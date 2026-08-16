/**
 * ICMS API — MULTI TABLE
 * Google Apps Script + Google Sheets
 *
 * Satu Web App URL ini bisa melayani BANYAK tabel (sheet), dibedakan
 * lewat parameter "sheet" yang dikirim dari frontend. Jika parameter
 * "sheet" tidak dikirim, default-nya adalah "rencana_kelas" (backward
 * compatible dengan index.html yang sudah ada).
 *
 * CARA MENAMBAH TABEL BARU:
 * 1. Tambahkan satu entri baru di object SCHEMAS di bawah ini.
 * 2. Jalankan fungsi setupAllSheets() sekali (Run > setupAllSheets).
 * 3. Di frontend, kirim parameter sheet:"nama_tabel_baru" pada apiGet/apiPost.
 * TIDAK PERLU deploy ulang Web App, dan TIDAK PERLU mengubah kode di bawah.
 */

const CONFIG = {
  SPREADSHEET_ID: "1pWrUcwejLGMmff0oSywZZcAPy-3pDvqr3czenc0dzfA",
  API_KEY: "2026-BNU-KARTUKONTROL",
  DEFAULT_SHEET: "rencana_kelas"
};

/**
 * DEFINISI SKEMA PER TABEL.
 * headers        : urutan kolom persis seperti di sheet (kolom pertama HARUS "id").
 * dateFields     : kolom yang harus diperlakukan sebagai tanggal.
 * numberFields   : kolom yang harus diperlakukan sebagai angka.
 * uniqueField    : (opsional) kolom yang tidak boleh duplikat, mis. "id_class".
 * importAliases  : (opsional) alias nama kolom asli (Excel) -> field key.
 */
const SCHEMAS = {

  rencana_kelas: {
    headers: [
      "id","fileName","id_class","nama_pelatihan","tgl_awal","tgl_akhir","lokasi",
      "npp_lc","pic_lc","pic_kelas","nama","total_input_pegawai","jenis_pelatihan",
      "metode","academy","unit","jumlah_hari","jumlah_peserta_didaftarkan",
      "jumlah_peserta_absen","biaya_pelatihan","status_kelas","tahun","created_at","updated_at"
    ],
    dateFields: ["tgl_awal","tgl_akhir","tahun"],
    numberFields: ["total_input_pegawai","jumlah_hari","jumlah_peserta_didaftarkan","jumlah_peserta_absen","biaya_pelatihan"],
    uniqueField: "id_class",
    importAliases: {
      fileName:["fileName","FILE_NAME","FILE NAME"],
      id_class:["id_class","IDCLASS","ID CLASS"],
      nama_pelatihan:["nama_pelatihan","NAMAPELATIHAN","NAMA PELATIHAN"],
      tgl_awal:["tgl_awal","TGLAWAL","TGL AWAL"],
      tgl_akhir:["tgl_akhir","TGLAKHIR","TGL AKHIR"],
      lokasi:["lokasi","LOKASI"],
      npp_lc:["npp_lc","NPPLC","NPP LC"],
      pic_lc:["pic_lc","PICLC","PIC LC"],
      pic_kelas:["pic_kelas","PICKELAS","PIC KELAS"],
      nama:["nama","NAMA"],
      total_input_pegawai:["total_input_pegawai","TOTALINPUTPEGAWAI","TOTAL INPUT PEGAWAI"],
      jenis_pelatihan:["jenis_pelatihan","JENISPELATIHAN","JENIS PELATIHAN"],
      metode:["metode","METODE"],
      academy:["academy","ACADEMY"],
      unit:["unit","UNIT"],
      jumlah_hari:["jumlah_hari","JUMLAHHARI","JUMLAH HARI"],
      jumlah_peserta_didaftarkan:["jumlah_peserta_didaftarkan","JUMLAHPESERTADIDAFTARKAN","JUMLAH PESERTA DIDAFTARKAN"],
      jumlah_peserta_absen:["jumlah_peserta_absen","JUMLAHPESERTAABSEN","JUMLAH PESERTA ABSEN"],
      biaya_pelatihan:["biaya_pelatihan","BIAYAPELATIHAN","BIAYA PELATIHAN"],
      status_kelas:["status_kelas","STATUSKELAS","STATUS KELAS"],
      tahun:["tahun","TAHUN"]
    }
  },

  learning_plan_realisasi: {
    headers: [
      "id","uic","nama_program","learning_goals",
      "awal_periode","akhir_periode","rev_awal_periode","rev_akhir_periode",
      "metode_pembelajaran","bulan_mulai","bulan_akhir","durasi_bulan",
      "status_progress_realisasi","nilai_per_bulan","total_estimasi_biaya",
      "revisi_total_estimasi_biaya","status_pelaksanaan_program","keterangan_perubahan",
      "bulan_1","bulan_2","bulan_3","bulan_4","bulan_5","bulan_6",
      "bulan_7","bulan_8","bulan_9","bulan_10","bulan_11","bulan_12",
      "cek",
      "akr_jan","akr_feb","akr_mar","akr_apr","akr_mei","akr_jun",
      "akr_jul","akr_agu","akr_sep","akr_okt","akr_nov","akr_des",
      "created_at","updated_at"
    ],
    dateFields: ["awal_periode","akhir_periode","rev_awal_periode","rev_akhir_periode"],
    numberFields: [
      "durasi_bulan","nilai_per_bulan","total_estimasi_biaya","revisi_total_estimasi_biaya",
      "bulan_1","bulan_2","bulan_3","bulan_4","bulan_5","bulan_6",
      "bulan_7","bulan_8","bulan_9","bulan_10","bulan_11","bulan_12",
      "akr_jan","akr_feb","akr_mar","akr_apr","akr_mei","akr_jun",
      "akr_jul","akr_agu","akr_sep","akr_okt","akr_nov","akr_des"
    ],
    uniqueField: "",
    importAliases: {
      uic:["UIC"],
      nama_program:["Nama Program"],
      learning_goals:["Learning Goals"],
      awal_periode:["Awal Periode Pelaksanaan"],
      akhir_periode:["Akhir Periode Pelaksanaan"],
      rev_awal_periode:["Rev. Awal Periode Pelaksanaan"],
      rev_akhir_periode:["Rev. Akhir Periode Pelaksanaan"],
      metode_pembelajaran:["Konfirmasi Metode Pembelajaran"],
      bulan_mulai:["Bulan Mulai"],
      bulan_akhir:["Bulan Akhir"],
      durasi_bulan:["Durasi Bulan"],
      status_progress_realisasi:["Status Progress Realisasi"],
      nilai_per_bulan:["Nilai per Bulan"],
      total_estimasi_biaya:["Total Estimasi Biaya"],
      revisi_total_estimasi_biaya:["Revisi Total Estimasi Biaya"],
      status_pelaksanaan_program:["Status Pelaksanaan Program"],
      keterangan_perubahan:["Keterangan Perubahan dan Strategi Close the Gap"],
      bulan_1:["Bulan 1"], bulan_2:["Bulan 2"], bulan_3:["Bulan 3"], bulan_4:["Bulan 4"],
      bulan_5:["Bulan 5"], bulan_6:["Bulan 6"], bulan_7:["Bulan 7"], bulan_8:["Bulan 8"],
      bulan_9:["Bulan 9"], bulan_10:["Bulan 10"], bulan_11:["Bulan 11"], bulan_12:["Bulan 12"],
      cek:["CEK"],
      akr_jan:["Jan (10% x Accr Jan)"],
      akr_feb:["Feb (20% x Accr Jan)"],
      akr_mar:["Mar (10% x Accr Jan + Accr Feb)"],
      akr_apr:["Apr (Accr Mar)"],
      akr_mei:["Mei (Accr Apr)"],
      akr_jun:["Jun (Accr Mei)"],
      akr_jul:["Jul (Accr Jun)"],
      akr_agu:["Agu (Accr Jul)"],
      akr_sep:["Sep (Accr Agu)"],
      akr_okt:["Okt (Accr Sep)"],
      akr_nov:["Nov (Accr Okt)"],
      akr_des:["Des (Accr Nov + Accr Des)"]
    }
  },

  rekap_program: {
    headers: [
      "id","uic","program","subprogram","pengelola","tgl_awal","tgl_akhir",
      "metode","jumlah_hari","learner","biaya_pelatihan","created_at","updated_at"
    ],
    dateFields: ["tgl_awal","tgl_akhir"],
    numberFields: ["jumlah_hari","learner","biaya_pelatihan"],
    uniqueField: "",
    importAliases: {
      uic:["UIC"],
      program:["Program"],
      subprogram:["Subprogram"],
      pengelola:["Pengelola"],
      tgl_awal:["Tgl Awal"],
      tgl_akhir:["Tgl Akhir"],
      metode:["Metode"],
      jumlah_hari:["Jumlah Hari"],
      learner:["Learner"],
      biaya_pelatihan:["Biaya Pelatihan"]
    }
  }

};

function setupAllSheets(){
  const results=[];
  Object.keys(SCHEMAS).forEach(name=>results.push(setupSheet_(name)));
  return results.join(" | ");
}

function setupSheet_(sheetName){
  const schema=getSchema_(sheetName);
  const ss=SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sh=ss.getSheetByName(sheetName);
  if(!sh) sh=ss.insertSheet(sheetName);
  sh.clear();
  sh.getRange(1,1,1,schema.headers.length).setValues([schema.headers]);
  sh.setFrozenRows(1);
  sh.getRange(1,1,1,schema.headers.length).setFontWeight("bold");
  return "Sheet siap: "+sheetName;
}

function getSchema_(sheetName){
  const schema=SCHEMAS[sheetName];
  if(!schema) throw new Error("Tabel/sheet tidak dikenal: "+sheetName);
  return schema;
}

function doGet(e){
  try{
    auth_(e.parameter.key);
    const action = e.parameter.action || "list";
    const sheetName = e.parameter.sheet || CONFIG.DEFAULT_SHEET;
    if(action==="list") return json_({ok:true,data:list_(sheetName)});
    if(action==="setup") return json_({ok:true,msg:setupSheet_(sheetName)});
    if(action==="setupAll") return json_({ok:true,msg:setupAllSheets()});
    if(action==="tables") return json_({ok:true,data:Object.keys(SCHEMAS)});
    return json_({ok:false,msg:"Action GET tidak dikenal."});
  }catch(err){return json_({ok:false,msg:String(err.message||err)})}
}

function doPost(e){
  try{
    const p = e.parameter || {};
    auth_(p.key);
    const action = p.action || "";
    const sheetName = p.sheet || CONFIG.DEFAULT_SHEET;
    if(action==="save") return json_(save_(sheetName,p));
    if(action==="delete") return json_(delete_(sheetName,p.id));
    if(action==="deleteAll") return json_(deleteAll_(sheetName));
    if(action==="import") return json_(import_(sheetName,p.rows));
    return json_({ok:false,msg:"Action POST tidak dikenal."});
  }catch(err){return json_({ok:false,msg:String(err.message||err)})}
}

function auth_(key){
  if(CONFIG.API_KEY && key !== CONFIG.API_KEY) throw new Error("API key tidak valid.");
}

function sheet_(sheetName){
  getSchema_(sheetName); // validasi tabel dikenal
  const ss=SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  let sh=ss.getSheetByName(sheetName);
  if(!sh){setupSheet_(sheetName);sh=ss.getSheetByName(sheetName);}
  return sh;
}

function list_(sheetName){
  const sh=sheet_(sheetName);
  const values=sh.getDataRange().getValues();
  if(values.length<2)return [];
  const headers=values[0];
  return values.slice(1).filter(r=>r.some(v=>v!=="")).map(r=>{
    const o={};
    headers.forEach((h,i)=>o[h]=normalize_(r[i]));
    return o;
  });
}

function save_(sheetName,p){
  const schema=getSchema_(sheetName);
  const sh=sheet_(sheetName);
  const now=new Date();
  const id=p.id || Utilities.getUuid();
  const rows=sh.getDataRange().getValues();
  const headers=rows[0];
  let rowIndex=-1;
  for(let i=1;i<rows.length;i++){
    if(String(rows[i][0])===String(id)){rowIndex=i+1;break;}
  }

  // Validasi kolom unik (jika didefinisikan di schema)
  if(schema.uniqueField){
    const uniqueColIdx=headers.indexOf(schema.uniqueField);
    const uniqueVal=String(p[schema.uniqueField]||"").trim();
    if(uniqueVal && uniqueColIdx>-1){
      for(let i=1;i<rows.length;i++){
        if(rowIndex===i+1)continue;
        if(String(rows[i][uniqueColIdx]).trim()===uniqueVal){
          throw new Error(schema.uniqueField+" sudah digunakan: "+uniqueVal);
        }
      }
    }
  }

  const createdAtIdx=headers.indexOf("created_at");
  const old = rowIndex>0 ? rows[rowIndex-1] : [];
  const createdAt = rowIndex>0 && createdAtIdx>-1 ? old[createdAtIdx] : now;

  const obj={id:id};
  headers.forEach(h=>{
    if(h==="id")return;
    if(h==="created_at"){obj[h]=createdAt;return;}
    if(h==="updated_at"){obj[h]=now;return;}
    if(schema.dateFields.includes(h)){obj[h]=parseDate_(p[h]);return;}
    if(schema.numberFields.includes(h)){obj[h]=num_(p[h]);return;}
    obj[h]=p[h]||"";
  });

  const row=headers.map(h=>obj[h]??"");
  if(rowIndex>0)sh.getRange(rowIndex,1,1,headers.length).setValues([row]);
  else sh.appendRow(row);
  return {ok:true,msg:rowIndex>0?"Data berhasil diperbarui.":"Data berhasil ditambahkan.",id:id};
}

function delete_(sheetName,id){
  if(!id)throw new Error("ID tidak diberikan.");
  const sh=sheet_(sheetName), rows=sh.getDataRange().getValues();
  for(let i=1;i<rows.length;i++){
    if(String(rows[i][0])===String(id)){sh.deleteRow(i+1);return {ok:true,msg:"Data berhasil dihapus."};}
  }
  throw new Error("Data tidak ditemukan.");
}

function deleteAll_(sheetName){
  const sh=sheet_(sheetName);
  if(sh.getLastRow()>1)sh.getRange(2,1,sh.getLastRow()-1,sh.getLastColumn()).clearContent();
  return {ok:true,msg:"Semua data berhasil dihapus."};
}

function import_(sheetName,raw){
  const schema=getSchema_(sheetName);
  if(!raw)throw new Error("Data import kosong.");
  const rows=JSON.parse(raw);
  if(!Array.isArray(rows)||!rows.length)throw new Error("Data import tidak valid.");
  let ok=0,fail=[];
  rows.forEach((r,i)=>{
    try{
      const p=mapImport_(schema,r);
      save_(sheetName,p);ok++;
    }catch(err){fail.push("Baris "+(i+2)+": "+err.message)}
  });
  return {ok:true,msg:ok+" data berhasil diimpor."+ (fail.length?" "+fail.length+" gagal.":""),inserted_count:ok,failed_count:fail.length,errors:fail};
}

function mapImport_(schema,r){
  const aliases=schema.importAliases||{};
  const get=(names)=>{
    for(const n of names){
      const key=Object.keys(r).find(k=>normalizeHeader_(k)===normalizeHeader_(n));
      if(key!==undefined)return r[key];
    }
    return "";
  };
  const p={id:""};
  schema.headers.forEach(h=>{
    if(h==="id"||h==="created_at"||h==="updated_at")return;
    const names=aliases[h] || [h];
    p[h]=get(names);
  });
  return p;
}

function normalizeHeader_(s){return String(s||"").toUpperCase().replace(/[^A-Z0-9]/g,"");}
function num_(v){if(v===""||v===null||v===undefined)return 0;const n=Number(String(v).replace(/[^0-9.-]/g,""));return isNaN(n)?0:n}
function parseDate_(v){
  if(!v)return "";
  if(Object.prototype.toString.call(v)==="[object Date]")return v;
  const s=String(v).trim();
  const d=new Date(s);
  if(!isNaN(d.getTime()))return d;
  const m=s.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})/);
  if(m)return new Date(Number(m[3]),Number(m[2])-1,Number(m[1]));
  return s;
}
function normalize_(v){
  if(v===null||v===undefined)return "";
  if(v instanceof Date)return Utilities.formatDate(v,Session.getScriptTimeZone(),"yyyy-MM-dd HH:mm:ss");
  return v;
}
function json_(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
