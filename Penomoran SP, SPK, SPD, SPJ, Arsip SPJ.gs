/**
 * ============================================================
 * SISTEM PENOMORAN OTOMATIS (SP, SPK, SPJ, SPD) + ARSIP SPJ
 * ============================================================
 * Update perbaikan (tambahan SPJ):
 * 1) SPJ: "Nomor SPTB" dibuat otomatis (tidak input manual).
 *    Format: [no urut]/SPTB/[kode otorisasi]/[tahun]
 *    No urut reset per tahun (berdasarkan max no urut SPTB pada tahun tsb).
 *    Diisi saat submit (dan dipastikan terisi saat ada perubahan status bila masih kosong).
 *
 * 2) SPJ: Email (submit / progress / dikembalikan / final) memuat semua kolom penting
 *    + tambahan "No. SPM" tepat di bawah "Status SPM" pada bagian Status Proses.
 *
 * 3) Format tanggal di email jadi Indonesia (dd/MM/yyyy) untuk semua kolom bertipe Date.
 *
 * 4) Sheet baru: "Arsip SPJ"
 *    onFormSubmit (sheet Arsip SPJ): jika "Nomor SPTB" pada Arsip SPJ sama dengan
 *    "Nomor SPTB" pada SPJ, maka kolom checkbox "Arsip" pada SPJ dicentang TRUE.
 *
 * TRIGGER:
 * 1) onFormSubmit  : From spreadsheet | On form submit
 * 2) onEdit        : From spreadsheet | On edit (khusus SPJ)
 * ============================================================
 */

const KODE_OTORITAS_DEFAULT = "Bbws2.a1";

function onFormSubmit(e) {
  if (!e || !e.namedValues || !e.range) return;

  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  if (sheetName === "SP")        return handleSP_(e);
  if (sheetName === "SPK")       return handleSPK_(e);
  if (sheetName === "SPJ")       return handleSPJ_(e);
  if (sheetName === "SPD")       return handleSPD_(e);
  if (sheetName === "Arsip SPJ") return handleArsipSPJ_(e);
}

function onEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  if (sheet.getName() !== "SPJ") return;

  const row = e.range.getRow();
  if (row < 2) return;

  const headerInfo = getHeaderInfo_(sheet, "SPJ");
  if (row <= headerInfo.row) return;

  const editedCol = e.range.getColumn();

  // 1) Override Nomor Bukti Kuitansi jika "Jumlah Kuitansi..." diedit
  const COL_JML_KUITANSI = mustCol_(headerInfo, "Jumlah Kuitansi Yang dibutuhkan", "SPJ");
  const COL_NO_BUKTI     = mustCol_(headerInfo, "Nomor Bukti Kuitansi", "SPJ");

  if (editedCol === COL_JML_KUITANSI) {
    fillNomorBuktiKuitansiForRow_(sheet, row, headerInfo, COL_JML_KUITANSI, COL_NO_BUKTI, true);
    // pastikan SPTB juga ada
    ensureNomorSPTB_(sheet, row, headerInfo, false);
    return;
  }

  // 2) Workflow status bertahap (kirim email sesuai validasi)
  const COL_STATUS_PPK  = mustCol_(headerInfo, "Status PPK", "SPJ");
  const COL_STATUS_SEK  = mustCol_(headerInfo, "Status Sekretariat", "SPJ");
  const COL_STATUS_BEN  = mustCol_(headerInfo, "Status Bendahara", "SPJ");
  const COL_STATUS_SPM  = mustCol_(headerInfo, "Status SPM", "SPJ");

  if (![COL_STATUS_PPK, COL_STATUS_SEK, COL_STATUS_BEN, COL_STATUS_SPM].includes(editedCol)) return;

  processSPJStatusChange_(sheet, row, headerInfo, editedCol);
}

/* ============================================================
 * HANDLER: SP
 * ============================================================ */
function handleSP_(e) {
  const TARGET_SHEET_NAME = "SP";
  const sheet = e.range.getSheet();
  if (sheet.getName() !== TARGET_SHEET_NAME) return;

  const row = e.range.getRow();
  const headerInfo = getHeaderInfo_(sheet, "SP");
  if (row <= headerInfo.row) return;

  const getVal = (headerName) => (e.namedValues[headerName] ? e.namedValues[headerName][0] : "");
  const safeGet = (name) => (getVal(name) ? getVal(name) : "-");

  const COL_TANGGAL  = mustCol_(headerInfo, "Tanggal Surat Pesanan", TARGET_SHEET_NAME);
  const COL_NO_URUT  = mustCol_(headerInfo, "Nomor Urut SP", TARGET_SHEET_NAME);
  const COL_NOMOR_SP = mustCol_(headerInfo, "Nomor Surat Pesanan", TARGET_SHEET_NAME);

  const emailPemohon = getVal("Email pemohon");
  const tanggalSPRaw = getVal("Tanggal Surat Pesanan");
  if (!emailPemohon || !tanggalSPRaw) return;

  const tanggalObj = parseTanggalID_(tanggalSPRaw);
  if (!tanggalObj) throw new Error(`Format "Tanggal Surat Pesanan" tidak valid: ${tanggalSPRaw}`);

  const tahun = tanggalObj.getFullYear();
  const nomorUrut = nextUrutPerTahun_(sheet, row, headerInfo, COL_TANGGAL, COL_NO_URUT, tahun);
  const nomorUrutFormatted = String(nomorUrut).padStart(2, "0");

  const nomorSP = `${nomorUrutFormatted}/SP/${KODE_OTORITAS_DEFAULT}/${tahun}`;

  sheet.getRange(row, COL_NO_URUT).setValue(nomorUrut);
  sheet.getRange(row, COL_NOMOR_SP).setValue(nomorSP);

  const subject = `Nomor Surat Pesanan (SP): ${nomorSP}`;

  const htmlBody = buildEmailSP_(nomorSP, tanggalObj, {
    "Nama Rekanan/Pihak Ketiga/Penyedia Jasa": safeGet("Pihak Ketiga/Penyedia Jasa") || safeGet("Nama Rekanan/Pihak Ketiga/Penyedia Jasa"),
    "Uraian Paket Pekerjaan": safeGet("Uraian Paket Pekerjaan"),
    "Mekanisme Pembayaran": safeGet("Mekanisme Pembayaran"),
    "Nilai SPJ": getVal("Nilai SPJ"),
    "Komponen/Subkomponen": safeGet("Komponen/Subkomponen"),
    "Akun": safeGet("Akun"),
    "Petugas Input Data": safeGet("Petugas Input Data"),
    "Keterangan": safeGet("Keterangan"),
  });

  const plainBody =
`Yth. Pemohon,

Pengajuan penomoran Surat Pesanan (SP) Anda telah berhasil diproses.

Nomor Surat Pesanan (SP):
${nomorSP}

Tanggal SP: ${formatTanggalID_(tanggalObj)}

Sistem Penomoran Otomatis BBWS Mesuji Sekampung`;

  MailApp.sendEmail({ to: emailPemohon, subject, htmlBody, body: plainBody });
}

/* ============================================================
 * HANDLER: SPK
 * ============================================================ */
function handleSPK_(e) {
  const TARGET_SHEET_NAME = "SPK";
  const sheet = e.range.getSheet();
  if (sheet.getName() !== TARGET_SHEET_NAME) return;

  const row = e.range.getRow();
  const headerInfo = getHeaderInfo_(sheet, "SPK");
  if (row <= headerInfo.row) return;

  const getVal = (headerName) => (e.namedValues[headerName] ? e.namedValues[headerName][0] : "");
  const safeGet = (name) => (getVal(name) ? getVal(name) : "-");

  const COL_TANGGAL   = mustCol_(headerInfo, "Tanggal SPK", TARGET_SHEET_NAME);
  const COL_NO_URUT   = mustCol_(headerInfo, "Nomor Urut SPK", TARGET_SHEET_NAME);
  const COL_NOMOR_SPK = mustCol_(headerInfo, "Nomor SPK", TARGET_SHEET_NAME);

  const emailPemohon = getVal("Email pemohon");
  const tanggalSPKRaw = getVal("Tanggal SPK");
  if (!emailPemohon || !tanggalSPKRaw) return;

  const tanggalObj = parseTanggalID_(tanggalSPKRaw);
  if (!tanggalObj) throw new Error(`Format "Tanggal SPK" tidak valid: ${tanggalSPKRaw}`);

  const tahun = tanggalObj.getFullYear();

  const nomorUrut = nextUrutPerTahun_(sheet, row, headerInfo, COL_TANGGAL, COL_NO_URUT, tahun);
  const nomorUrutFormatted = String(nomorUrut).padStart(2, "0");
  const nomorSPK = `${nomorUrutFormatted}/SPK/${KODE_OTORITAS_DEFAULT}/${tahun}`;

  sheet.getRange(row, COL_NO_URUT).setValue(nomorUrut);
  sheet.getRange(row, COL_NOMOR_SPK).setValue(nomorSPK);

  const subject = `Nomor SPK: ${nomorSPK}`;

  const htmlBody = buildEmailSPK_(nomorSPK, tanggalObj, {
    "Nama Rekanan/Pihak Ketiga/Penyedia Jasa": safeGet("Nama Rekanan/Pihak Ketiga/Penyedia Jasa"),
    "Uraian SPK": safeGet("Uraian SPK"),
    "Mekanisme Pembayaran": safeGet("Mekanisme Pembayaran"),
    "Nilai Kontrak": getVal("Nilai Kontrak"),
    "Komponen/Sub Komponen": safeGet("Komponen/Sub Komponen"),
    "Akun": safeGet("Akun"),
    "Keterangan": safeGet("Keterangan"),
    "Petugas Input Data": safeGet("Petugas Input Data"),
  });

  const plainBody =
`Yth. Pemohon,

Pengajuan penomoran SPK Anda telah berhasil diproses.

Nomor SPK:
${nomorSPK}

Tanggal SPK: ${formatTanggalID_(tanggalObj)}

Sistem Penomoran Otomatis BBWS Mesuji Sekampung`;

  MailApp.sendEmail({ to: emailPemohon, subject, htmlBody, body: plainBody });
}

/* ============================================================
 * HANDLER: SPJ (SUBMIT)
 * ============================================================ */
function handleSPJ_(e) {
  const TARGET_SHEET_NAME = "SPJ";
  const sheet = e.range.getSheet();
  if (sheet.getName() !== TARGET_SHEET_NAME) return;

  const row = e.range.getRow();
  const headerInfo = getHeaderInfo_(sheet, "SPJ");
  if (row <= headerInfo.row) return;

  const getVal = (headerName) => (e.namedValues[headerName] ? e.namedValues[headerName][0] : "");

  const COL_EMAIL         = mustCol_(headerInfo, "Email pemohon", "SPJ");
  const COL_JML_KUITANSI  = mustCol_(headerInfo, "Jumlah Kuitansi Yang dibutuhkan", "SPJ");
  const COL_NO_BUKTI      = mustCol_(headerInfo, "Nomor Bukti Kuitansi", "SPJ");
  const COL_STATUS_PPK    = mustCol_(headerInfo, "Status PPK", "SPJ");

  const emailPemohon = String(sheet.getRange(row, COL_EMAIL).getValue() || getVal("Email pemohon") || "").trim();
  if (!emailPemohon) return;

  // Default status PPK
  const curPPK = String(sheet.getRange(row, COL_STATUS_PPK).getValue() || "").trim();
  if (!curPPK) sheet.getRange(row, COL_STATUS_PPK).setValue("Menunggu");

  // Nomor Bukti Kuitansi
  fillNomorBuktiKuitansiForRow_(sheet, row, headerInfo, COL_JML_KUITANSI, COL_NO_BUKTI, false);

  // Nomor SPTB otomatis
  ensureNomorSPTB_(sheet, row, headerInfo, false);

  const subject = "Pengajuan SPJ diterima: menunggu persetujuan PPK";

  const ctx = buildSPJContext_(sheet, row, headerInfo);
  const htmlBody = buildEmailSPJSubmitFull_(ctx);

  const plainBody =
`Yth. Pemohon,

Pengajuan SPJ Anda sudah diterima.
Status saat ini: MENUNGGU persetujuan PPK.

Nomor Bukti Kuitansi: ${ctx["Nomor Bukti Kuitansi"] || "-"}
Nomor SPTB: ${ctx["Nomor SPTB"] || "-"}

Sistem Penomoran Otomatis BBWS Mesuji Sekampung`;

  MailApp.sendEmail({ to: emailPemohon, subject, htmlBody, body: plainBody });
}

/* ============================================================
 * HANDLER: SPD (UPDATE)
 * ============================================================ */
/* ============================================================
 * HANDLER: SPD (UPDATE: Checkbox Pelaksana & Auto Count)
 * ============================================================ */
function handleSPD_(e) {
  const TARGET_SHEET_NAME = "SPD";
  const sheet = e.range.getSheet();
  if (sheet.getName() !== TARGET_SHEET_NAME) return;

  const row = e.range.getRow();
  const headerInfo = getHeaderInfo_(sheet, "SPD");
  if (row <= headerInfo.row) return;

  const getVal = (headerName) => (e.namedValues[headerName] ? e.namedValues[headerName][0] : "");
  const safeGet = (name) => (getVal(name) ? getVal(name) : "-");

  const COL_NO_URUT = mustCol_(headerInfo, "Nomor Urut SPD", TARGET_SHEET_NAME);
  const COL_NOMOR   = mustCol_(headerInfo, "Nomor SPD", TARGET_SHEET_NAME);
  const COL_PELAKSANA = mustCol_(headerInfo, "Nama Pelaksana Perjalanan Dinas", TARGET_SHEET_NAME);
  
  const emailPemohon = String(getVal("Email pemohon") || "").trim();
  if (!emailPemohon) return;

  const maksudIsi = getVal("Maksud / Keperluan Perjalanan Dinas") || "-";

  // --- LOGIKA BARU: HITUNG ORANG DARI CHECKBOX ---
  const rawPelaksana = String(getVal("Nama Pelaksana Perjalanan Dinas") || "");
  
  // 1. Split berdasarkan titik koma (;) sesuai instruksi
  // 2. Trim untuk hapus spasi/koma sisa format Google Form (misal "Nama A;, Nama B;")
  // 3. Filter yang kosong
  const listPelaksana = rawPelaksana.split(";")
    .map(n => n.replace(/^,|,$/g, "").trim()) // Hapus koma sisa checkbox form
    .filter(n => n !== "");

  const totalOrang = listPelaksana.length;

  if (totalOrang === 0) {
    // Safety jika kosong (jarang terjadi di required field)
    return;
  }

  // Update Tampilan Kolom Pelaksana (biar rapi ke bawah)
  // Format: Nama A; (Enter) Nama B;
  const pelaksanaFormatted = listPelaksana.join(";\n") + (listPelaksana.length > 0 ? ";" : ""); 
  sheet.getRange(row, COL_PELAKSANA).setValue(pelaksanaFormatted);

  // Update Kolom "Jumlah Orang" jika ada di Spreadsheet
  if (hasHeader_(headerInfo, "Jumlah Orang")) {
    const COL_JML_ORANG = mustCol_(headerInfo, "Jumlah Orang", TARGET_SHEET_NAME);
    sheet.getRange(row, COL_JML_ORANG).setValue(totalOrang);
  }
  // ------------------------------------------------

  let tahun = new Date().getFullYear();
  if (hasHeader_(headerInfo, "Tanggal SPT")) {
    const t = parseTanggalID_(getVal("Tanggal SPT"));
    if (t) tahun = t.getFullYear();
  }

  // Generate Nomor
  const startUrut = nextUrutSPDStart_(sheet, row, headerInfo, COL_NO_URUT, tahun);
  const endUrut = startUrut + totalOrang - 1;

  const nomorList = [];
  for (let n = startUrut; n <= endUrut; n++) {
    const noFmt = String(n).padStart(2, "0");
    nomorList.push(`${noFmt}/SPD/${KODE_OTORITAS_DEFAULT}/${tahun}`);
  }

  sheet.getRange(row, COL_NOMOR).setValue(nomorList.join("\n"));

  const startFmt = String(startUrut).padStart(2, "0");
  const endFmt   = String(endUrut).padStart(2, "0");
  const rangeText = (startUrut === endUrut) ? startFmt : `${startFmt} s/d ${endFmt}`;
  sheet.getRange(row, COL_NO_URUT).setValue(rangeText);

  // Persiapan Data untuk Email
  // Nama pertama dianggap Pelaksana Utama, sisanya Pengikut
  const namaUtama = listPelaksana[0]; 
  const sisaPengikut = listPelaksana.slice(1); // Array sisa nama
  const jumlahPengikut = sisaPengikut.length;
  const namaPengikutClean = sisaPengikut.length > 0 ? sisaPengikut.join(";\n") : "-";

  const subject = `Nomor SPD: ${nomorList[0]}`;
  const htmlBody = buildEmailSPD_(nomorList, {
    namaPelaksana: namaUtama,        // Hanya nama pertama
    pengikut: String(jumlahPengikut),// Sisa orang
    namaPengikut: namaPengikutClean, // List nama sisa
    maksud: maksudIsi,
    jenisTujuan: safeGet("Jenis Tujuan"),
    tanggalSPT: safeGet("Tanggal SPT"),
    tanggalBerangkat: safeGet("Tanggal Berangkat"),
    tanggalKembali: safeGet("Tanggal Kembali"),
    komponen: safeGet("Komponen/Subkomponen"),
    akun: safeGet("Akun"),
    keterangan: safeGet("Keterangan"),
    petugas: safeGet("Petugas Input Data"),
  });

  const plainBody = `Yth. Pemohon,\n\nPengajuan SPD Anda telah diproses.\n\nNomor SPD:\n${nomorList.join("\n")}\n\nSistem Penomoran Otomatis`;
  MailApp.sendEmail({ to: emailPemohon, subject, htmlBody, body: plainBody });
}

/* ============================================================
 * HANDLER: Arsip SPJ (SUBMIT)
 * ============================================================ */
function handleArsipSPJ_(e) {
  const TARGET_SHEET_NAME = "Arsip SPJ";
  const sheet = e.range.getSheet();
  if (sheet.getName() !== TARGET_SHEET_NAME) return;

  const ss = sheet.getParent();
  const row = e.range.getRow();

  // Header Arsip SPJ
  const hArsip = getHeaderInfo_(sheet, "Arsip SPJ");
  if (row <= hArsip.row) return;

  const colSPTB_arsip = mustCol_(hArsip, "Nomor SPTB", "Arsip SPJ");
  const sptb = String(sheet.getRange(row, colSPTB_arsip).getValue() || "").trim();
  if (!sptb) return;

  // Cari di sheet SPJ dan centang kolom "Arsip" jika Nomor SPTB sama
  const sheetSPJ = ss.getSheetByName("SPJ");
  if (!sheetSPJ) throw new Error('Sheet "SPJ" tidak ditemukan.');

  const hSPJ = getHeaderInfo_(sheetSPJ, "SPJ");
  const colSPTB_spj = mustCol_(hSPJ, "Nomor SPTB", "SPJ");
  const colArsip_spj = mustCol_(hSPJ, "Arsip", "SPJ");

  const lastRowSPJ = sheetSPJ.getLastRow();
  if (lastRowSPJ <= hSPJ.row) return;

  // Scan SPJ
  for (let r = hSPJ.row + 1; r <= lastRowSPJ; r++) {
    const v = String(sheetSPJ.getRange(r, colSPTB_spj).getValue() || "").trim();
    if (v && v === sptb) {
      sheetSPJ.getRange(r, colArsip_spj).setValue(true);
      break;
    }
  }
}

/* ============================================================
 * SPJ: AUTO NOMOR BUKTI KUITANSI
 * ============================================================ */
function fillNomorBuktiKuitansiForRow_(sheet, row, headerInfo, colJumlah, colNomorBukti, forceOverwrite) {
  const raw = sheet.getRange(row, colJumlah).getValue();
  const jumlah = parseJumlahKuitansi_(raw);

  if (!jumlah || jumlah <= 0) return "";

  const existing = String(sheet.getRange(row, colNomorBukti).getValue() || "").trim();
  if (existing && !forceOverwrite) return existing;

  const maxUsed = findMaxKuitansiUsed_(sheet, headerInfo, colNomorBukti);

  const start = maxUsed + 1;
  const end = start + jumlah - 1;

  const val = (jumlah === 1) ? String(start) : `${start} s/d ${end}`;
  sheet.getRange(row, colNomorBukti).setValue(val);
  return val;
}

function parseJumlahKuitansi_(value) {
  const s = String(value ?? "").trim();
  if (!s) return 0;
  const m = s.match(/\d+/);
  if (!m) return 0;
  return Number(m[0]) || 0;
}

function findMaxKuitansiUsed_(sheet, headerInfo, colNomorBukti) {
  const lastRow = sheet.getLastRow();
  let maxN = 0;

  for (let r = headerInfo.row + 1; r <= lastRow; r++) {
    const val = String(sheet.getRange(r, colNomorBukti).getValue() || "").trim();
    if (!val) continue;

    const nums = val.match(/\d+/g);
    if (!nums || nums.length === 0) continue;

    for (let i = 0; i < nums.length; i++) {
      const n = Number(nums[i]) || 0;
      if (n > maxN) maxN = n;
    }
  }

  return maxN;
}

/* ============================================================
 * SPJ: AUTO NOMOR SPTB (FORMAT BARU)
 * ============================================================ */
function ensureNomorSPTB_(sheet, row, headerInfo, forceOverwrite) {
  if (!hasHeader_(headerInfo, "Nomor SPTB")) return "";

  const colSPTB = mustCol_(headerInfo, "Nomor SPTB", "SPJ");
  const existing = String(sheet.getRange(row, colSPTB).getValue() || "").trim();
  if (existing && !forceOverwrite) return existing;

  // Tahun dari Timestamp jika ada
  let tahun = new Date().getFullYear();
  if (hasHeader_(headerInfo, "Timestamp")) {
    const colTS = mustCol_(headerInfo, "Timestamp", "SPJ");
    const tsVal = sheet.getRange(row, colTS).getValue();
    const ts = parseTanggalID_(tsVal);
    if (ts) tahun = ts.getFullYear();
  }

  // Cari max urut SPTB untuk tahun tsb
  const maxUsed = findMaxSPTBUsedByYear_(sheet, headerInfo, colSPTB, tahun);

  const nextNo = maxUsed + 1;
  const noFmt = String(nextNo).padStart(2, "0");

  const nomorSPTB = `${noFmt}/SPTB/${KODE_OTORITAS_DEFAULT}/${tahun}`;
  sheet.getRange(row, colSPTB).setValue(nomorSPTB);

  return nomorSPTB;
}

function findMaxSPTBUsedByYear_(sheet, headerInfo, colSPTB, tahun) {
  const lastRow = sheet.getLastRow();
  let maxN = 0;

  for (let r = headerInfo.row + 1; r <= lastRow; r++) {
    const v = String(sheet.getRange(r, colSPTB).getValue() || "").trim();
    if (!v) continue;

    // NN/SPTB/KODE/TAHUN
    const m = v.match(/^(\d+)\s*\/\s*SPTB\s*\/\s*[^\/]+\s*\/\s*(\d{4})$/i);
    if (!m) continue;

    const n = Number(m[1]) || 0;
    const y = Number(m[2]) || 0;

    if (y === tahun && n > maxN) maxN = n;
  }

  return maxN;
}

/* ============================================================
 * SPJ: WORKFLOW STATUS BERTAHAP + EMAIL (EMAIL FULL)
 * ============================================================ */
function processSPJStatusChange_(sheet, row, headerInfo, editedCol) {
  const COL_EMAIL        = mustCol_(headerInfo, "Email pemohon", "SPJ");
  const COL_STATUS_PPK   = mustCol_(headerInfo, "Status PPK", "SPJ");
  const COL_STATUS_SEK   = mustCol_(headerInfo, "Status Sekretariat", "SPJ");
  const COL_STATUS_BEN   = mustCol_(headerInfo, "Status Bendahara", "SPJ");
  const COL_STATUS_SPM   = mustCol_(headerInfo, "Status SPM", "SPJ");

  const email = String(sheet.getRange(row, COL_EMAIL).getValue() || "").trim();
  if (!email) return;

  const statusPPK = normalizeStatus_(sheet.getRange(row, COL_STATUS_PPK).getValue());
  const statusSEK = normalizeStatus_(sheet.getRange(row, COL_STATUS_SEK).getValue());
  const statusBEN = normalizeStatus_(sheet.getRange(row, COL_STATUS_BEN).getValue());
  const statusSPM = normalizeStatus_(sheet.getRange(row, COL_STATUS_SPM).getValue());

  // Pastikan SPTB terisi
  ensureNomorSPTB_(sheet, row, headerInfo, false);

  let tahap = "";
  let statusNow = "";
  let valid = false;
  let waitingText = "";

  if (editedCol === COL_STATUS_PPK) {
    tahap = "PPK";
    statusNow = statusPPK;
    valid = true;
    waitingText = "Menunggu Sekretariat, Bendahara, dan SPM";
  } else if (editedCol === COL_STATUS_SEK) {
    tahap = "Sekretariat";
    statusNow = statusSEK;
    valid = (statusPPK === "DISETUJUI");
    waitingText = "Menunggu Bendahara dan SPM";
  } else if (editedCol === COL_STATUS_BEN) {
    tahap = "Bendahara";
    statusNow = statusBEN;
    valid = (statusPPK === "DISETUJUI" && statusSEK === "DISETUJUI");
    waitingText = "Menunggu SPM";
  } else if (editedCol === COL_STATUS_SPM) {
    tahap = "SPM";
    statusNow = statusSPM;
    valid = (statusPPK === "DISETUJUI" && statusSEK === "DISETUJUI" && statusBEN === "DISETUJUI");
    waitingText = "Selesai";
  }

  if (!valid) return;
  if (statusNow !== "DISETUJUI" && statusNow !== "DIKEMBALIKAN") return;

  const ctx = buildSPJContext_(sheet, row, headerInfo);
  ctx["Status PPK"] = statusPPK || ctx["Status PPK"];
  ctx["Status Sekretariat"] = statusSEK || ctx["Status Sekretariat"];
  ctx["Status Bendahara"] = statusBEN || ctx["Status Bendahara"];
  ctx["Status SPM"] = statusSPM || ctx["Status SPM"];

  if (statusNow === "DIKEMBALIKAN") {
    const subject = `SPJ dikembalikan oleh ${tahap}`;
    const htmlBody = buildEmailSPJReturnedFull_(tahap, ctx);
    const plainBody =
`Yth. Pemohon,

Pengajuan SPJ Anda DIKEMBALIKAN oleh ${tahap}.

Keterangan:
${ctx["Keterangan"] || "-"}

Sistem Penomoran Otomatis BBWS Mesuji Sekampung`;
    MailApp.sendEmail({ to: email, subject, htmlBody, body: plainBody });
    return;
  }

  if (tahap === "SPM") {
    const subject = "SPJ disetujui: selesai (SPM)";
    const htmlBody = buildEmailSPJFinalBySPMFull_(ctx);
    const plainBody =
`Yth. Pemohon,

Pengajuan SPJ Anda sudah DISETUJUI sampai tahap SPM (SELESAI).

Nomor Bukti Kuitansi: ${ctx["Nomor Bukti Kuitansi"] || "-"}
Nomor SPTB: ${ctx["Nomor SPTB"] || "-"}
Nomor SPBY: ${ctx["Nomor SPBY"] || "-"}

Sistem Penomoran Otomatis BBWS Mesuji Sekampung`;
    MailApp.sendEmail({ to: email, subject, htmlBody, body: plainBody });
    return;
  }

  const subject = `SPJ disetujui oleh ${tahap}: ${waitingText}`;
  const htmlBody = buildEmailSPJProgressFull_(tahap, waitingText, ctx);
  const plainBody =
`Yth. Pemohon,

Pengajuan SPJ Anda telah DISETUJUI oleh ${tahap}.
Status berikutnya: ${waitingText}

Sistem Penomoran Otomatis BBWS Mesuji Sekampung`;
  MailApp.sendEmail({ to: email, subject, htmlBody, body: plainBody });
}

function normalizeStatus_(v) {
  const s = String(v || "").trim().toUpperCase();
  if (!s) return "";
  if (s === "SETUJU") return "DISETUJUI";
  if (s === "DISETUJUI") return "DISETUJUI";
  if (s === "DIKEMBALIKAN") return "DIKEMBALIKAN";
  if (s === "KEMBALI") return "DIKEMBALIKAN";
  if (s === "MENUNGGU") return "MENUNGGU";
  return s;
}

/* ============================================================
 * UTIL: DETEKSI HEADER OTOMATIS PER SHEET
 * ============================================================ */
function normalizeHeader_(h) {
  return String(h ?? "")
    .replace(/\u00A0/g, " ")
    .trim()
    .replace(/\s+/g, " ")
    .toLowerCase();
}

function getHeaderInfo_(sheet, sheetName) {
  const lastCol = sheet.getLastColumn();
  const maxScanRow = Math.min(6, sheet.getLastRow());

  const requiredBySheet = {
    "SP":        ["tanggal surat pesanan", "email pemohon"],
    "SPK":       ["tanggal spk", "email pemohon"],
    "SPD":       ["nomor spd", "email pemohon"],
    "SPJ":       ["email pemohon", "status ppk"],
    "Arsip SPJ": ["nomor sptb"]
  };

  const required = requiredBySheet[sheetName] || ["email pemohon"];

  let best = { row: 1, headers: [], headersNorm: [], score: -999 };

  for (let r = 1; r <= maxScanRow; r++) {
    const headers = sheet.getRange(r, 1, 1, lastCol).getValues()[0];
    const headersNorm = headers.map(normalizeHeader_);

    const filled = headersNorm.filter(Boolean).length;

    let hit = 0;
    for (const kw of required) {
      if (headersNorm.includes(normalizeHeader_(kw))) hit += 1;
    }

    const score = filled + (hit * 100);

    if (score > best.score) best = { row: r, headers, headersNorm, score };
  }

  return best;
}

function hasHeader_(headerInfo, headerName) {
  return headerInfo.headersNorm.includes(normalizeHeader_(headerName));
}

function mustCol_(headerInfo, headerName, sheetName) {
  const idx = headerInfo.headersNorm.indexOf(normalizeHeader_(headerName));
  if (idx === -1) {
    const sample = headerInfo.headers.slice(0, 12).map(x => String(x ?? "")).join(" | ");
    throw new Error(
      `Kolom header tidak ditemukan: "${headerName}" (Sheet: ${sheetName}). ` +
      `Header terdeteksi di baris ${headerInfo.row}. Contoh header: ${sample}`
    );
  }
  return idx + 1;
}

/* ============================================================
 * FORMAT TANGGAL INDONESIA (dd/MM/yyyy)
 * ============================================================ */
function formatTanggalID_(dateObj) {
  try {
    return Utilities.formatDate(dateObj, "Asia/Jakarta", "dd/MM/yyyy");
  } catch (e) {
    return "-";
  }
}

function getCellByHeader_(sheet, row, headerInfo, headerName) {
  const idx = headerInfo.headersNorm.indexOf(normalizeHeader_(headerName));
  if (idx === -1) return "-";

  const v = sheet.getRange(row, idx + 1).getValue();

  // Jika bertipe Date, tampilkan format Indonesia
  if (v instanceof Date && !isNaN(v.getTime())) {
    return formatTanggalID_(v);
  }

  const s = String(v ?? "").trim();
  return s ? s : "-";
}

/* ============================================================
 * UTIL: Nomor urut per tahun (SP/SPK)
 * ============================================================ */
function nextUrutPerTahun_(sheet, currentRow, headerInfo, colTanggal, colNoUrut, tahun) {
  const lastRow = sheet.getLastRow();
  let maxUrut = 0;

  for (let r = headerInfo.row + 1; r <= lastRow; r++) {
    if (r === currentRow) continue;

    const tglCell = sheet.getRange(r, colTanggal).getValue();
    const urutCell = sheet.getRange(r, colNoUrut).getValue();

    const t = parseTanggalID_(tglCell);
    if (t && urutCell !== "" && urutCell !== null) {
      if (t.getFullYear() === tahun) {
        maxUrut = Math.max(maxUrut, Number(urutCell) || 0);
      }
    }
  }

  return maxUrut + 1;
}

/* ============================================================
 * UTIL: SPD start urut global per tahun
 * ============================================================ */
/* ============================================================
 * UTIL: SPD start urut global per tahun (PERBAIKAN)
 * ============================================================ */
function nextUrutSPDStart_(sheet, currentRow, headerInfo, colNoUrut, tahun) {
  const lastRow = sheet.getLastRow();
  let maxUrut = 0;

  let timestampCol = -1;
  if (hasHeader_(headerInfo, "Timestamp")) timestampCol = mustCol_(headerInfo, "Timestamp", "SPD");

  for (let r = headerInfo.row + 1; r <= lastRow; r++) {
    // Lewati baris yang sedang diedit agar tidak menghitung dirinya sendiri
    if (r === currentRow) continue;

    // Cek Tahun (jika ada kolom Timestamp)
    if (timestampCol !== -1) {
      const tsVal = sheet.getRange(r, timestampCol).getValue();
      const ts = parseTanggalID_(tsVal);
      // Jika tanggal tidak valid atau tahun beda, lewati
      if (ts && ts.getFullYear() !== tahun) continue;
    }

    // Ambil nilai nomor urut
    const urutCell = String(sheet.getRange(r, colNoUrut).getValue() || "");
    if (!urutCell) continue;

    // --- BAGIAN PERBAIKAN ---
    // Cari semua angka dalam sel tersebut (misal: "05 s/d 09" -> ketemu [5, 9])
    const matches = urutCell.match(/\d+/g);
    
    if (matches && matches.length > 0) {
      // Ambil angka terakhir (biasanya angka terbesar dalam range)
      const currentVal = Number(matches[matches.length - 1]);
      if (currentVal > maxUrut) {
        maxUrut = currentVal;
      }
    }
    // ------------------------
  }

  return maxUrut + 1;
}

/* ============================================================
 * SPJ CONTEXT
 * ============================================================ */
function buildSPJContext_(sheet, row, headerInfo) {
  const pick = (name) => hasHeader_(headerInfo, name) ? getCellByHeader_(sheet, row, headerInfo, name) : "-";

  const nomorSPBY =
    (hasHeader_(headerInfo, "Nomor SPBY") ? pick("Nomor SPBY") : "") ||
    (hasHeader_(headerInfo, "Nomor SPBy") ? pick("Nomor SPBy") : "") ||
    (hasHeader_(headerInfo, "No. SPP/SPM") ? pick("No. SPP/SPM") : "-") ||
    "-";

  const noSPM =
    (hasHeader_(headerInfo, "No. SPM") ? pick("No. SPM") : "") ||
    (hasHeader_(headerInfo, "No SPM") ? pick("No SPM") : "") ||
    "-";

  const ctx = {
    "Program": pick("Program"),
    "Kegiatan": pick("Kegiatan"),
    "Output": pick("Output"),
    "Sub Output": pick("Sub Output"),
    "Komponen / Subkomponen": (hasHeader_(headerInfo, "Komponen / Subkomponen") ? pick("Komponen / Subkomponen") : pick("Komponen/Subkomponen")),
    "Uraian Kegiatan": pick("Uraian Kegiatan"),

    "Nama Penerima": pick("Nama Penerima"),
    "Mekanisme Pembayaran": pick("Mekanisme Pembayaran"),
    "Jumlah": pick("Jumlah"),

    "PPN (jika ada)": pick("PPN (jika ada)"),
    "PPh 21 (jika ada)": pick("PPh 21 (jika ada)"),
    "PPh 22 (jika ada)": pick("PPh 22 (jika ada)"),
    "PPh 23 (jika ada)": pick("PPh 23 (jika ada)"),
    "PPh Pasal 4 Ayat 2 Final (jika ada)": pick("PPh Pasal 4 Ayat 2 Final (jika ada)"),

    "Nomor Bukti Kuitansi": pick("Nomor Bukti Kuitansi"),
    "Nomor SPTB": pick("Nomor SPTB"),
    "Nomor SPBY": nomorSPBY,

    "Status PPK": pick("Status PPK"),
    "Tgl Kirim PPK": (hasHeader_(headerInfo, "Tgl Kirim PPK") ? pick("Tgl Kirim PPK") : "-"),
    "Status Sekretariat": pick("Status Sekretariat"),
    "Tgl Kirim Sekretariat": (hasHeader_(headerInfo, "Tgl Kirim Sekretariat") ? pick("Tgl Kirim Sekretariat") : "-"),
    "Status Bendahara": pick("Status Bendahara"),
    "Tgl Kirim Bendahara": (hasHeader_(headerInfo, "Tgl Kirim Bendahara") ? pick("Tgl Kirim Bendahara") : "-"),
    "Status SPM": pick("Status SPM"),
    "No. SPM": noSPM,
    "Tgl Kirim SPM": (hasHeader_(headerInfo, "Tgl Kirim SPM") ? pick("Tgl Kirim SPM") : "-"),

    "Petugas Input Data": pick("Petugas Input Data"),
    "Keterangan": pick("Keterangan"),
  };

  return ctx;
}

/* ============================================================
 * EMAIL BUILDERS (SP/SPK/SPD)
 * ============================================================ */
function buildEmailSP_(nomorSP, tanggalObj, data) {
  const esc = escHtml_;
  const rupiah = formatRupiah_;

  return `
    <div style="font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif;color:#333;max-width:650px;border:1px solid #e0e0e0;padding:25px;border-radius:10px;box-shadow:0 2px 8px rgba(0,0,0,0.06);">
      <h2 style="color:#2c3e50;border-bottom:3px solid #6f42c1;padding-bottom:15px;margin-top:0;">Notifikasi Penomoran Surat Pesanan (SP)</h2>
      <p style="font-size:16px;line-height:1.6;margin:0 0 10px 0;">Yth. Pemohon,</p>

      <table style="width:100%;border-collapse:collapse;margin:18px 0;font-size:15px;">
        <tr style="background-color:#f8f9fa;">
          <td style="padding:12px;border:1px solid #e0e0e0;font-weight:bold;width:35%;">Nomor SP</td>
          <td style="padding:12px;border:1px solid #e0e0e0;color:#4a148c;font-weight:bold;font-size:1.12em;">${esc(nomorSP)}</td>
        </tr>
        <tr><td style="padding:12px;border:1px solid #e0e0e0;font-weight:bold;">Nama Penyedia Jasa</td><td style="padding:12px;border:1px solid #e0e0e0;">${esc(data["Nama Rekanan/Pihak Ketiga/Penyedia Jasa"])}</td></tr>
        <tr style="background-color:#fcfcfd;"><td style="padding:12px;border:1px solid #e0e0e0;font-weight:bold;">Uraian Paket</td><td style="padding:12px;border:1px solid #e0e0e0;">${esc(data["Uraian Paket Pekerjaan"])}</td></tr>
        <tr><td style="padding:12px;border:1px solid #e0e0e0;font-weight:bold;">Mekanisme Pembayaran</td><td style="padding:12px;border:1px solid #e0e0e0;">${esc(data["Mekanisme Pembayaran"])}</td></tr>
        <tr style="background-color:#fcfcfd;"><td style="padding:12px;border:1px solid #e0e0e0;font-weight:bold;">Nilai SPJ</td><td style="padding:12px;border:1px solid #e0e0e0;">Rp ${esc(rupiah(data["Nilai SPJ"]))}</td></tr>
        <tr><td style="padding:12px;border:1px solid #e0e0e0;font-weight:bold;">Tanggal SP</td><td style="padding:12px;border:1px solid #e0e0e0;">${esc(formatTanggalID_(tanggalObj))}</td></tr>
        <tr style="background-color:#fcfcfd;"><td style="padding:12px;border:1px solid #e0e0e0;font-weight:bold;">Komponen/Subkomponen</td><td style="padding:12px;border:1px solid #e0e0e0;">${esc(data["Komponen/Subkomponen"])}</td></tr>
        <tr><td style="padding:12px;border:1px solid #e0e0e0;font-weight:bold;">Akun</td><td style="padding:12px;border:1px solid #e0e0e0;">${esc(data["Akun"])}</td></tr>
        <tr style="background-color:#fcfcfd;"><td style="padding:12px;border:1px solid #e0e0e0;font-weight:bold;">Petugas Input Data</td><td style="padding:12px;border:1px solid #e0e0e0;">${esc(data["Petugas Input Data"])}</td></tr>
        <tr><td style="padding:12px;border:1px solid #e0e0e0;font-weight:bold;">Keterangan</td><td style="padding:12px;border:1px solid #e0e0e0;">${esc(data["Keterangan"])}</td></tr>
      </table>

      <p style="font-size:12px;color:#757575;margin-top:26px;text-align:center;">Sistem Penomoran Otomatis BBWS Mesuji Sekampung &copy; ${new Date().getFullYear()}</p>
    </div>
  `;
}

function buildEmailSPK_(nomorSPK, tanggalObj, data) {
  const esc = escHtml_;
  const rupiah = formatRupiah_;

  return `
    <div style="font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif;color:#333;max-width:650px;border:1px solid #e0e0e0;padding:25px;border-radius:10px;box-shadow:0 2px 8px rgba(0,0,0,0.06);">
      <h2 style="color:#2c3e50;border-bottom:3px solid #0d6efd;padding-bottom:15px;margin-top:0;">Notifikasi Penomoran SPK</h2>
      <p style="font-size:16px;line-height:1.6;margin:0 0 10px 0;">Yth. Pemohon,</p>

      <table style="width:100%;border-collapse:collapse;margin:18px 0;font-size:15px;">
        <tr style="background-color:#f8f9fa;">
          <td style="padding:12px;border:1px solid #e0e0e0;font-weight:bold;width:35%;">Nomor SPK</td>
          <td style="padding:12px;border:1px solid #e0e0e0;color:#0b5ed7;font-weight:bold;font-size:1.12em;">${esc(nomorSPK)}</td>
        </tr>
        <tr><td style="padding:12px;border:1px solid #e0e0e0;font-weight:bold;">Nama Penyedia Jasa</td><td style="padding:12px;border:1px solid #e0e0e0;">${esc(data["Nama Rekanan/Pihak Ketiga/Penyedia Jasa"])}</td></tr>
        <tr style="background-color:#fcfcfd;"><td style="padding:12px;border:1px solid #e0e0e0;font-weight:bold;">Uraian SPK</td><td style="padding:12px;border:1px solid #e0e0e0;">${esc(data["Uraian SPK"])}</td></tr>
        <tr><td style="padding:12px;border:1px solid #e0e0e0;font-weight:bold;">Mekanisme Pembayaran</td><td style="padding:12px;border:1px solid #e0e0e0;">${esc(data["Mekanisme Pembayaran"])}</td></tr>
        <tr style="background-color:#fcfcfd;"><td style="padding:12px;border:1px solid #e0e0e0;font-weight:bold;">Nilai Kontrak</td><td style="padding:12px;border:1px solid #e0e0e0;">Rp ${esc(rupiah(data["Nilai Kontrak"]))}</td></tr>
        <tr><td style="padding:12px;border:1px solid #e0e0e0;font-weight:bold;">Tanggal SPK</td><td style="padding:12px;border:1px solid #e0e0e0;">${esc(formatTanggalID_(tanggalObj))}</td></tr>
      </table>

      <p style="font-size:12px;color:#757575;margin-top:26px;text-align:center;">Sistem Penomoran Otomatis BBWS Mesuji Sekampung &copy; ${new Date().getFullYear()}</p>
    </div>
  `;
}

function buildEmailSPD_(nomorList, d) {
  const esc = escHtml_;

  const listHtml = nomorList
    .map(n => `<li style="margin:6px 0;"><b style="color:#0b5ed7;">${esc(n)}</b></li>`)
    .join("");

  return `
    <div style="font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif;color:#333;max-width:650px;border:1px solid #e0e0e0;padding:25px;border-radius:10px;box-shadow:0 2px 8px rgba(0,0,0,0.06);">
      <h2 style="color:#2c3e50;border-bottom:3px solid #20c997;padding-bottom:15px;margin-top:0;">Notifikasi Penomoran SPD</h2>
      <p style="font-size:16px;line-height:1.6;margin:0 0 12px 0;">Yth. Pemohon,</p>

      <table style="width:100%;border-collapse:collapse;margin:18px 0;font-size:14px;">
        <tr style="background:#f8f9fa;">
          <td style="padding:10px;border:1px solid #ddd;font-weight:bold;width:35%;">Pelaksana</td>
          <td style="padding:10px;border:1px solid #ddd;">${esc(d.namaPelaksana)}</td>
        </tr>
        <tr>
          <td style="padding:10px;border:1px solid #ddd;font-weight:bold;">Jumlah Pengikut</td>
          <td style="padding:10px;border:1px solid #ddd;">${esc(d.pengikut)} Orang</td>
        </tr>
        <tr style="background:#f8f9fa;">
          <td style="padding:10px;border:1px solid #ddd;font-weight:bold;">Nama Pengikut</td>
          <td style="padding:10px;border:1px solid #ddd;white-space:pre-line;">${esc(d.namaPengikut)}</td>
        </tr>
        <tr>
          <td style="padding:10px;border:1px solid #ddd;font-weight:bold;">Maksud Perjalanan</td>
          <td style="padding:10px;border:1px solid #ddd;">${esc(d.maksud)}</td>
        </tr>
      </table>

      <div style="background:#f8fbff;border:1px solid #dbe8ff;border-radius:8px;padding:14px 16px;margin:14px 0;">
        <p style="margin:0 0 8px 0;font-weight:bold;">Daftar Nomor SPD:</p>
        <ol style="margin:0 0 0 18px;padding:0;">
          ${listHtml}
        </ol>
      </div>

      <p style="font-size:12px;color:#757575;margin-top:26px;text-align:center;">
        Sistem Penomoran Otomatis BBWS Mesuji Sekampung &copy; ${new Date().getFullYear()}
      </p>
    </div>
  `;
}

/* ============================================================
 * EMAIL SPJ (FULL)
 * ============================================================ */
function buildEmailSPJSubmitFull_(ctx) {
  return buildEmailSPJBaseFull_("SPJ diterima", "Pengajuan SPJ Anda sudah diterima dan menunggu persetujuan PPK.", ctx, "#ffc107");
}

function buildEmailSPJProgressFull_(tahap, waitingText, ctx) {
  const subtitle = `Pengajuan SPJ Anda telah disetujui oleh <b>${escHtml_(tahap)}</b>. Status berikutnya: <b>${escHtml_(waitingText)}</b>.`;
  return buildEmailSPJBaseFull_(`SPJ disetujui oleh ${tahap}`, subtitle, ctx, "#198754");
}

function buildEmailSPJReturnedFull_(tahap, ctx) {
  const subtitle = `Pengajuan SPJ Anda <b>DIKEMBALIKAN</b> oleh <b>${escHtml_(tahap)}</b>. Silakan cek kolom <b>Keterangan</b>.`;
  return buildEmailSPJBaseFull_(`SPJ dikembalikan oleh ${tahap}`, subtitle, ctx, "#dc3545");
}

function buildEmailSPJFinalBySPMFull_(ctx) {
  const subtitle = "Pengajuan SPJ Anda sudah <b>DISETUJUI</b> sampai tahap SPM (SELESAI).";
  return buildEmailSPJBaseFull_("SPJ selesai (Disetujui sampai SPM)", subtitle, ctx, "#0d6efd");
}

function buildEmailSPJBaseFull_(title, subtitleHtml, ctx, accentColor) {
  const esc = escHtml_;
  const rupiah = formatRupiah_;

  const rowsTop = [
    ["Nomor Bukti Kuitansi", ctx["Nomor Bukti Kuitansi"]],
    ["Nomor SPTB", ctx["Nomor SPTB"]],
    ["Nomor SPBY", ctx["Nomor SPBY"]],
    ["Nama Penerima", ctx["Nama Penerima"]],
    ["Mekanisme Pembayaran", ctx["Mekanisme Pembayaran"]],
    ["Jumlah", ctx["Jumlah"] ? `Rp ${esc(rupiah(ctx["Jumlah"]))}` : "-"],
  ];

  const rowsProgram = [
    ["Program", ctx["Program"]],
    ["Kegiatan", ctx["Kegiatan"]],
    ["Output", ctx["Output"]],
    ["Sub Output", ctx["Sub Output"]],
    ["Komponen / Subkomponen", ctx["Komponen / Subkomponen"]],
    ["Uraian Kegiatan", ctx["Uraian Kegiatan"]],
  ];

  const rowsPajak = [
    ["PPN (jika ada)", ctx["PPN (jika ada)"]],
    ["PPh 21 (jika ada)", ctx["PPh 21 (jika ada)"]],
    ["PPh 22 (jika ada)", ctx["PPh 22 (jika ada)"]],
    ["PPh 23 (jika ada)", ctx["PPh 23 (jika ada)"]],
    ["PPh Pasal 4 Ayat 2 Final (jika ada)", ctx["PPh Pasal 4 Ayat 2 Final (jika ada)"]],
  ];

  // Status Proses + No. SPM tepat di bawah Status SPM
  const rowsStatus = [
    ["Status PPK", ctx["Status PPK"]],
    ["Tgl Kirim PPK", ctx["Tgl Kirim PPK"]],
    ["Status Sekretariat", ctx["Status Sekretariat"]],
    ["Tgl Kirim Sekretariat", ctx["Tgl Kirim Sekretariat"]],
    ["Status Bendahara", ctx["Status Bendahara"]],
    ["Tgl Kirim Bendahara", ctx["Tgl Kirim Bendahara"]],
    ["Status SPM", ctx["Status SPM"]],
    ["No. SPM", ctx["No. SPM"]],
    ["Tgl Kirim SPM", ctx["Tgl Kirim SPM"]],
  ];

  const rowsMeta = [
    ["Petugas Input Data", ctx["Petugas Input Data"]],
    ["Keterangan", ctx["Keterangan"]],
  ];

  const block = (label, rows) => `
    <h3 style="margin:18px 0 10px 0;font-size:16px;color:#2c3e50;">${esc(label)}</h3>
    ${buildTable_(rows)}
  `;

  return `
    <div style="font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif;color:#333;max-width:700px;border:1px solid #e0e0e0;padding:25px;border-radius:10px;box-shadow:0 2px 8px rgba(0,0,0,0.06);">
      <h2 style="color:#2c3e50;border-bottom:3px solid ${esc(accentColor)};padding-bottom:15px;margin-top:0;">${esc(title)}</h2>
      <p style="font-size:15px;line-height:1.7;margin:0 0 12px 0;">Yth. Pemohon,</p>
      <p style="font-size:15px;line-height:1.7;margin:0 0 14px 0;">${subtitleHtml}</p>

      ${block("Identitas & Nomor", rowsTop)}
      ${block("Program & Rincian Kegiatan", rowsProgram)}
      ${block("Pajak", rowsPajak)}
      ${block("Status Proses", rowsStatus)}
      ${block("Catatan & Petugas", rowsMeta)}

      <p style="font-size:12px;color:#757575;margin-top:26px;text-align:center;">Sistem Penomoran Otomatis BBWS Mesuji Sekampung &copy; ${new Date().getFullYear()}</p>
    </div>
  `;
}

function buildTable_(rows) {
  const esc = escHtml_;
  const filtered = (rows || []).filter(r => r && r.length >= 2);

  let html = `<table style="width:100%;border-collapse:collapse;font-size:14.5px;">`;
  for (let i = 0; i < filtered.length; i++) {
    const bg = (i % 2 === 0) ? "#fcfcfd" : "#ffffff";
    html += `
      <tr style="background-color:${bg};">
        <td style="padding:11px 12px;border:1px solid #e0e0e0;font-weight:600;width:38%;">${esc(filtered[i][0])}</td>
        <td style="padding:11px 12px;border:1px solid #e0e0e0;">${esc(filtered[i][1] ?? "-")}</td>
      </tr>
    `;
  }
  html += `</table>`;
  return html;
}

/* ============================================================
 * UTIL: escape html dan format rupiah + parse tanggal
 * ============================================================ */
function escHtml_(s) {
  return String(s ?? "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function formatRupiah_(val) {
  const n = String(val ?? "").replace(/[^\d]/g, "");
  if (!n) return "-";
  return n.replace(/\B(?=(\d{3})+(?!\d))/g, ".");
}

function parseTanggalID_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value;

  const s = String(value || "").trim();
  if (!s) return null;

  // dd/MM/yyyy HH:mm:ss
  let m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})\s+(\d{1,2}):(\d{2})(?::(\d{2}))?$/);
  if (m) {
    const d = Number(m[1]), mo = Number(m[2]), y = Number(m[3]);
    const hh = Number(m[4]), mm = Number(m[5]), ss = Number(m[6] || 0);
    const dt = new Date(y, mo - 1, d, hh, mm, ss);
    if (dt.getFullYear() === y && dt.getMonth() === (mo - 1) && dt.getDate() === d) return dt;
    return null;
  }

  // dd/MM/yyyy
  m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/);
  if (m) {
    const d = Number(m[1]);
    const mo = Number(m[2]);
    const y = Number(m[3]);
    const dt = new Date(y, mo - 1, d);
    if (dt.getFullYear() === y && dt.getMonth() === (mo - 1) && dt.getDate() === d) return dt;
    return null;
  }

  const dt2 = new Date(s);
  if (!isNaN(dt2.getTime())) return dt2;

  return null;
}
