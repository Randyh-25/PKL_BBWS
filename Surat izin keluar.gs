/***********************************************************
 * KONFIGURASI UNIVERSAL (GEDUNG 1 & 2)
 ***********************************************************/
const FORM_URL_GEDUNG_1 = "https://docs.google.com/forms/d/e/1FAIpQLScUv5jIBgIiIqagknQUgxxzqFXnxOT6DyfTdud67XryIexCDw/viewform?usp=pp_url";
const FORM_URL_GEDUNG_2 = "https://docs.google.com/forms/d/e/1FAIpQLSczKcTK7mCuHNvMUXeMxcQQUBdb3BKt0GDxYjdd7dOdXMGpTw/viewform?usp=pp_url";

const ENTRY_REQUEST_ID = "entry.1199147859";
const ENTRY_KEPUTUSAN  = "entry.1396566311";

function mainOnFormSubmit(e) {
  if (!e || !e.range) return;
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();

  if (sheetName.includes("DATA IZIN")) {
    prosesFormIzin(sheet, e.range.getRow());
  } else if (sheetName.includes("APPROVAL")) {
    prosesFormApprovalUniversal(sheet, e.range.getRow());
  }
}

/**
 * 1. PROSES IZIN MASUK (Email Ke Atasan)
 */
function prosesFormIzin(sheet, row) {
  const sheetName = sheet.getName();
  const tz = Session.getScriptTimeZone();
  
  // Generate Request ID Otomatis
  const timestamp = sheet.getRange(row, 1).getValue(); 
  const formattedDate = Utilities.formatDate(new Date(timestamp), tz, "yyyyMMdd-HHmmss");
  const requestId = "REQ-" + formattedDate;
  
  // Buat kolom jika belum ada
  getOrCreateCol_(sheet, "Request ID");
  updateCellByHeader(sheet, row, "Request ID", requestId);
  
  // Sinkronisasi Nama Master
  setupNamaMaster(sheet, row);

  const nama = ambilNilaiSmart(sheet, row, "Nama (Master)");
  const emailAtas = ambilNilaiSmart(sheet, row, "Email Atasan Langsung") || ambilNilaiSmart(sheet, row, "Atasan Langsung");
  const tglIzin = ambilNilaiSmart(sheet, row, "Tanggal Izin");

  const baseUrl = sheetName.includes("GEDUNG 1") ? FORM_URL_GEDUNG_1 : FORM_URL_GEDUNG_2;

  const setujui = `${baseUrl}&${ENTRY_REQUEST_ID}=${encodeURIComponent(requestId)}&${ENTRY_KEPUTUSAN}=DISETUJUI`;
  const tolak   = `${baseUrl}&${ENTRY_REQUEST_ID}=${encodeURIComponent(requestId)}&${ENTRY_KEPUTUSAN}=DITOLAK`;

  const subject = `[PERSETUJUAN] Izin Keluar Kantor – ${nama || "(Nama belum terisi)"}`;
  
  const htmlBodyAtasan = `
    <p>Yth. Bapak/Ibu,</p>
    <p>Mohon persetujuan izin keluar kantor dengan rincian berikut:</p>
    <table cellpadding="4" cellspacing="0">
      <tr><td><b>Request ID</b></td><td>: ${requestId}</td></tr>
      <tr><td><b>Nama</b></td><td>: ${nama || "-"}</td></tr>
      <tr><td><b>Unit Kerja</b></td><td>: ${ambilNilaiSmart(sheet, row, "Unit Kerja") || "-"}</td></tr>
      <tr><td><b>Hari/Tanggal</b></td><td>: ${formatTanggal(new Date(tglIzin))}</td></tr>
      <tr><td><b>Waktu</b></td><td>: ${Utilities.formatDate(new Date(ambilNilaiSmart(sheet, row, "Jam Keluar")), tz, "HH.mm")} - ${Utilities.formatDate(new Date(ambilNilaiSmart(sheet, row, "Jam Kembali")), tz, "HH.mm")} WIB</td></tr>
      <tr><td><b>Keperluan</b></td><td>: ${ambilNilaiSmart(sheet, row, "Keperluan", true) || "-"}</td></tr>
      <tr><td><b>Uraian</b></td><td>: ${ambilNilaiSmart(sheet, row, "Uraian Keperluan") || "-"}</td></tr>
      <tr><td><b>Email Pemohon</b></td><td>: ${ambilNilaiSmart(sheet, row, "Email Pemohon") || "-"}</td></tr>
    </table>
    <p><b>Silakan pilih:</b></p>
    <div style="margin-top: 15px;">
      <a href="${setujui}" style="background-color: #28a745; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px; font-weight: bold; display: inline-block; margin-right: 10px;">SETUJUI</a>
      <a href="${tolak}" style="background-color: #dc3545; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px; font-weight: bold; display: inline-block;">TOLAK</a>
    </div>
    <br><p>BBWS Mesuji Sekampung</p>`;

  if (emailAtas) {
    MailApp.sendEmail({ to: String(emailAtas).trim(), subject: subject, htmlBody: htmlBodyAtasan });
  }
}

/**
 * 2. PROSES APPROVAL (Update SPS & Email Hasil Ke Pemohon)
 */
function prosesFormApprovalUniversal(approvalSheet, row) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const tz = Session.getScriptTimeZone();
  const sheetName = approvalSheet.getName();
  const targetSheetName = sheetName.replace("APPROVAL", "DATA IZIN");
  const sheetIzin = ss.getSheetByName(targetSheetName);

  const requestId = String(ambilNilaiSmart(approvalSheet, row, "Request ID")).trim();
  const keputusan = ambilNilaiSmart(approvalSheet, row, "Keputusan");
  const catatan   = ambilNilaiSmart(approvalSheet, row, "Catatan Atasan") || "-";

  if (!requestId || !sheetIzin) return;

  const headersIzin = sheetIzin.getRange(1, 1, 1, sheetIzin.getLastColumn()).getValues()[0];
  const colReqId = headersIzin.findIndex(h => h.toString().trim() === "Request ID") + 1;
  const dataReqId = sheetIzin.getRange(2, colReqId, sheetIzin.getLastRow()-1, 1).getValues().flat();
  const targetIdx = dataReqId.indexOf(requestId);

  if (targetIdx !== -1) {
    const targetRow = targetIdx + 2;
    
    const emailAtasanAsli = ambilNilaiSmart(sheetIzin, targetRow, "Email Atasan Langsung") || ambilNilaiSmart(sheetIzin, targetRow, "Atasan Langsung");

    // Otomatis buat kolom di DATA IZIN jika belum ada
    getOrCreateCol_(sheetIzin, "Status");
    getOrCreateCol_(sheetIzin, "Approved By (Email Atasan)");
    getOrCreateCol_(sheetIzin, "Catatan Atasan");
    getOrCreateCol_(sheetIzin, "Waktu Persetujuan");

    updateCellByHeader(sheetIzin, targetRow, "Status", keputusan);
    updateCellByHeader(sheetIzin, targetRow, "Catatan Atasan", catatan);
    updateCellByHeader(sheetIzin, targetRow, "Approved By (Email Atasan)", emailAtasanAsli);
    updateCellByHeader(sheetIzin, targetRow, "Waktu Persetujuan", new Date());

    const emailPemohon = ambilNilaiSmart(sheetIzin, targetRow, "Email Pemohon");
    const nama         = ambilNilaiSmart(sheetIzin, targetRow, "Nama (Master)");
    const unit         = ambilNilaiSmart(sheetIzin, targetRow, "Unit Kerja");
    const tglIzin      = ambilNilaiSmart(sheetIzin, targetRow, "Tanggal Izin");
    const jamKeluar    = ambilNilaiSmart(sheetIzin, targetRow, "Jam Keluar");
    const jamKembali   = ambilNilaiSmart(sheetIzin, targetRow, "Jam Kembali");
    const keperluan    = ambilNilaiSmart(sheetIzin, targetRow, "Keperluan", true);
    const uraian       = ambilNilaiSmart(sheetIzin, targetRow, "Uraian Keperluan");

    if (emailPemohon) {
      const subject = `Hasil Persetujuan Izin Keluar Kantor – ${keputusan}`;
      const hariTanggalIndo = formatTanggal(new Date(tglIzin));
      const jamMulai    = Utilities.formatDate(new Date(jamKeluar), tz, "HH.mm");
      const jamAkhir    = Utilities.formatDate(new Date(jamKembali), tz, "HH.mm");

      const htmlBodyPemohon = `
        <p>Yth. ${nama || "Bapak/Ibu"},</p>
        <p>Berikut hasil persetujuan izin keluar kantor Anda:</p>
        <table cellpadding="4" cellspacing="0">
          <tr><td><b>Request ID</b></td><td>: ${requestId}</td></tr>
          <tr><td><b>Nama</b></td><td>: ${nama || "-"}</td></tr>
          <tr><td><b>Unit Kerja</b></td><td>: ${unit || "-"}</td></tr>
          <tr><td><b>Hari/Tanggal</b></td><td>: ${hariTanggalIndo}</td></tr>
          <tr><td><b>Waktu</b></td><td>: ${jamMulai} - ${jamAkhir} WIB</td></tr>
          <tr><td><b>Keperluan</b></td><td>: ${keperluan || "-"}</td></tr>
          <tr><td><b>Uraian</b></td><td>: ${uraian || "-"}</td></tr>
          <tr><td><b>Catatan Atasan</b></td><td>: ${catatan || "-"}</td></tr>
          <tr><td><b>Status</b></td><td>: <b>${keputusan}</b></td></tr>
        </table>
        <p>BBWS Mesuji Sekampung</p>`;

      MailApp.sendEmail({ to: String(emailPemohon).trim(), subject: subject, htmlBody: htmlBodyPemohon });
    }
  }
}

/**
 * HELPER: Format Tanggal Indonesia
 */
function formatTanggal(dateObj) {
  const hari = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
  const bulan = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  return `${hari[dateObj.getDay()]}, ${dateObj.getDate()} ${bulan[dateObj.getMonth()]} ${dateObj.getFullYear()}`;
}

/**
 * HELPER: Buat kolom jika belum ada
 */
function getOrCreateCol_(sheet, headerName) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const idx = headers.findIndex(h => h.toString().trim().toLowerCase() === headerName.toLowerCase());
  if (idx !== -1) return idx + 1;
  const lastCol = sheet.getLastColumn();
  sheet.getRange(1, lastCol + 1).setValue(headerName);
  return lastCol + 1;
}

/**
 * HELPER: Sinkronisasi Nama
 */
function setupNamaMaster(sheet, row) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let colMaster = headers.findIndex(h => h.toString().trim().toLowerCase() === "nama (master)") + 1;
  if (colMaster === 0) colMaster = getOrCreateCol_(sheet, "Nama (Master)");
  const namaCols = headers.map((h, i) => ({h: h.toString().toLowerCase(), idx: i+1})).filter(x => x.h.includes("nama") && !x.h.includes("master"));
  let namaFinal = "";
  for (let col of namaCols) {
    let val = sheet.getRange(row, col.idx).getValue();
    if (val) { namaFinal = val; break; }
  }
  if (namaFinal) sheet.getRange(row, colMaster).setValue(namaFinal);
}

/**
 * HELPER: Ambil Nilai Smart
 */
function ambilNilaiSmart(sheet, row, headerName, strict = false) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  let res = "";
  for (let i = 0; i < headers.length; i++) {
    const head = headers[i].toString().trim().toLowerCase();
    const search = headerName.toLowerCase();
    if (strict ? head === search : head.includes(search)) {
      const val = sheet.getRange(row, i + 1).getValue();
      if (val !== "") res = val;
    }
  }
  return res;
}

/**
 * HELPER: Update Cell
 */
function updateCellByHeader(sheet, row, headerName, value) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const col = headers.findIndex(h => h.toString().trim() === headerName) + 1;
  if (col > 0) sheet.getRange(row, col).setValue(value);
}