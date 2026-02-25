/**
 * ============================================================
 * CONFIGURATION & HELPER FUNCTIONS
 * ============================================================
 */

function formatTanggalIndo(date) {
  if (!(date instanceof Date)) return date;
  var bulanIndo = [
    "Januari", "Februari", "Maret", "April", "Mei", "Juni",
    "Juli", "Agustus", "September", "Oktober", "November", "Desember"
  ];
  return date.getDate() + " " + bulanIndo[date.getMonth()] + " " + date.getFullYear();
}

function formatWaktu(waktu) {
  if (waktu instanceof Date) {
    return Utilities.formatDate(waktu, "GMT+7", "HH:mm");
  }
  return waktu ? waktu.toString() : "";
}

/**
 * ============================================================
 * 1. FUNGSI SAAT FORM DISUBMIT
 * TRIGGER: ON FORM SUBMIT (INSTALLABLE)
 * ============================================================
 */
function kirimNotifikasiSubmit(e) {
  try {
    if (!e || !e.range) return;

    var sheet = e.range.getSheet();
    var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    var row = e.range.getRow();

    // Set Status = DIPROSES
    var colStatusIndex = headers.indexOf("Status");
    if (colStatusIndex !== -1) {
      sheet.getRange(row, colStatusIndex + 1).setValue("DIPROSES");
    }

    // Ambil email berdasarkan header (AMAN)
    var colEmailIndex = headers.indexOf("Email Address");
    if (colEmailIndex === -1) return;

    var rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];
    var emailPemohon = rowData[colEmailIndex];

    if (!emailPemohon || !emailPemohon.toString().includes("@")) return;

    var subjek = "Notifikasi – Permohonan Diterima";
    var pesan =
      "Yth. Bapak/Ibu Pemohon,\n\n" +
      "Permohonan peminjaman ruangan yang Anda ajukan telah kami terima dan dicatat dalam sistem.\n\n" +
      "Selanjutnya permohonan akan kami tindaklanjuti sesuai ketersediaan ruangan dan jadwal.\n\n" +
      "Terima kasih.\n\n" +
      "Hormat kami,\n" +
      "Pengelola Peminjaman Ruangan\n" +
      "BBWS Mesuji Sekampung";

    MailApp.sendEmail(emailPemohon, subjek, pesan);

  } catch (err) {
    Logger.log("Error Submit: " + err.toString());
  }
}

/**
 * ============================================================
 * 2. FUNGSI SAAT EDIT STATUS
 * TRIGGER: ON EDIT (INSTALLABLE)
 * ============================================================
 */
function kirimNotifikasiEdit(e) {

  // === VALIDASI EVENT DASAR ===
  if (!e || !e.range) return;

  var range = e.range;
  var sheet = range.getSheet();

  // 1. Batasi hanya sheet tertentu
  if (sheet.getName() !== "Status Pengajuan Ruangan") return;

  // 2. Abaikan header
  if (range.getRow() === 1) return;

  // 3. Abaikan paste / edit banyak sel
  if (range.getNumRows() > 1 || range.getNumColumns() > 1) return;

  // 4. Tentukan kolom STATUS dari header
  var headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  var colStatusIndex = headers.indexOf("Status");
  if (colStatusIndex === -1) return;

  // 5. HANYA lanjut jika kolom STATUS yang diedit
  if (range.getColumn() !== colStatusIndex + 1) return;

  var statusValue = range.getValue();
  if (statusValue !== "DISETUJUI" && statusValue !== "DITOLAK") return;

  // === LOCK ===
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(10000)) return;

  try {
    var row = range.getRow();
    var rowData = sheet.getRange(row, 1, 1, sheet.getLastColumn()).getValues()[0];

    var data = {};
    headers.forEach(function (h, i) {
      data[h] = rowData[i];
    });

    var emailPemohon = data["Email Address"];
    if (!emailPemohon || !emailPemohon.toString().includes("@")) return;

    var tempat  = data["Tempat"];
    var agenda  = data["Agenda"];

    var tglString = formatTanggalIndo(data["Tanggal Mulai Kegiatan"]);
    var waktuGabung =
      data["Waktu Mulai"] && data["Waktu Selesai"]
        ? formatWaktu(data["Waktu Mulai"]) + " - " + formatWaktu(data["Waktu Selesai"])
        : "";

    var subjek = "";
    var pesan  = "";

    // === LOGIKA STATUS ===
    if (statusValue === "DISETUJUI") {

      var sudahAda = cekDuplikatDiAgenda(tempat, tglString, waktuGabung, agenda);

      if (!sudahAda) {
        salinKeAgenda(tempat, tglString, waktuGabung, agenda);
      }

      subjek = "Notifikasi – Permohonan Disetujui";
      pesan =
        "Yth. Bapak/Ibu Pemohon,\n\n" +
        "Permohonan peminjaman ruangan yang Anda ajukan telah disetujui dengan rincian sebagai berikut:\n\n" +
        "Ruangan : " + tempat + "\n" +
        "Tanggal : " + tglString + "\n" +
        "Waktu   : " + waktuGabung + "\n" +
        "Kegiatan: " + agenda + "\n\n" +
        "Mohon menggunakan ruangan sesuai jadwal dan menjaga ketertiban serta kebersihan.\n\n" +
        "Terima kasih.\n\n" +
        "Hormat kami,\n" +
        "Pengelola Peminjaman Ruangan\n" +
        "BBWS Mesuji Sekampung";

    } else if (statusValue === "DITOLAK") {

      subjek = "Notifikasi – Permohonan Ditolak";
      pesan =
        "Yth. Bapak/Ibu Pemohon,\n\n" +
        "Permohonan peminjaman ruangan yang Anda ajukan belum dapat kami setujui karena keterbatasan ruangan atau benturan jadwal.\n\n" +
        "Atas pengertiannya, kami ucapkan terima kasih.\n\n" +
        "Hormat kami,\n" +
        "Pengelola Peminjaman Ruangan\n" +
        "BBWS Mesuji Sekampung";
    }

    MailApp.sendEmail(emailPemohon, subjek, pesan);

  } catch (err) {
    Logger.log("Error Edit: " + err.toString());
  } finally {
    lock.releaseLock();
  }
}

/**
 * ============================================================
 * 3. CEK DUPLIKAT & SALIN KE AGENDA
 * ============================================================
 */
function cekDuplikatDiAgenda(ruangan, tanggalStr, waktuStr, agenda) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var agendaSheet = ss.getSheetByName("Agenda Ruangan");
  if (!agendaSheet) return false;

  var lastRow = agendaSheet.getLastRow();
  if (lastRow < 2) return false;

  var dataAgenda = agendaSheet.getRange(2, 1, lastRow - 1, 4).getValues();

  for (var i = 0; i < dataAgenda.length; i++) {
    var r = dataAgenda[i];
    if (
      String(r[0]) === String(ruangan) &&
      String(r[1]) === String(tanggalStr) &&
      String(r[2]) === String(waktuStr) &&
      String(r[3]) === String(agenda)
    ) {
      return true;
    }
  }
  return false;
}

function salinKeAgenda(ruangan, tanggal, waktu, agenda) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var agendaSheet = ss.getSheetByName("Agenda Ruangan");
  if (agendaSheet) {
    agendaSheet.appendRow([ruangan, tanggal, waktu, agenda]);
  }
}
