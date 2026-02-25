/**
 * -----------------------------------------------------------------------
 * SISTEM PENOMORAN SURAT OTOMATIS (CLUSTER NUMBERING 2 DIGIT)
 * -----------------------------------------------------------------------
 * Fitur:
 * 1. Logika: Mengelompokkan nomor berdasarkan 2 HURUF DEPAN kode.
 *    (Contoh: BK01 dan BK02 dianggap satu kelompok "BK").
 * 2. Tampilan Email: Menggunakan format "Bagus" (Tabel & Style Segoe UI).
 * 3. Kode Otoritas: Otomatis menyesuaikan berdasarkan "Pejabat Penandatanganan".
 * 4. Derajat Kecepatan: Diambil dari "Sifat Surat (Derajat Kecepatan)" (Biasa/Segera/Amat Segera),
 *    agar tidak bentrok dengan "Kode Keamanan Naskah Dinas" yang juga punya kata "Biasa".
 * 5. KHUSUS "Keputusan" (Produk Hukum): Format nomor berbeda:
 *    (NoUrut)/KPTS/(KodeOtoritas)/(Tahun)
 *    dan nomor urut hanya menghitung sesama "Keputusan" pada tahun yang sama.
 *    Selain "Keputusan", tetap pakai format normal & urut berdasarkan 2 huruf depan kode + tahun.
 */

function onFormSubmit(e) {
  if (!e || !e.namedValues) {
    console.error("Script harus via Trigger.");
    return;
  }

  const sheet = e.range.getSheet();

  // ===================== TAMBAHAN FITUR SINKRON (TANPA EDIT BAGIAN LAIN) =====================
  // Jika submit berasal dari sheet "Digitalisasi", cek nomor -> centang checkbox "Digitalisasi" di sheet "Penomoran"
  if (sheet.getName() === "Digitalisasi") {
    sinkronDigitalisasiKePenomoran(e);
    return;
  }
  // ==========================================================================================

  if (sheet.getName() !== "Penomoran") return; // HANYA berlaku di sheet "Penomoran"

  const row = e.range.getRow();

  // --- 1. AMBIL DATA (SAFE GET) ---
  const getVal = (headerName) => {
    if (e.namedValues[headerName]) return e.namedValues[headerName][0];
    const keys = Object.keys(e.namedValues);
    const match = keys.find(k => k.trim() === headerName.trim());
    return match ? e.namedValues[match][0] : "";
  };

  const emailUser    = getVal("Email");
  const kodeArsipRaw = getVal("Kode Klasifikasi Naskah Dinas");

  // DIPISAH: keamanan vs derajat kecepatan
  const keamananRaw  = getVal("Kode Keamanan Naskah Dinas");
  const derajatRaw   = getVal("Sifat Surat (Derajat Kecepatan)");

  const hal          = getVal("Hal Naskah Dinas");
  const tanggalSurat = getVal("Tanggal Naskah Dinas");
  const keterangan   = getVal("Keterangan");
  const jenisNaskah  = getVal("Jenis Naskah Dinas");

  // pejabat penandatanganan
  const pejabat = getVal("Pejabat Penandatanganan");

  // Flag keputusan (produk hukum)
  const isKeputusan = String(jenisNaskah || "").toLowerCase().trim() === "keputusan";

  // --- 2. LOGIKA KODE (2 HURUF DEPAN) ---
  const kodeDepan = (kodeArsipRaw || "").substring(0, 2).toUpperCase() || "UM";
  const currentYear = new Date().getFullYear();

  // --- 2A. LOGIKA KODE KEAMANAN (dari "Kode Keamanan Naskah Dinas") ---
  let kodeKeamanan = "B";
  if ((keamananRaw || "").includes("(SR)")) kodeKeamanan = "SR";
  else if ((keamananRaw || "").includes("(R)"))  kodeKeamanan = "R";
  else if ((keamananRaw || "").includes("(T)"))  kodeKeamanan = "T";

  // --- 3. LOGIKA NOMOR URUT (KEPUTUSAN DIPISAH) ---
  let nomorUrut = 1;

  if (row > 2) {
    const headersRow = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

    // Cari index kolom
    let idxKode = headersRow.findIndex(h => String(h).trim() === "Kode Klasifikasi Naskah Dinas");
    let idxTgl  = headersRow.findIndex(h => {
      const t = String(h).trim();
      return t === "Tanggal Surat" || t === "Tanggal Naskah Dinas";
    });
    let idxJenis = headersRow.findIndex(h => String(h).trim() === "Jenis Naskah Dinas");

    // fallback (hindari error total kalau header berubah)
    if (idxKode === -1) idxKode = 2;
    if (idxTgl === -1) idxTgl = 7;

    const values = sheet.getRange(2, 1, row - 2, sheet.getLastColumn()).getValues();
    let counter = 0;

    for (let i = 0; i < values.length; i++) {
      const barisData = values[i];

      // Tahun baris
      let yearRow = currentYear;
      if (barisData[idxTgl]) {
        const dt = new Date(barisData[idxTgl]);
        if (!isNaN(dt.getTime())) yearRow = dt.getFullYear();
      }
      if (yearRow !== currentYear) continue;

      // Jenis baris (keputusan / bukan)
      const jenisRow = idxJenis >= 0 ? String(barisData[idxJenis] || "").toLowerCase().trim() : "";
      const rowIsKeputusan = jenisRow === "keputusan";

      if (isKeputusan) {
        // Keputusan: hanya hitung sesama keputusan
        if (!rowIsKeputusan) continue;
        counter++;
      } else {
        // Non-keputusan: hanya hitung non-keputusan + sesuai 2 huruf depan
        if (rowIsKeputusan) continue;

        const valKode = String(barisData[idxKode] || "");
        const kodeRow = valKode.substring(0, 2).toUpperCase();

        if (kodeRow === kodeDepan) counter++;
      }
    }

    nomorUrut = counter + 1;
  }

  // --- 4. FORMAT FINAL ---
  const kodeOutput = String(kodeArsipRaw || "").split(" ")[0] || "UM";

  // Kode Otoritas otomatis dari pejabat
  const kodeOtoritas = getKodeOtoritas(pejabat);

  // Format nomor surat: keputusan vs normal
  const nomorSuratLengkap = isKeputusan
    ? `${nomorUrut}/KPTS/${kodeOtoritas}/${currentYear}`
    : `${kodeOutput}/${kodeKeamanan}/${kodeOtoritas}/${currentYear}/${nomorUrut}`;

  // --- 5. SIMPAN KE SPREADSHEET ---
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const writeToColumn = (colName, value, bgColor) => {
    let colIndex = headers.indexOf(colName) + 1;
    if (colIndex === 0) {
      colIndex = sheet.getLastColumn() + 1;
      sheet.getRange(1, colIndex).setValue(colName).setBackground(bgColor).setFontWeight("bold");
      headers.push(colName);
    }
    sheet.getRange(row, colIndex).setValue(value);
  };

  writeToColumn("No. Urut", nomorUrut, "#cfe2f3");
  writeToColumn("Nomor Surat Lengkap", nomorSuratLengkap, "#d9ead3");

  // Simpan pejabat & kode otoritas
  writeToColumn("Pejabat Penandatanganan", pejabat, "#fff2cc");
  writeToColumn("Kode Otoritas", kodeOtoritas, "#fff2cc");

  // Simpan derajat kecepatan (opsional)
  writeToColumn("Sifat Surat (Derajat Kecepatan)", derajatRaw, "#fff2cc");

  // Simpan jenis naskah (opsional, biar konsisten)
  writeToColumn("Jenis Naskah Dinas", jenisNaskah, "#fff2cc");

  // --- 6. KIRIM EMAIL ---
  if (emailUser && emailUser.includes("@")) {
    kirimEmail(emailUser, nomorSuratLengkap, jenisNaskah, keamananRaw, derajatRaw, hal, tanggalSurat);
  } else {
    console.log("Email tidak valid: " + emailUser);
  }
}

/**
 * ===================== TAMBAHAN FITUR SINKRON =====================
 * Jika ada submit dari sheet "Digitalisasi":
 * Ambil nilai kolom "Nomor Naskah Dinas" (di Digitalisasi),
 * cocokkan dengan kolom "Nomor Surat Lengkap" (di Penomoran),
 * lalu set TRUE pada kolom checkbox "Digitalisasi" (di Penomoran).
 */
function sinkronDigitalisasiKePenomoran(e) {
  const sheetDigitalisasi = e.range.getSheet();

  // Ambil nomor dari form Digitalisasi (kolom: "Nomor Naskah Dinas")
  const getVal = (headerName) => {
    if (e.namedValues[headerName]) return e.namedValues[headerName][0];
    const keys = Object.keys(e.namedValues);
    const match = keys.find(k => k.trim() === headerName.trim());
    return match ? e.namedValues[match][0] : "";
  };

  const nomorInput = String(getVal("Nomor Naskah Dinas") || "").trim();
  if (!nomorInput) {
    console.log('Digitalisasi: "Nomor Naskah Dinas" kosong.');
    return;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetPenomoran = ss.getSheetByName("Penomoran");
  if (!sheetPenomoran) {
    console.log('Sheet "Penomoran" tidak ditemukan.');
    return;
  }

  const lastRow = sheetPenomoran.getLastRow();
  if (lastRow < 2) return;

  const headers = sheetPenomoran.getRange(1, 1, 1, sheetPenomoran.getLastColumn()).getValues()[0];

  const idxNomorSuratLengkap = headers.findIndex(h => String(h).trim() === "Nomor Surat Lengkap");
  const idxCheckboxDigitalisasi = headers.findIndex(h => String(h).trim() === "Digitalisasi");

  if (idxNomorSuratLengkap === -1) {
    console.log('Kolom "Nomor Surat Lengkap" tidak ditemukan di sheet Penomoran.');
    return;
  }
  if (idxCheckboxDigitalisasi === -1) {
    console.log('Kolom "Digitalisasi" tidak ditemukan di sheet Penomoran.');
    return;
  }

  // Ambil semua nomor surat lengkap untuk dicocokkan
  const nomorRange = sheetPenomoran.getRange(2, idxNomorSuratLengkap + 1, lastRow - 1, 1).getValues();

  for (let i = 0; i < nomorRange.length; i++) {
    const nomorDiPenomoran = String(nomorRange[i][0] || "").trim();
    if (nomorDiPenomoran === nomorInput) {
      // Centang checkbox Digitalisasi pada baris yang sama
      sheetPenomoran.getRange(i + 2, idxCheckboxDigitalisasi + 1).setValue(true);
      return;
    }
  }

  console.log('Nomor tidak ditemukan di Penomoran untuk sinkron: ' + nomorInput);
}

/**
 * HELPER: MAPPING KODE OTORITAS BERDASARKAN PEJABAT PENANDATANGANAN
 */
function getKodeOtoritas(pejabatRaw) {
  const norm = (s) =>
    String(s || "")
      .toLowerCase()
      .replace(/[.,/()]/g, " ")
      .replace(/\s+/g, " ")
      .trim();

  const p = norm(pejabatRaw);

  const map = {
    // A. Satker BBWS Mesuji Sekampung
    "kepala satker bbws mesuji sekampung": "Bbws2.a",

    // A.I PPK Ketatalaksanaan
    "ppk ketatalaksanaan": "Bbws2.a1",

    // A.II PPK Perencanaan dan Program
    "ppk perencanaan dan program": "Bbws2.a2",

    // A.III PPK PSDA
    "ppk psda": "Bbws2.a3",
  };
  return map[p] || "Bbws2"; // fallback default
}

/**
 * FUNGSI KIRIM EMAIL (Alert keamanan untuk T/R/SR, derajat kecepatan masuk tabel)
 */
function kirimEmail(targetEmail, nomorSurat, jenis, sifat, derajat, hal, tanggal) {
  const subject = `Nomor Surat Anda (${nomorSurat})`;

  const sifatRaw = String(sifat || "").trim();
  const sifatLower = sifatRaw.toLowerCase();

  // Deteksi klasifikasi keamanan: SR / R / T / B
  let klasifikasi = "B";
  if (sifatLower.includes("(sr)") || sifatLower.includes("sangat rahasia")) klasifikasi = "SR";
  else if (sifatLower.includes("(r)") || sifatLower.includes("rahasia")) klasifikasi = "R";
  else if (sifatLower.includes("(t)") || sifatLower.includes("terbatas")) klasifikasi = "T";
  else if (sifatLower.includes("(b)") || sifatLower.includes("biasa") || sifatLower.includes("umum") || sifatLower.includes("terbuka")) klasifikasi = "B";

  // Alert keamanan (hanya untuk T/R/SR)
  let keamananAlertHtml = "";
  if (klasifikasi === "SR") {
    keamananAlertHtml = `
      <div style="background-color:#ffebee;color:#b71c1c;padding:12px;border-radius:6px;margin-bottom:20px;border-left:5px solid #b71c1c;font-weight:bold;">
        PERHATIAN: Naskah dinas ini berkategori SANGAT RAHASIA (SR). Akses dan distribusi harus sangat dibatasi sesuai ketentuan.
      </div>
    `;
  } else if (klasifikasi === "R") {
    keamananAlertHtml = `
      <div style="background-color:#fff3e0;color:#e65100;padding:12px;border-radius:6px;margin-bottom:20px;border-left:5px solid #e65100;font-weight:bold;">
        PERHATIAN: Naskah dinas ini berkategori RAHASIA (R). Mohon batasi akses dan distribusi sesuai ketentuan.
      </div>
    `;
  } else if (klasifikasi === "T") {
    keamananAlertHtml = `
      <div style="background-color:#fff8e1;color:#6d4c41;padding:12px;border-radius:6px;margin-bottom:20px;border-left:5px solid #ffb300;font-weight:bold;">
        PERHATIAN: Naskah dinas ini berkategori TERBATAS (T). Pastikan distribusi hanya kepada pihak yang berwenang.
      </div>
    `;
  }

  // Derajat kecepatan tampil di tabel
  const derajatText = String(derajat || "").trim() || "Biasa";

  const htmlBody = `
    <div style="font-family:'Segoe UI',Tahoma,Geneva,Verdana,sans-serif;color:#333;max-width:600px;border:1px solid #e0e0e0;padding:25px;border-radius:8px;box-shadow:0 2px 5px rgba(0,0,0,0.05);">

      <h2 style="color:#2c3e50;border-bottom:3px solid #3498db;padding-bottom:15px;margin-top:0;">
        Notifikasi Penomoran Surat
      </h2>

      ${keamananAlertHtml}

      <p style="font-size:16px;">Halo,</p>
      <p style="font-size:16px;">Permohonan nomor naskah dinas Anda telah berhasil diproses:</p>

      <table style="width:100%;border-collapse:collapse;margin:20px 0;font-size:15px;">
        <tr style="background-color:#f8f9fa;">
          <td style="padding:12px;border:1px solid #e0e0e0;font-weight:bold;width:35%;">Nomor Naskah Dinas</td>
          <td style="padding:12px;border:1px solid #e0e0e0;color:#1565c0;font-weight:bold;font-size:1.1em;">${nomorSurat}</td>
        </tr>
        <tr>
          <td style="padding:12px;border:1px solid #e0e0e0;font-weight:bold;">Jenis Naskah Dinas</td>
          <td style="padding:12px;border:1px solid #e0e0e0;">${jenis || "-"}</td>
        </tr>
        <tr>
          <td style="padding:12px;border:1px solid #e0e0e0;font-weight:bold;">Kode Keamanan Naskah Dinas</td>
          <td style="padding:12px;border:1px solid #e0e0e0;">${sifatRaw || "-"}</td>
        </tr>
        <tr>
          <td style="padding:12px;border:1px solid #e0e0e0;font-weight:bold;">Sifat Surat (Derajat Kecepatan)</td>
          <td style="padding:12px;border:1px solid #e0e0e0;font-weight:bold;">${derajatText}</td>
        </tr>
        <tr>
          <td style="padding:12px;border:1px solid #e0e0e0;font-weight:bold;">Perihal Naskah Dinas</td>
          <td style="padding:12px;border:1px solid #e0e0e0;">${hal || "-"}</td>
        </tr>
        <tr>
          <td style="padding:12px;border:1px solid #e0e0e0;font-weight:bold;">Tanggal Naskah Dinas</td>
          <td style="padding:12px;border:1px solid #e0e0e0;">${tanggal || "-"}</td>
        </tr>
      </table>

      <div style="background-color:#e3f2fd;padding:20px;border-radius:8px;border:1px solid #bbdefb;margin-top:30px;">
        <h3 style="margin-top:0;color:#0d47a1;font-size:16px;">Langkah Selanjutnya:</h3>
        <p style="margin-bottom:5px;"><strong>1. Pengajuan TTE:</strong><br>
        <a href="https://forms.gle/wwBHrNpDBoLv619P8" style="text-decoration:none;color:#1976d2;font-weight:bold;">Form Pengajuan TTE</a></p>
        <hr style="border:0;border-top:1px solid #bbdefb;margin:15px 0;">
        <p style="margin-bottom:5px;"><strong>2. Upload Arsip Surat:</strong><br>
        <a href="https://forms.gle/FTmYZuH7Ffbx2gEM8" style="text-decoration:none;color:#2e7d32;font-weight:bold;">Form Upload Surat</a></p>
      </div>

      <p style="font-size:12px;color:#757575;margin-top:30px;text-align:center;">
        Sistem Penomoran Otomatis BBWS Mesuji Sekampung &copy; ${new Date().getFullYear()}
      </p>
    </div>
  `;

  try {
    MailApp.sendEmail({ to: targetEmail, subject: subject, htmlBody: htmlBody });
  } catch (e) {
    console.error("Gagal mengirim email: " + e.message);
  }
}
