// ===========================================================================
// KONFIGURASI UTAMA
// ===========================================================================
const ADMIN_EMAIL = "YOUR_ADMIN_EMAIL@example.com"; 

// ⚠️ PENTING: NAMA HEADER HARUS SAMA PERSIS 100% DENGAN DI SPREADSHEET
// Cek Kolom B dan C di Spreadsheet Anda, apakah "Email" atau "Email Address"?
const HEADERS = {
  EMAIL: "Email",        // Default Google biasanya "Email Address" (Ganti jika beda)
  NAMA: "Nama",                  // Sesuaikan dengan judul kolom Nama
  STATUS: "Status Permohonan",   
  NO_REG: "Nomor Register",      
  LINK_SURAT: "Link Surat",      
  ALASAN: "Alasan Penolakan",    
  NOTIFIKASI: "Notifikasi Terkirim" 
};

// ===========================================================================
// FUNGSI 1: SAAT FORMULIR DIKIRIM (TRIGGER: ON FORM SUBMIT)
// ===========================================================================
function onFormSubmit(e) {
  // Cek apakah e ada (untuk mencegah error jika dijalankan manual)
  if (!e) {
    console.error("❌ JANGAN JALANKAN onFormSubmit SECARA MANUAL. Gunakan Preview Form untuk mengetes.");
    return;
  }

  try {
    var sheet = e.range.getSheet();
    var row = e.range.getRow();
    
    // 1. Cari Nomor Kolom
    var colMap = getColumnMap(sheet);
    
    // DEBUG: Cek apakah kolom penting ditemukan
    if (!colMap[HEADERS.EMAIL]) console.error("❌ Kolom '" + HEADERS.EMAIL + "' TIDAK DITEMUKAN. Cek ejaan Header di Spreadsheet.");
    if (!colMap[HEADERS.NAMA]) console.error("❌ Kolom '" + HEADERS.NAMA + "' TIDAK DITEMUKAN. Cek ejaan Header di Spreadsheet.");
    
    // 2. Buat Nomor Register
    var urutan = row - 1; 
    var urutanFormat = ("000" + urutan).slice(-3);
    var today = new Date();
    var month = ("0" + (today.getMonth() + 1)).slice(-2);
    var year = today.getFullYear();
    var noRegister = urutanFormat + "/Magang/Bbws2/" + month + "/" + year;
    
    // 3. Set Status "Diterima" & No Register
    if (colMap[HEADERS.STATUS]) sheet.getRange(row, colMap[HEADERS.STATUS]).setValue("Diterima");
    if (colMap[HEADERS.NO_REG]) sheet.getRange(row, colMap[HEADERS.NO_REG]).setValue(noRegister);
    
    SpreadsheetApp.flush(); // Simpan perubahan segera

    // 4. Ambil Data Email & Nama (DENGAN PENGECEKAN)
    var emailPemohon = "";
    var namaPemohon = "";

    if (colMap[HEADERS.EMAIL]) emailPemohon = sheet.getRange(row, colMap[HEADERS.EMAIL]).getValue();
    if (colMap[HEADERS.NAMA]) namaPemohon = sheet.getRange(row, colMap[HEADERS.NAMA]).getValue();
    
    // 5. Kirim Email ke PEMOHON
    if (emailPemohon && emailPemohon.includes("@")) {
      sendEmailNotification(emailPemohon, namaPemohon, "Diterima", noRegister, "", "");
      catatWaktuNotifikasi(sheet, row, colMap[HEADERS.NOTIFIKASI]);
    } else {
      console.error("❌ Email pemohon kosong atau kolom tidak ditemukan. Notifikasi batal.");
    }

    // 6. Kirim Email ke ADMIN
    var spreadsheetUrl = SpreadsheetApp.getActiveSpreadsheet().getUrl();
    sendAdminNotification(noRegister, spreadsheetUrl);

  } catch (error) {
    console.error("❌ Error Fatal di onFormSubmit: " + error.toString());
  }
}

// ===========================================================================
// FUNGSI 2: SAAT STATUS DIEDIT (TRIGGER: ON EDIT)
// ===========================================================================
function onEditStatus(e) {
  if (!e) return; // Mencegah run manual
  
  var range = e.range;
  var sheet = range.getSheet();
  var row = range.getRow();
  var col = range.getColumn();
  
  var colMap = getColumnMap(sheet);
  var statusColIndex = colMap[HEADERS.STATUS];

  // Pastikan kolom status ditemukan sebelum lanjut
  if (!statusColIndex) return;

  if (col === statusColIndex && row > 1) {
    var status = range.getValue();
    
    // Ambil data dengan pengecekan aman
    var email = colMap[HEADERS.EMAIL] ? sheet.getRange(row, colMap[HEADERS.EMAIL]).getValue() : "";
    var nama = colMap[HEADERS.NAMA] ? sheet.getRange(row, colMap[HEADERS.NAMA]).getValue() : "";
    var noReg = colMap[HEADERS.NO_REG] ? sheet.getRange(row, colMap[HEADERS.NO_REG]).getValue() : "";
    var linkSurat = colMap[HEADERS.LINK_SURAT] ? sheet.getRange(row, colMap[HEADERS.LINK_SURAT]).getValue() : "";
    var alasan = colMap[HEADERS.ALASAN] ? sheet.getRange(row, colMap[HEADERS.ALASAN]).getValue() : "";
    
    // VALIDASI "Disetujui"
    if (status === "Disetujui") {
      if (!linkSurat) {
        SpreadsheetApp.getUi().alert("⛔ PERHATIAN: Kolom 'Link Surat' kosong! Email TIDAK dikirim.");
        range.setValue(""); 
        return; 
      }
    }
    
    // KIRIM EMAIL
    if (status === "Disetujui" || status === "Ditolak" || status === "Isi Survei IKM dan IPK") {
      if (email && email.includes("@")) {
        sendEmailNotification(email, nama, status, noReg, linkSurat, alasan);
        catatWaktuNotifikasi(sheet, row, colMap[HEADERS.NOTIFIKASI]);
        SpreadsheetApp.getActiveSpreadsheet().toast("Email terkirim ke " + email, "Sukses");
      } else {
         SpreadsheetApp.getActiveSpreadsheet().toast("Gagal: Email tidak valid/kolom tidak ditemukan", "Error");
      }
    }
  }
}

// ===========================================================================
// HELPER FUNCTIONS
// ===========================================================================

function getColumnMap(sheet) {
  var lastCol = sheet.getLastColumn();
  if (lastCol < 1) return {};
  
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]; 
  var map = {};
  for (var i = 0; i < headers.length; i++) {
    var headerName = headers[i].toString().trim(); // Hapus spasi berlebih
    map[headerName] = i + 1;
  }
  return map;
}

function catatWaktuNotifikasi(sheet, row, colIndex) {
  if (colIndex) {
    var now = new Date();
    var formattedDate = Utilities.formatDate(now, "Asia/Jakarta", "HH:mm dd MMMM yyyy");
    sheet.getRange(row, colIndex).setValue(formattedDate);
    SpreadsheetApp.flush(); 
  }
}

function sendAdminNotification(noReg, sheetUrl) {
  if (!ADMIN_EMAIL) return; 

  var subject = "🔔 Permohonan Magang Mahasiswa Baru Masuk!";
  var htmlBody = `
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: auto; border: 1px solid #e0e0e0; border-radius: 8px; overflow: hidden;">
      <div style="padding: 20px;">
        <h3 style="color: #1a73e8; margin-top: 0;">Ada Permohonan Magang Mahasiswa Baru</h3>
        <p>Halo Petugas Layanan,</p>
        <p>Sistem mencatat pemohon baru dengan <b>Nomor Register: ${noReg}</b>.</p>
        <p>Sistem telah otomatis mengirim email konfirmasi "Diterima" ke pemohon dan mengatur status di Spreadsheet menjadi "Diterima".</p>
        <p>Silakan buka Spreadsheet untuk menindaklanjuti:</p>
        <div style="margin: 25px 0;">
          <a href="${sheetUrl}" style="background-color: #1565C0; color: white; padding: 12px 24px; text-decoration: none; border-radius: 4px; font-weight: bold;">Buka Spreadsheet Sekarang</a>
        </div>
      </div>
    </div>
  `;

  MailApp.sendEmail({to: ADMIN_EMAIL, subject: subject, htmlBody: htmlBody});
}

function sendEmailNotification(email, nama, status, noReg, linkSurat, alasan) {
  var subject = "";
  var bodyContent = "";
  var headerColor = "#E5BA41"; 

  if (status === "Diterima") {
    subject = "Layanan Magang Mahasiswa - Permohonan Diterima";
    bodyContent = `
      <p>Yth. Saudara/i <b>${nama}</b>,</p>
      <p>Permohonan magang mahasiswa yang Anda ajukan telah kami terima dan dicatat sistem dengan Nomor Register: <b>${noReg}</b>.</p>
      <p>Permohonan akan segera kami tindaklanjuti sesuai ketentuan yang berlaku.<br>
      Apabila diperlukan informasi lebih lanjut, Saudara/i dapat menghubungi Call Center BBWS Mesuji Sekampung di [NOMOR_CALL_CENTER].</p>
      <p>Terima kasih atas perhatian dan kerja sama Saudara/i.</p>
      <br><p>Hormat kami,<br><b>BBWS Mesuji Sekampung</b></p>
    `;
  } 
  else if (status === "Disetujui") {
    var noIzin = noReg.replace("Magang", "IZP"); 
    subject = "Layanan Magang Mahasiswa - Permohonan Disetujui";
    bodyContent = `
      <p>Yth. Saudara/i <b>${nama}</b>,</p>
      <p>Dengan hormat kami sampaikan bahwa permohonan Magang Mahasiswa dengan Nomor Register: <b>${noIzin}</b> telah disetujui.</p>
      <p>Persetujuan Magang Mahasiswa disampaikan melalui surat dan dapat diakses pada tautan berikut:</p>
      <div style="text-align: center; margin: 20px 0;">
        <a href="${linkSurat}" style="background-color: ${headerColor}; color: white; padding: 12px 25px; text-decoration: none; border-radius: 5px; font-weight: bold;">👉 LIHAT SURAT PERSETUJUAN</a>
      </div>
      <p>Pelaksanaan magang mahasiswa agar dilaksanakan sesuai lokasi, waktu, dan ketentuan yang berlaku di lingkungan BBWS Mesuji Sekampung.</p>
      <p>Apabila diperlukan informasi lebih lanjut, Saudara/i dapat menghubungi Call Center BBWS Mesuji Sekampung di [NOMOR_CALL_CENTER].</p>
      <p>Atas kerja sama Suadara/i, kami ucapkan terima kasih.</p>
      <br><p>Hormat kami,<br><b>BBWS Mesuji Sekampung</b></p>
    `;
  } 
  else if (status === "Ditolak") {
    subject = "Layanan Magang Mahasiswa - Permohonan Ditolak";
    var alasanText = alasan ? `<p><strong>Alasan:</strong> ${alasan}</p>` : "";
    bodyContent = `
      <p>Yth. Saudara/i <b>${nama}</b>,</p>
      <p>Dengan hormat kami sampaikan bahwa permohonan Magang Mahasiswa dengan Nomor Register: <b>${noReg}</b> belum dapat kami setujui.</p>
      <div style="background-color: #ffebee; padding: 15px; border-left: 5px solid #d32f2f; margin: 10px 0;">
        ${alasanText}
      </div>
      <p>Penolakan dilakukan dengan mempertimbangkan ketentuan dan kebijakan yang berlaku di lingkungan BBWS Mesuji Sekampung.<br>
      Saudara/1 dapat mengajukan permohonan kembali dengan menyesuaikan ketentuan yang berlaku.</p>
      <p>Apabila diperlukan informasi lebih lanjut, Saudara/i dapat menghubungi Call Center BBWS Mesuji Sekampung di [NOMOR_CALL_CENTER].</p>
      <p>Atas perhatian dan pengertiannya, kami ucapkan terima kasih.</p>
      <br><p>Hormat kami,<br><b>BBWS Mesuji Sekampung</b></p>
    `;
  }
  else if (status === "Isi Survei IKM dan IPK") {
    subject = "Permohonan Pengisian Survei (IKM & IPK)";
    bodyContent = `
      <p>Yth. Saudara/i <b>${nama}</b>,</p>
      <p align="justify">Terima kasih telah menggunakan layanan Magang Mahasiswa Balai Besar Wilayah Sungai Mesuji Sekampung. Sehubungan dengan telah selesainya proses permohonan magang mahasiswa Saudara/i dengan Nomor Registrasi: <b>${noReg}</b>, kami mohon kesediaan Saudara/i untuk mengisi Survei Indeks Kepuasan Masyarakat (IKM) dan Indeks Persepsi Korupsi (IPK) sebagai bahan evaluasi dan peningkatan kualitas layanan kami.</p>
      <p>Mohon kesediaan Saudara/i untuk mengisi Survei IKM & IPK pada tautan berikut:</p>
      <div style="text-align: center; margin: 20px 0;">
        <a href="[LINK_SURVEI_IKM_IPK]" style="background-color: #2E7D32; color: white; padding: 12px 25px; text-decoration: none; border-radius: 5px; font-weight: bold;">📝 ISI SURVEI</a>
      </div>

      <p align="justify">Pengisian survei bersifat sukarela, tidak dipungut biaya, dan tidak mempengaruhi layanan yang telah atau akan diterima. Seluruh jawaban Saudara/i dijamin kerahasiaannya.</p>
      
      <p align="justify">Partisipasi Saudara/i sangat berarti bagi peningkatan transparansi, akuntabilitas, dan kualitas pelayanan informasi publik di lingkungan BBWS Mesuji Sekampung.</p>

      <p>Atas perhatian dan kerja sama Saudara/i, kami ucapkan terima kasih.</p>
      <br>
      <p style="font-size: 13px; color: #555;">Hormat kami,<br>
      <b>Layanan Magang Mahasiswa<br>
      Balai Besar Wilayah Sungai Mesuji Sekampung</b><br>
      Call Center PPID: [NOMOR_CALL_CENTER]</p>
    `;
  }

  var htmlTemplate = `
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: auto; border: 1px solid #e0e0e0; border-radius: 8px;">
      <div style="background-color: ${headerColor}; padding: 20px; color: white; text-align: center;">
        <h2 style="margin:0;">BBWS Mesuji Sekampung</h2>
      </div>
      <div style="padding: 20px; color: #333;">
        ${bodyContent}
        <br><hr><p style="font-size: 12px; color: #777;">Sistem Otomatis BBWS Mesuji Sekampung</p>
      </div>
    </div>
  `;

  if (subject && email) {
    MailApp.sendEmail({to: email, subject: subject, htmlBody: htmlTemplate});
  }
}