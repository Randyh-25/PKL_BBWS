// KONFIGURASI NAMA KOLOM (Sesuaikan persis dengan header di Spreadsheet Anda)
const CONFIG = {
  colEmail: "Alamat Email Aktif",       // Nama kolom email pemohon
  colStatus: "Status Permohonan",       // Nama kolom dropdown status
  colReg: "Nomor Register",             // Nama kolom Nomor Register
  colLink: "Link Surat",                // Nama kolom Link Surat (untuk yang disetujui)
  colLog: "Notifikasi Terkirim",        // Nama kolom untuk log pengiriman
  sheetName: "Status Izin Penelitian",  // Nama Sheet
  adminEmail: "ning.kurniasi@gmail.com",     
  spreadsheetLink: "https://docs.google.com/spreadsheets/d/1Z_8shV5PAdt8a-KeZ1-6hDezIiC8gK6DVNqYT5fEqy4/edit"
};

// =================================================================
// 1. TRIGGER: SAAT FORMULIR DIKIRIM (ON FORM SUBMIT)
// (Akan dipanggil oleh trigger: Dari Spreadsheet > Saat mengirim formulir)
// =================================================================
function onFormSubmitTrigger(e) {
  // Jika dieksekusi manual dari editor, hentikan
  if (!e || !e.range) return;
  
  const sheet = e.range.getSheet();
  if (sheet.getName() !== CONFIG.sheetName) return;

  const row = e.range.getRow();
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  
  const colIndexStatus = headers.indexOf(CONFIG.colStatus) + 1;
  const colIndexEmail = headers.indexOf(CONFIG.colEmail) + 1;
  const colIndexReg = headers.indexOf(CONFIG.colReg) + 1;
  const colIndexLog = headers.indexOf(CONFIG.colLog) + 1;

  const email = sheet.getRange(row, colIndexEmail).getValue();

  // --- LOGIKA PENOMORAN (Hanya buat jika kosong) ---
  let noReg = sheet.getRange(row, colIndexReg).getValue();
  if (noReg === "") {
    const today = new Date();
    const month = ("0" + (today.getMonth() + 1)).slice(-2);
    const year = today.getFullYear();
    let urutan = row - 1; // Baris 2 jadi 001
    const rowPad = ("00" + urutan).slice(-3); 
    noReg = `${rowPad}/IZP/Bbws2/${month}/${year}`;
    sheet.getRange(row, colIndexReg).setValue(noReg);
  }

  // --- SET STATUS OTOMATIS MENJADI DITERIMA ---
  sheet.getRange(row, colIndexStatus).setValue("Diterima");

  const waktuKirim = timestampIndo();

  // --- KIRIM EMAIL KE PEMOHON ---
  if (email && email.toString().includes("@")) {
    sendEmailDiterima(email, noReg);
    sheet.getRange(row, colIndexLog).setValue(`Terkirim - Diterima: ${waktuKirim}`);
  }

  // --- KIRIM NOTIFIKASI KE ADMIN ---
  notifikasiAdminOtomatis(noReg);
}

// =================================================================
// 2. TRIGGER: SAAT ADMIN EDIT SPREADSHEET (ON EDIT)
// (Akan dipanggil oleh trigger: Dari Spreadsheet > Saat diedit)
// Hanya menangani perubahan manual ke "Disetujui" atau "Ditolak"
// =================================================================
function onEditTrigger(e) {
  if (!e || !e.source) return;

  const sheet = e.source.getActiveSheet();
  const range = e.range;
  
  if (sheet.getName() !== CONFIG.sheetName) return;

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colIndexStatus = headers.indexOf(CONFIG.colStatus) + 1;
  const colIndexEmail = headers.indexOf(CONFIG.colEmail) + 1;
  const colIndexReg = headers.indexOf(CONFIG.colReg) + 1;
  const colIndexLink = headers.indexOf(CONFIG.colLink) + 1;
  const colIndexLog = headers.indexOf(CONFIG.colLog) + 1;

  // Cek jika yang diedit adalah kolom Status
  if (range.getColumn() === colIndexStatus && range.getRow() > 1) {
    const row = range.getRow();
    const status = range.getValue();
    const emailCell = sheet.getRange(row, colIndexEmail);
    const email = emailCell.getValue();
    const noReg = sheet.getRange(row, colIndexReg).getValue();
    const linkSurat = sheet.getRange(row, colIndexLink).getValue(); 
    
    // Status Diterima sudah dihandle saat Submit Form, jadi abaikan jika diedit manual ke Diterima.
    if (status === "Diterima" || status === "") return;

    if (email && email.toString().includes("@")) {
        const waktuKirim = timestampIndo();

        if (status === "Disetujui") {
          if (linkSurat === "") {
            SpreadsheetApp.getUi().alert("PERINGATAN: Link Surat belum diisi! Email TIDAK dikirim. Isi link surat terlebih dahulu lalu pilih status ulang.");
            range.setValue(""); // Kosongkan kembali statusnya
            return;
          }
          sendEmailDisetujui(email, noReg, linkSurat);
          sheet.getRange(row, colIndexLog).setValue(`Terkirim - Disetujui: ${waktuKirim}`);
        } 
        else if (status === "Ditolak") {
          sendEmailDitolak(email, noReg);
          sheet.getRange(row, colIndexLog).setValue(`Terkirim - Ditolak: ${waktuKirim}`);
        }
    } else {
       if(status !== "") {
         SpreadsheetApp.getUi().alert("Email pemohon kosong atau tidak valid.");
         range.setValue("");
       }
    }
  }
}

// =================================================================
// FUNGSI TIMESTAMP FORMAT INDONESIA
// =================================================================
function timestampIndo() {
  const now = new Date();
  const hari = ["Minggu", "Senin", "Selasa", "Rabu", "Kamis", "Jumat", "Sabtu"];
  const bulan = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];

  const namaHari = hari[now.getDay()];
  const tgl = now.getDate();
  const namaBulan = bulan[now.getMonth()];
  const thn = now.getFullYear();
  const jam = ("0" + now.getHours()).slice(-2);
  const menit = ("0" + now.getMinutes()).slice(-2);

  return `${namaHari}, ${tgl} ${namaBulan} ${thn} Pukul ${jam}:${menit} WIB`;
}

// =================================================================
// FUNGSI EMAIL KEPADA PEMOHON (Sesuai Permintaan)
// =================================================================

function sendEmailDiterima(to, noReg) {
  const subject = "Permohonan Diterima";
  const htmlBody = `
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: auto; padding: 20px; border: 1px solid #ddd; border-radius: 5px;">
      <h3 style="color: #333;">Permohonan Diterima</h3>
      <p>Yth. Bapak/Ibu Pemohon,</p>
      
      <p>Permohonan izin penelitian yang Anda ajukan telah kami terima dan dicatat dalam sistem dengan Nomor Register: <b>${noReg}</b>.</p>
      
      <p>Permohonan akan ditindaklanjuti sesuai ketentuan yang berlaku.<br>
      Apabila diperlukan informasi lebih lanjut, Bapak/Ibu dapat menghubungi Call Center BBWS Mesuji Sekampung di 0811-7215-700.</p>
      
      <p>Terima kasih atas perhatian dan kerja sama Bapak/Ibu.</p>
      
      <br>
      <p>Hormat kami,<br>
      <b>BBWS Mesuji Sekampung</b></p>
    </div>
  `;
  MailApp.sendEmail({to: to, subject: subject, htmlBody: htmlBody});
}

function sendEmailDisetujui(to, noReg, linkSurat) {
  const subject = "Notifikasi – Permohonan Disetujui";
  const htmlBody = `
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: auto; padding: 20px; border: 1px solid #ddd; border-radius: 5px; border-top: 5px solid #27ae60;">
      <h3 style="color: #27ae60;">Notifikasi – Permohonan Disetujui</h3>
      <p>Yth. Bapak/Ibu Pemohon,</p>
      
      <p>Dengan hormat kami sampaikan bahwa permohonan izin penelitian dengan Nomor Register: <b>${noReg}</b> telah disetujui.</p>
      
      <p>Persetujuan izin penelitian disampaikan melalui surat dan dapat diakses pada tautan berikut:<br>
      👉 <a href="${linkSurat}" style="color: #27ae60; font-weight: bold; text-decoration: none;">Klik Disini untuk Mengunduh Surat</a></p>
      
      <p>Pelaksanaan penelitian agar dilaksanakan sesuai lokasi, waktu, dan ketentuan yang berlaku di lingkungan BBWS Mesuji Sekampung.</p>
      
      <p>Apabila diperlukan informasi lebih lanjut, Bapak/Ibu dapat menghubungi Call Center BBWS Mesuji Sekampung di 0811-7215-700.</p>
      
      <p>Atas kerja sama Bapak/Ibu, kami ucapkan terima kasih.</p>
      
      <br>
      <p>Hormat kami,<br>
      <b>BBWS Mesuji Sekampung</b></p>
    </div>
  `;
  MailApp.sendEmail({to: to, subject: subject, htmlBody: htmlBody});
}

function sendEmailDitolak(to, noReg) {
  const subject = "Notifikasi – Permohonan Ditolak";
  const htmlBody = `
    <div style="font-family: Arial, sans-serif; max-width: 600px; margin: auto; padding: 20px; border: 1px solid #ddd; border-radius: 5px; border-top: 5px solid #c0392b;">
      <h3 style="color: #c0392b;">Notifikasi – Permohonan Ditolak</h3>
      <p>Yth. Bapak/Ibu Pemohon,</p>
      
      <p>Dengan hormat kami sampaikan bahwa permohonan izin penelitian dengan Nomor Register: <b>${noReg}</b> belum dapat kami setujui.</p>
      
      <p>Penolakan dilakukan dengan mempertimbangkan ketentuan dan kebijakan yang berlaku di lingkungan BBWS Mesuji Sekampung.<br>
      Bapak/Ibu dapat mengajukan permohonan kembali dengan menyesuaikan ketentuan yang berlaku.</p>
      
      <p>Apabila diperlukan informasi lebih lanjut, Bapak/Ibu dapat menghubungi Call Center BBWS Mesuji Sekampung di 0811-7215-700.</p>
      
      <p>Atas perhatian dan pengertiannya, kami ucapkan terima kasih.</p>
      
      <br>
      <p>Hormat kami,<br>
      <b>BBWS Mesuji Sekampung</b></p>
    </div>
  `;
  MailApp.sendEmail({to: to, subject: subject, htmlBody: htmlBody});
}

// =================================================================
// FUNGSI NOTIFIKASI KE ADMIN SAAT FORM DIISI (Diupdate)
// =================================================================
function notifikasiAdminOtomatis(noReg) {
  const subject = "🔔 Permohonan Izin Penelitian Baru Masuk!";
  const htmlBody = `
    <div style="font-family: Arial, sans-serif; padding: 20px; border: 1px solid #ddd; border-radius: 5px;">
      <h3 style="color: #2980b9;">Ada Permohonan Izin Penelitian Baru</h3>
      <p>Halo Petugas Layanan,</p>
      <p>Sistem mencatat ada pemohon baru yang telah mengisi Form Izin Penelitian dengan <b>Nomor Register: ${noReg}</b>.</p>
      <p>Sistem telah otomatis mengirim email konfirmasi "Diterima" ke pemohon dan mengatur status di Spreadsheet menjadi "Diterima".</p>
      <p>Silakan buka Spreadsheet untuk menindaklanjuti (menyetujui/menolak) permohonan tersebut:</p>
      <br>
      <a href="${CONFIG.spreadsheetLink}" style="background-color: #2980b9; color: white; padding: 10px 20px; text-decoration: none; border-radius: 5px; font-weight: bold;">Buka Spreadsheet Sekarang</a>
      <br><br>
      <p>Terima kasih,<br>Sistem Otomatis BBWS Mesuji Sekampung</p>
    </div>
  `;
  
  MailApp.sendEmail({to: CONFIG.adminEmail, subject: subject, htmlBody: htmlBody});
}