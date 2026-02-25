/**
 * SISTEM NOTIFIKASI PPID BBWS MESUJI SEKAMPUNG
 * Versi: Final - Fixed Mapping Kolom E & F + Notifikasi Admin
 */

// --- 1. KONFIGURASI NAMA KOLOM (WAJIB SESUAIKAN DENGAN HEADER DI SPREADSHEET) ---

// A. Header Bawaan Google Form
const HEADER_EMAIL = "Email"; 
const HEADER_NAMA = "Nama Pemohon"; 
const HEADER_JENIS = "Rincian Informasi yang Dibutuhkan"; 
const HEADER_NO_KTP = "Nomor KTP"; 
const HEADER_ALAMAT = "Alamat Tempat Tinggal"; 
const HEADER_HP = "No. Telp / HP"; 

// B. Header Kolom Admin (Diisi Manual oleh Petugas)
const KOLOM_NO_FORMULIR = "Nomor Formulir";
const KOLOM_STATUS = "Status Verifikasi"; 
const KOLOM_LINK_DOKUMEN = "Link Dokumen Pemohon"; 
// Catatan: Variabel ini dipakai untuk status "Selesai" & "Tidak Lengkap"
const KOLOM_CATATAN = "Keterangan Tambahan"; 

// C. Header Kolom Rincian (Untuk Status Diproses - Pemberitahuan Tertulis)
const KOLOM_STATUS_INFO = "Status Informasi"; 
const KOLOM_PENGUASAAN = "Penguasaan Informasi"; 
const KOLOM_BENTUK = "Bentuk Informasi Tersedia"; 
const KOLOM_BIAYA = "Biaya"; 

// [UPDATE] Sesuaikan nama kolom di sini agar Poin E & F terbaca
const KOLOM_WAKTU = "Waktu Penyediaan Informasi"; // Sesuaikan dgn nama kolom baru
const KOLOM_KET_TAMBAHAN = "Keterangan Tambahan"; // Sesuaikan dgn nama kolom baru

// D. Data Petugas
const NAMA_PETUGAS = "Nindy Widyawati";
const CONTACT_CENTER = "08117215700";
const EMAIL_ADMIN = "nindindoy19@gmail.com"; // Email Admin untuk Notifikasi


// --- FUNGSI 1: SAAT FORM DISUBMIT (Otomatis Kirim Tanda Terima & Notif Admin) ---
function onFormSubmit(e) {
  try {
    const sheet = e.range.getSheet();
    const row = e.range.getRow();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const responses = e.namedValues;

    console.log("New Form Submit at Row: " + row);

    // 1. Set Status Awal
    updateCell(sheet, row, headers, KOLOM_STATUS, "Permohonan Diterima");

    // 2. Generate Nomor Formulir
    const date = new Date();
    const noForm = `${row - 1}/PPID/Bbws2/${date.getMonth() + 1}/${date.getFullYear()}`;
    updateCell(sheet, row, headers, KOLOM_NO_FORMULIR, noForm);

    // 3. Ambil Data
    const email = responses[HEADER_EMAIL] ? responses[HEADER_EMAIL][0] : "";
    const nama = responses[HEADER_NAMA] ? responses[HEADER_NAMA][0] : "-";
    const jenis = responses[HEADER_JENIS] ? responses[HEADER_JENIS][0] : "-";
    const ktp = responses[HEADER_NO_KTP] ? responses[HEADER_NO_KTP][0] : "-";
    const alamat = responses[HEADER_ALAMAT] ? responses[HEADER_ALAMAT][0] : "-";
    // Tambahan: Ambil No HP juga untuk info Admin
    const hp = responses[HEADER_HP] ? responses[HEADER_HP][0] : "-";

    if (!email) return;

    // 4. Kirim Email Tanda Terima ke PEMOHON
    const htmlBodyPemohon = `
      <div style="font-family: 'Times New Roman', serif; border: 1px solid #ccc; padding: 25px; max-width: 600px;">
        <h3 style="text-align:center; text-decoration: underline;">TANDA BUKTI PENERIMAAN PERMOHONAN</h3>
        <p>Permohonan Informasi Publik Anda telah diterima sistem dengan rincian:</p>
        <table style="width:100%">
          <tr><td width="35%">No. Registrasi</td><td>: <strong>${noForm}</strong></td></tr>
          <tr><td>Nama Pemohon</td><td>: ${nama}</td></tr>
          <tr><td>No. KTP</td><td>: ${ktp}</td></tr>
          <tr><td>Alamat</td><td>: ${alamat}</td></tr>
          <tr><td>Informasi diminta</td><td>: ${jenis}</td></tr>
        </table>
        <br>
        <p>Bandar Lampung, ${formatTanggalIndo(date)}</p>
        <p>Petugas PPID, <br><br><strong>${NAMA_PETUGAS}</strong></p>
      </div>
    `;

    MailApp.sendEmail({to: email, subject: "Tanda Bukti Penerimaan Permohonan Informasi Publik", htmlBody: htmlBodyPemohon});

    // 5. Kirim Notifikasi ke ADMIN (NEW)
    const subjectAdmin = `[ADMIN PPID] Permohonan Baru: ${nama}`;
    const htmlBodyAdmin = `
      <div style="font-family: Arial, sans-serif; padding: 15px; border: 1px solid #333; max-width: 600px;">
        <h3 style="margin-top:0;">🔔 Notifikasi Permohonan Masuk</h3>
        <p>Halo Admin, terdapat data permohonan informasi baru yang perlu diverifikasi:</p>
        <table style="width:100%; border-collapse: collapse;">
          <tr><td width="30%"><strong>No. Registrasi</strong></td><td>: ${noForm}</td></tr>
          <tr><td><strong>Nama</strong></td><td>: ${nama}</td></tr>
          <tr><td><strong>No. HP</strong></td><td>: ${hp}</td></tr>
          <tr><td><strong>Informasi</strong></td><td>: ${jenis}</td></tr>
          <tr><td><strong>Email</strong></td><td>: ${email}</td></tr>
        </table>
        <p>Silakan buka Spreadsheet untuk memverifikasi berkas.</p>
        <p><small>Waktu terima: ${formatTanggalIndo(date)}</small></p>
      </div>
    `;
    
    MailApp.sendEmail({to: EMAIL_ADMIN, subject: subjectAdmin, htmlBody: htmlBodyAdmin});

  } catch (err) { console.error(err.toString()); }
}


// --- FUNGSI 2: SAAT EDIT STATUS MANUAL ---
function onEditStatus(e) {
  try {
    const range = e.range;
    const sheet = range.getSheet();
    const row = range.getRow();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Validasi Kolom Status
    const idxStatus = headers.indexOf(KOLOM_STATUS);
    if (range.getColumn() !== idxStatus + 1 || row <= 1) return;

    const newValue = e.value; 
    
    // Helper Ambil Data
    const getVal = (colName) => {
      const idx = headers.indexOf(colName);
      return (idx > -1) ? sheet.getRange(row, idx + 1).getValue() : "-";
    };

    const email = getVal(HEADER_EMAIL);
    const nama = getVal(HEADER_NAMA);
    const noReg = getVal(KOLOM_NO_FORMULIR);
    const jenisInfo = getVal(HEADER_JENIS); // Uraian Informasi
    
    // Ambil Tanggal Permohonan (Kolom 1 / Timestamp)
    const tglRaw = sheet.getRange(row, 1).getValue();
    const tglMohon = (tglRaw instanceof Date) ? formatTanggalIndo(tglRaw) : "-";

    if (!email || email === "-" || !email.includes("@")) return;

    let subject = "";
    let htmlContent = "";

    // --- LOGIKA PER STATUS ---

    // 1. DIVERIFIKASI
    if (newValue == "Permohonan Diverifikasi - Berkas Lengkap") {
      subject = "Pemberitahuan: Permohonan Dinyatakan Lengkap";
      htmlContent = `
        <p>Yth. Saudara/i <strong>${nama}</strong>,</p>
        <p>Permohonan dengan No. Registrasi <strong>${noReg}</strong> telah diverifikasi dan dinyatakan <strong>LENGKAP</strong>.</p>
        <p>Selanjutnya permohonan akan diproses sesuai ketentuan.</p>
      `;
    }

    // 2. BERKAS TIDAK LENGKAP
    else if (newValue == "Permohonan Diverifikasi - Berkas Tidak Lengkap") {
      const ketTambahan = getVal(KOLOM_CATATAN); // Ambil dari Keterangan Tambahan
      
      subject = "Pemberitahuan: Permohonan Belum Lengkap";
      htmlContent = `
        <div style="font-family: Arial, sans-serif; line-height: 1.5; color: #000;">
          <p>Yth. Saudara/i <strong>${nama}</strong>,</p>
          <p>Permohonan informasi publik Anda dengan Nomor Registrasi: <strong>${noReg}</strong> telah diverifikasi oleh PPID BBWS Mesuji Sekampung.</p>
          <p>Berdasarkan hasil verifikasi, permohonan Anda belum lengkap dan memerlukan perbaikan sesuai dengan catatan verifikasi sebagai berikut:</p>
          <p style="background-color: #fff3cd; padding: 10px; border: 1px solid #ffeeba;"><strong>${ketTambahan}</strong></p>
          <p>Permohonan informasi publik dengan nomor registrasi ini dinyatakan tidak diproses lebih lanjut.</p>
          <p>Pemohon dipersilakan mengajukan permohonan informasi publik baru dengan melampirkan kelengkapan sesuai catatan di atas.</p>
        </div>
      `;
    }

    // 3. DIPROSES (PEMBERITAHUAN TERTULIS)
    else if (newValue == "Permohonan Diproses") {
      const statusInfo = getVal(KOLOM_STATUS_INFO);
      const penguasaan = getVal(KOLOM_PENGUASAAN);
      const bentuk = getVal(KOLOM_BENTUK);
      const biaya = getVal(KOLOM_BIAYA);
      const waktu = getVal(KOLOM_WAKTU);        // Mengambil dari kolom "Waktu Penyediaan Informasi"
      const ket = getVal(KOLOM_KET_TAMBAHAN);   // Mengambil dari kolom "Keterangan Tambahan"

      subject = "Pemberitahuan Tertulis PPID BBWS Mesuji Sekampung";
      htmlContent = `
        <div style="font-family: Arial, sans-serif; line-height: 1.5; color: #000;">
          <p>Yth. Saudara/i <strong>${nama}</strong>,</p>
          <p>Menindaklanjuti permohonan informasi publik yang Saudara/i ajukan, dengan ini kami sampaikan <strong>Pemberitahuan Tertulis</strong>:</p>
          
          <table style="width:100%; border-collapse: collapse; margin-bottom:15px;">
            <tr><td width="30%">No. Registrasi</td><td>: <strong>${noReg}</strong></td></tr>
            <tr><td>Tanggal</td><td>: ${tglMohon}</td></tr>
          </table>

          <div style="background-color: #f8f9fa; padding: 15px; border: 1px solid #dee2e6;">
            <p><strong>A. Status Permohonan</strong><br>Status: <span style="font-weight: bold;">${statusInfo}</span></p>
            <p><strong>B. Penguasaan Informasi</strong><br>${penguasaan}</p>
            <p><strong>C. Bentuk Informasi yang Tersedia</strong><br>${bentuk}</p>
            <p><strong>D. Biaya yang Dibutuhkan</strong><br>${biaya}</p>
            <p><strong>E. Waktu Penyediaan Informasi</strong><br>${waktu}</p>
            <p><strong>F. Keterangan Tambahan</strong><br>${ket}</p>
          </div>
           <p style="text-align: justify; margin-top: 15px; font-size: 0.9em;">
            Pemberitahuan tertulis ini disampaikan sebagai jawaban resmi PPID sesuai UU No. 14 Tahun 2008.
            Apabila Saudara/i tidak puas, Saudara/i berhak mengajukan keberatan melalui tautan: 
            <a href="https://s.pu.go.id/MTMwOA/Form_Keberatan_PPID">Form Keberatan PPID</a>
          </p>
        </div>`;
    }

    // 4. SELESAI
    else if (newValue == "Permohonan Selesai") {
      const alamat = getVal(HEADER_ALAMAT);
      const hp = getVal(HEADER_HP);
      const linkDok = getVal(KOLOM_LINK_DOKUMEN);
      const bentukSelesai = getVal(KOLOM_BENTUK); 
      
      const now = new Date();
      const tglSelesai = formatTanggalIndo(now);
      const jamSelesai = now.toLocaleTimeString('id-ID', { hour: '2-digit', minute: '2-digit' }).replace('.', ':');

      subject = "Penyerahan Informasi Publik - PPID BBWS Mesuji Sekampung";
      
      htmlContent = `
        <div style="font-family: Arial, sans-serif; line-height: 1.6; color: #000;">
          <p>Yth. <strong>${nama}</strong>,</p>
          <p>Sehubungan dengan permohonan informasi publik yang Saudara/i ajukan melalui layanan PPID Balai Besar Wilayah Sungai Mesuji Sekampung, dengan ini kami sampaikan bahwa penyerahan informasi publik telah dilaksanakan, dengan rincian sebagai berikut:</p>
          
          <h4 style="margin-bottom: 5px;">A. Data Penyerahan Informasi</h4>
          <table style="width: 100%; border-collapse: collapse;">
             <tr><td width="35%" style="vertical-align: top;">• Nomor Registrasi Permohonan</td><td width="2%">:</td><td><strong>${noReg}</strong></td></tr>
             <tr><td style="vertical-align: top;">• Informasi yang Diberikan</td><td>:</td><td>${jenisInfo}</td></tr>
             <tr><td style="vertical-align: top;">• Format Informasi</td><td>:</td><td>☑ Terekam / Softcopy (${bentukSelesai})</td></tr>
             <tr><td style="vertical-align: top;">• Cara Penyerahan Informasi</td><td>:</td><td>☑ Email (tautan Google Drive)</td></tr>
             <tr><td style="vertical-align: top;">• Tautan Akses Informasi Publik</td><td>:</td><td><a href="${linkDok}">${linkDok}</a></td></tr>
          </table>
          <p style="font-size: 0.9em; font-style: italic; background: #eee; padding: 5px;">Catatan: Tautan Google Drive hanya dapat diakses menggunakan alamat email yang didaftarkan saat registrasi permohonan informasi publik.</p>

          <h4 style="margin-bottom: 5px;">B. Identitas Pemohon Informasi</h4>
          <table style="width: 100%; border-collapse: collapse;">
             <tr><td width="35%">• Nama</td><td width="2%">:</td><td>${nama}</td></tr>
             <tr><td>• Alamat</td><td>:</td><td>${alamat}</td></tr>
             <tr><td>• No. HP</td><td>:</td><td>${hp}</td></tr>
             <tr><td>• Email Terdaftar</td><td>:</td><td>${email}</td></tr>
          </table>

          <h4 style="margin-bottom: 5px;">C. Waktu Penyerahan</h4>
          <table style="width: 100%; border-collapse: collapse;">
             <tr><td width="35%">• Tanggal</td><td width="2%">:</td><td>${tglSelesai}</td></tr>
             <tr><td>• Pukul</td><td>:</td><td>${jamSelesai} WIB</td></tr>
          </table>

          <h4 style="margin-bottom: 5px;">D. Konfirmasi Pemenuhan Informasi</h4>
          <p style="text-align: justify;">Kami mohon Saudara/i memeriksa data/informasi yang telah disampaikan. Konfirmasi pemenuhan informasi disampaikan melalui Call Center PPID BBWS Mesuji Sekampung: 📞 <strong>${CONTACT_CENTER}</strong>. <br⏱️ Batas waktu konfirmasi: <strong>1 x 24 jam</strong> sejak notifikasi ini diterima. <br>Apabila dalam jangka waktu tersebut tidak terdapat konfirmasi, maka informasi publik dianggap telah diterima dan sesuai, serta proses permohonan informasi publik dinyatakan selesai.</p>

          <h4 style="margin-bottom: 5px;">E. Keterangan</h4>
          <ol>
            <li>Notifikasi ini merupakan tanda bukti sah penyerahan informasi publik secara elektronik.</li>
            <li>Penyerahan informasi dilakukan sesuai dengan Undang-Undang Nomor 14 Tahun 2008 tentang Keterbukaan Informasi Publik.</li>
            <li>Hak-hak Pemohon Informasi Publik dapat dibaca melalui tautan berikut: <a href="https://s.pu.go.id/MTMwOA/Hak_Pemohon_Informasi">Link Hak Pemohon Informasi</a></li>
          </ol>

          <h4 style="margin-bottom: 5px;">F. Petugas PPID</h4>
          <p>Petugas yang Menyerahkan Informasi<br>PPID BBWS Mesuji Sekampung<br><strong>${NAMA_PETUGAS}</strong></p>
        </div>
      `;
    }

    // 5. ISI SURVEY IKM DAN IPK
    else if (newValue == "Isi Survei IKM dan IPK") {
      subject = "Permohonan Pengisian Survei Kepuasan Masyarakat (IKM & IPK)";
      htmlContent = `
        <div style="font-family: Arial, sans-serif; line-height: 1.6; color: #000;">
          <p>Yth. Saudara/i <strong>${nama}</strong>,</p>
          <p>Terima kasih telah menggunakan layanan Informasi Publik PPID Balai Besar Wilayah Sungai Mesuji Sekampung. Sehubungan dengan telah selesainya proses permohonan informasi publik Saudara/i dengan Nomor Registrasi: <strong>${noReg}</strong>, kami mohon kesediaan Saudara/i untuk mengisi Survei Indeks Kepuasan Masyarakat (IKM) dan Indeks Persepsi Korupsi (IPK) sebagai bahan evaluasi dan peningkatan kualitas layanan kami.</p>

          <p>📝 <strong>Tautan Survei IKM & IPK:</strong><br>
          🔗 <a href="https://s.pu.go.id/MTMwOA/Survei_IKMdanIPK_BBWSMS">https://s.pu.go.id/MTMwOA/Survei_IKMdanIPK_BBWSMS</a></p>

          <p>Pengisian survei bersifat sukarela, tidak dipungut biaya, dan tidak mempengaruhi layanan yang telah atau akan diterima. Seluruh jawaban Saudara/i dijamin kerahasiaannya.</p>
          <p>Partisipasi Saudara/i sangat berarti bagi peningkatan transparansi, akuntabilitas, dan kualitas pelayanan informasi publik di lingkungan BBWS Mesuji Sekampung.</p>
          <p>Atas perhatian dan kerja sama Saudara/i, kami ucapkan terima kasih.</p>
        </div>
      `;
    }

    // --- KIRIM EMAIL ---
    if (subject && email) {
      const footer = `
        <br><hr>
        <div style="font-size: 0.9em; color: #555;">
          <p>Hormat kami,<br><strong>Pejabat Pengelola Informasi dan Dokumentasi (PPID)</strong><br>Balai Besar Wilayah Sungai Mesuji Sekampung</p>
          <p><em>Call Center PPID: ${CONTACT_CENTER}</em></p>
        </div>
      `;
      MailApp.sendEmail({to: email, subject: subject, htmlBody: htmlContent + footer});
    }

  } catch (err) { console.error(err.toString()); }
}


// --- HELPER FUNCTIONS ---
function updateCell(sheet, row, headers, colName, value) {
  const idx = headers.indexOf(colName);
  if (idx > -1) sheet.getRange(row, idx + 1).setValue(value);
}

function formatTanggalIndo(dateObj) {
  const bulan = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  return `${dateObj.getDate()} ${bulan[dateObj.getMonth()]} ${dateObj.getFullYear()}`;
}