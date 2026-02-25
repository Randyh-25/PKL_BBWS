function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Sinkronisasi Data')
    .addItem('Sinkronkan Data', 'importDataDinas')
    .addSeparator()
    .addItem('Update Tabel Database (Hitung Warna)', 'extractDataKeTabel')
    .addSeparator()
    .addItem('Reset Warna & Lokasi (Sheet Aktif)', 'resetSheetAktif')
    .addToUi();
}

function resetSemuaBulan(ssRekap, bulanNames) {
  bulanNames.forEach(namaBulan => {
    const sheet = ssRekap.getSheetByName(namaBulan);
    if (sheet) {
      sheet.getRange("C9:C500").clearContent();
      sheet.getRange("F9:AJ500").clearContent().setBackground("#ffffff");
    }
  });
}

function importDataDinas() {
  const ssRekap = SpreadsheetApp.getActiveSpreadsheet();
  const sourceIds = [
    "1W2G-8-jjP8eOz-OdalQbJsgDaCILjVAjj4wmFoQ5pJE", "11JUlIawqyFi6rPQhxI2ZUWtSdaVaA0gir9ZJHpk-wwo",
    "1-AaPdHm9Hvq-pcqAkEh-dU-dIqT59S4LIs4lQGrKVJI", "1-QhSBxzRVTuox2A8P9HLVOS-Xlizfma86qyD8EuUQRQ",
    "1pxk5Z66B-aNZSNlsc-3foKksvWAOeZ4LNdnmVPKgPqI", "1xi9hxoBVepiZfcJ3eokNFliL7_7SYxCEzED28BaORMI",
    "1SStQfuBVwqU7MPlTcMCWyaiV4iA4Hi1m8rVLrccZFhc", "1L1VrAZH-fjdVUU71_aaH-RHum5amb6S4F6SvLPBeyDw",
    "1FTi_7Nq5GeUStRak3UyUBSAXqHganciWUMJ25lwg3rg", "1ZNKKeSOTEs19bts3ya6-MZa0sU-HUwj-Xa2yoU_JJs4",
    "1hmIcCjysDSspOL_vWBa2tTeiCK-b97Y06Ry-_NAyFlc", "13hGqRVJgPQfUm46EucemXT8mEUMUN1iuaR4B9dnjucU",
    "1kXGB-iUpVH1TPLh77O6Etiuni1Q_cAS0AmBYIcnDKsg"
  ];

  const bulanNames = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", 
                      "Juli", "Agustus", "September", "Oktober", "November", "Desember"];

  resetSemuaBulan(ssRekap, bulanNames);

  const sheetRef = ssRekap.getSheetByName("Januari"); 
  if (!sheetRef) {
    SpreadsheetApp.getUi().alert("Sheet Januari tidak ditemukan.");
    return;
  }
  const warnaB = sheetRef.getRange("B2:B6").getBackgrounds(); 
  const warnaC = sheetRef.getRange("C2:C6").getBackgrounds(); 
  const daftarWarna = [...warnaB.flat(), ...warnaC.flat()];

  sourceIds.forEach((id) => {
    try {
      const ssSumber = SpreadsheetApp.openById(id);
      const sheetSPD = ssSumber.getSheetByName("SPD");
      if (!sheetSPD) return;

      const data = sheetSPD.getDataRange().getValues();
      data.shift(); 

      data.forEach((row) => {
        const pelaksana      = row[2];  
        const lokasiSumber   = row[5];  
        const tglBerangkat   = parseTanggalIndo(row[7]); 
        const tglKembali     = parseTanggalIndo(row[8]); 
        const pengikutRaw    = row[14]; 

        if (!tglBerangkat || (!pelaksana && !pengikutRaw)) return;

        let daftarNamaDinas = [];
        if (pelaksana && pelaksana.toString().trim().length > 2) {
          daftarNamaDinas.push(pelaksana.toString().trim());
        }
        
        if (pengikutRaw) {
          let pengikutStr = pengikutRaw.toString().trim();
          if (pengikutStr !== "") {
            let pengikutArray = pengikutStr.split(";");
            pengikutArray.forEach(p => { 
              let n = p.trim();
              if (n !== "" && !n.toLowerCase().includes("orang") && !/^\d+$/.test(n) && n.length > 2) {
                daftarNamaDinas.push(n); 
              }
            });
          }
        }
        
        let namaUnik = [...new Set(daftarNamaDinas)];
        const namaSheetBulan = bulanNames[tglBerangkat.getMonth()];
        const sheetTujuan = ssRekap.getSheetByName(namaSheetBulan);

        if (sheetTujuan) {
          namaUnik.forEach(namaPersonil => {
            prosesRekap(sheetTujuan, namaPersonil, lokasiSumber, tglBerangkat, tglKembali, daftarWarna);
          });
        }
      });
    } catch (e) {
      Logger.log("Error ID " + id + ": " + e.message);
    }
  });

  SpreadsheetApp.getUi().alert('Sinkronisasi selesai.');
}

function prosesRekap(sheet, nama, lokasi, mulai, selesai, daftarWarna) {
  const rangeNama = sheet.getRange("B9:B500").getValues(); 
  let barisTarget = -1;
  const namaSumberBersih = bersihkanNama(nama);

  if (namaSumberBersih.length < 3) return;

  for (let i = 0; i < rangeNama.length; i++) {
    const namaRekapBersih = bersihkanNama(rangeNama[i][0]);
    if (namaRekapBersih !== "" && (namaRekapBersih.includes(namaSumberBersih) || namaSumberBersih.includes(namaRekapBersih))) {
      barisTarget = i + 9;
      break;
    }
  }

  if (barisTarget !== -1) {
    const cellLokasi = sheet.getRange(barisTarget, 3);
    let valLama = cellLokasi.getValue().toString().trim();
    let nomor = valLama === "" ? 1 : valLama.split("\n").length + 1;
    cellLokasi.setValue(valLama === "" ? "1. " + lokasi : valLama + "\n" + nomor + ". " + lokasi);

    const hariMulai = mulai.getDate();
    const hariSelesai = selesai ? selesai.getDate() : hariMulai;
    const durasi = Math.max(1, (hariSelesai - hariMulai) + 1);

    const areaWarna = sheet.getRange(barisTarget, 6, 1, 31).getBackgrounds()[0];
    let blok = 0, aktif = false;
    for (let c of areaWarna) {
      if (c !== "#ffffff" && c !== "white") { if (!aktif) { blok++; aktif = true; } } else { aktif = false; }
    }

    const rangeTgl = sheet.getRange(barisTarget, hariMulai + 5, 1, durasi);
    rangeTgl.setBackground(daftarWarna[blok % daftarWarna.length]);
    cellLokasi.setWrap(true);
  }
}

function extractDataKeTabel() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const bulanNames = ["Januari", "Februari", "Maret", "April", "Mei", "Juni", "Juli", "Agustus", "September", "Oktober", "November", "Desember"];
  
  let statsHari = {};      
  let statsFreq = {}; 
  let statsPerSatkerFreq = {}; 

  bulanNames.forEach(namaBulan => {
    const sheet = ss.getSheetByName(namaBulan);
    if (sheet) {
      const dataNama = sheet.getRange("B9:B500").getValues();
      const dataLokasi = sheet.getRange("C9:C500").getValues();
      const dataSatker = sheet.getRange("D9:D500").getValues();
      const dataWarna = sheet.getRange("F9:AJ500").getBackgrounds();

      for (let i = 0; i < dataNama.length; i++) {
        let nama = dataNama[i][0].toString().trim();
        let satker = dataSatker[i][0].toString().trim();
        let lokasiRaw = dataLokasi[i][0].toString().trim();
        
        if (nama === "") continue;

        let hari = dataWarna[i].filter(c => c !== "#ffffff" && c !== "white").length;
        let frekuensi = lokasiRaw === "" ? 0 : lokasiRaw.split("\n").length;

        statsHari[nama] = (statsHari[nama] || 0) + hari;
        statsFreq[nama] = (statsFreq[nama] || 0) + frekuensi;
        
        if (satker !== "") {
          if (!statsPerSatkerFreq[satker]) statsPerSatkerFreq[satker] = {};
          statsPerSatkerFreq[satker][nama] = (statsPerSatkerFreq[satker][nama] || 0) + frekuensi;
        }
      }
    }
  });

  let dbSheet = ss.getSheetByName("DATABASE_REKAP");
  if (!dbSheet) { dbSheet = ss.insertSheet("DATABASE_REKAP"); }
  
  dbSheet.clearContents();

  // Fungsi pembantu untuk mengambil Top 10 dan mengisi baris kosong dengan 0
  function getTop10Rows(obj) {
    let sorted = Object.entries(obj).sort((a,b) => b[1] - a[1]).slice(0, 10);
    while (sorted.length < 10) { sorted.push(["-", 0]); }
    return sorted;
  }

  let colStart = 1;

  // 1. TABEL GLOBAL HARI
  dbSheet.getRange(1, colStart).setValue("TOP 10 HARI DINAS").setFontWeight("bold");
  dbSheet.getRange(2, colStart, 1, 2).setValues([["Nama", "Total Hari"]]).setBackground("#efefef");
  dbSheet.getRange(3, colStart, 10, 2).setValues(getTop10Rows(statsHari));
  colStart += 3;

  // 2. TABEL GLOBAL FREKUENSI
  dbSheet.getRange(1, colStart).setValue("TOP 10 PERJALANAN DINAS").setFontWeight("bold");
  dbSheet.getRange(2, colStart, 1, 2).setValues([["Nama", "Total Kali"]]).setBackground("#efefef");
  dbSheet.getRange(3, colStart, 10, 2).setValues(getTop10Rows(statsFreq));
  colStart += 3;

  // 3. TABEL PER SATKER
  let listSatker = Object.keys(statsPerSatkerFreq).sort();
  listSatker.forEach(satker => {
    dbSheet.getRange(1, colStart).setValue("TOP 10: " + satker).setFontWeight("bold");
    dbSheet.getRange(2, colStart, 1, 2).setValues([["Nama", "Total Kali"]]).setBackground("#efefef");
    dbSheet.getRange(3, colStart, 10, 2).setValues(getTop10Rows(statsPerSatkerFreq[satker]));
    colStart += 3;
  });

  dbSheet.activate();
  SpreadsheetApp.getUi().alert('Tabel database diperbarui.');
}

function bersihkanNama(nama) {
  if (!nama) return "";
  const gelar = ["dr", "ir", "st", "mt", "msi", "phd", "dra", "drs", "h", "hj", "amd", "mm", "sh", "se", "skom", "kom", "ssos", "sos", "sst", "sp", "spd", "spi", "sip", "ip", "mtech", "msc", "mh", "mp"];
  let n = nama.toString().toLowerCase().replace(/\u00A0/g, " ").replace(/[,.]/g, " ");
  let kataKata = n.split(/\s+/);
  let namaInti = kataKata.filter(k => k.length > 2 && !gelar.includes(k));
  return namaInti.join("").replace(/[^a-z]/g, "");
}

function parseTanggalIndo(val) {
  if (val instanceof Date && !isNaN(val.getTime())) return val;
  if (!val) return null;
  let parts = val.toString().split("/");
  if (parts.length === 3) {
    return new Date(parseInt(parts[2]), parseInt(parts[1]) - 1, parseInt(parts[0]));
  }
  return null;
}

function resetSheetAktif() {
  const sheet = SpreadsheetApp.getActiveSheet();
  sheet.getRange("C9:C500").clearContent();
  sheet.getRange("F9:AJ500").clearContent().setBackground("#ffffff");
  SpreadsheetApp.getUi().alert('Reset selesai.');
}