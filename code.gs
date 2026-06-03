/**
 * Core.gs - Router & Helper Utama
 */
const SPREADSHEET_ID = SpreadsheetApp.getActiveSpreadsheet().getId();

function doGet(e) {
  // WAJIB Menggunakan createTemplateFromFile dan evaluate().
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('LMS')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Helper: Menyisipkan file HTML modular
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// Helper: Ambil Data Range
function getSheetData_(sheetName) {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const sheet = ss.getSheetByName(sheetName);
  return sheet ? sheet.getDataRange().getValues() : [];
}

// Helper: Format Timestamp
function getTimestamp_() {
  return Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
}

// ====================================================================
// Fungsi trackVisitor() untuk Mencatat Pengunjung Unik
// ====================================================================
function trackVisitor(sessionId) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName('Log_Kunjungan');
    
    // Jika sheet Log_Kunjungan belum ada, buat baru
    if (!sheet) {
      sheet = ss.insertSheet('Log_Kunjungan');
      sheet.appendRow(['Tanggal', 'Session_ID', 'Timestamp']);
    }
    
    const tglHariIni = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'yyyy-MM-dd');
    const timestamp = getTimestamp_();
    const activeSession = sessionId || 'anonymous';
    
    // Periksa apakah session ini sudah dicatat hari ini
    const data = sheet.getDataRange().getValues();
    let exists = false;
    for (let i = 1; i < data.length; i++) {
      let cellTgl = data[i][0];
      // Jika Sheets terlanjur membaca data lama sebagai format Date, ubah kembali ke String
      if (cellTgl instanceof Date) {
          cellTgl = Utilities.formatDate(cellTgl, 'Asia/Jakarta', 'yyyy-MM-dd');
      } else {
          cellTgl = cellTgl.toString().trim();
      }
      
      if (cellTgl === tglHariIni && data[i][1] === activeSession) {
        exists = true;
        break;
      }
    }
    
    // Jika belum dicatat hari ini, tambahkan baris baru
    // REVISI: Menyimpan tanggal dan timestamp dengan penambahan tanda kutip di depan agar dibaca utuh sebagai Teks
    if (!exists) {
      sheet.appendRow(["'" + tglHariIni, activeSession, "'" + timestamp]);
    }
    
    return { status: 'success' };
  } catch (error) {
    Logger.log("Error di trackVisitor: " + error.message);
    return { status: 'error', message: error.message };
  }
}

// ====================================================================
// Fungsi getAdminStats() untuk Real-time Dashboard (Updated with Visitor & Chart)
// ====================================================================
function getAdminStats() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Ambil Data Siswa & Hitung Kelas Unik
    const sheetSiswa = ss.getSheetByName('Data_Siswa');
    let totalSiswa = 0;
    let totalKelas = 0;
    
    if (sheetSiswa) {
      const dataSiswa = sheetSiswa.getDataRange().getValues();
      if (dataSiswa.length > 1) {
        // Pemetaan Index dinamis agar tidak error jika kolom digeser
        const headers = dataSiswa[0].map(h => h.toString().toLowerCase().trim());
        const idxStatus = headers.indexOf('status');
        const idxKelas = headers.indexOf('kelas');
        
        const kelasSet = new Set();
        
        for (let i = 1; i < dataSiswa.length; i++) {
          const row = dataSiswa[i];
          const status = idxStatus !== -1 && row[idxStatus] ? row[idxStatus].toString().toLowerCase() : '';
          
          // Hitung siswa jika Status = 'aktif' atau kolom Status tidak ditemukan (bypass)
          if (status === 'aktif' || status === '' || idxStatus === -1) {
             totalSiswa++;
          }
          
          // Tampung nama kelas unik
          if (idxKelas !== -1 && row[idxKelas] && row[idxKelas].toString().trim() !== '') {
             kelasSet.add(row[idxKelas].toString().trim());
          }
        }
        totalKelas = kelasSet.size;
      }
    }

    // 2. Ambil Data Guru (Dari Sheet Pengguna)
    const sheetPengguna = ss.getSheetByName('Pengguna');
    let totalGuru = 0;
    
    if (sheetPengguna) {
      const dataPengguna = sheetPengguna.getDataRange().getValues();
      if(dataPengguna.length > 1) {
        const headers = dataPengguna[0].map(h => h.toString().toLowerCase().trim());
        const idxRole = headers.indexOf('role');
        
        for (let i = 1; i < dataPengguna.length; i++) {
          if (idxRole !== -1 && dataPengguna[i][idxRole].toString().toLowerCase() === 'guru') {
            totalGuru++;
          }
        }
      }
    }

    // 3. Ambil Data Modul LMS / Tugas (Fallback opsional jika belum ada sheetnya)
    let totalLMS = 0;
    const sheetLMS = ss.getSheetByName('Materi_LMS'); 
    if(sheetLMS) {
      totalLMS = Math.max(0, sheetLMS.getLastRow() - 1); // Hitung total baris dikurangi Header
    }

    // ==========================================
    // REVISI: Hitung Jumlah Visitor dan Buat Grafik Harian
    // ==========================================
    let totalVisitor = 0;
    let chartLabels = [];
    let chartValues = [];
    const logSheet = ss.getSheetByName('Log_Kunjungan');
    
    // Generate label 7 hari terakhir secara dinamis
    const MS_PER_DAY = 24 * 60 * 60 * 1000;
    const hariIni = new Date();
    const listTanggalTerakhir = [];
    
    for (let i = 6; i >= 0; i--) {
      const d = new Date(hariIni.getTime() - (i * MS_PER_DAY));
      const formattedDate = Utilities.formatDate(d, 'Asia/Jakarta', 'yyyy-MM-dd');
      listTanggalTerakhir.push(formattedDate);
    }

    const kunjunganMap = {};
    listTanggalTerakhir.forEach(tgl => {
      kunjunganMap[tgl] = 0;
    });

    if (logSheet) {
      const logData = logSheet.getDataRange().getValues();
      totalVisitor = Math.max(0, logData.length - 1); // Menghitung total semua baris log kunjungan
      
      for (let i = 1; i < logData.length; i++) {
        let tglLog = logData[i][0]; // Format yyyy-MM-dd
        
        // Amankan jika format yang diterima secara dinamis berubah menjadi object Date
        if (tglLog instanceof Date) {
            tglLog = Utilities.formatDate(tglLog, 'Asia/Jakarta', 'yyyy-MM-dd');
        } else {
            tglLog = tglLog.toString().trim();
        }

        if (kunjunganMap[tglLog] !== undefined) {
          kunjunganMap[tglLog]++;
        }
      }
    }

    // Ubah ke format display tanggal yang mudah dibaca (DD/MM) untuk grafik
    listTanggalTerakhir.forEach(tgl => {
      const parts = tgl.split('-');
      chartLabels.push(`${parts[2]}/${parts[1]}`);
      chartValues.push(kunjunganMap[tgl]);
    });

    const stats = {
      totalSiswa: totalSiswa,
      totalGuru: totalGuru,
      totalKelas: totalKelas,
      totalLMS: totalLMS,
      totalVisitor: totalVisitor,
      chartLabels: chartLabels,
      chartValues: chartValues
    };

    return { status: 'success', data: stats };
  } catch (error) {
    Logger.log("Error di getAdminStats: " + error.message);
    return { status: 'error', message: error.message };
  }
}
