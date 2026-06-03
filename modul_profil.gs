/**
 * Modul_Profil.gs
 * Logic backend terpadu untuk Pengguna, Tahun Ajaran, Setting Aplikasi, dan Pengumuman
 */

function getPenggunaData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Pengguna');
    
    if (!sheet) {
      sheet = ss.insertSheet('Pengguna');
      sheet.appendRow(['ID', 'Nama Pengguna', 'Username', 'Password', 'Role', 'Status']);
      sheet.appendRow(['U-001', 'Admin Master', 'admin', 'admin123', 'admin', 'Aktif']);
      sheet.appendRow(['U-002', 'Guru Sosiologi', 'guru1', 'guru123', 'guru', 'Aktif']);
      sheet.appendRow(['U-003', 'Siswa Teladan', 'siswa1', 'siswa123', 'siswa', 'Non-Aktif']);
      
      sheet.getRange("A1:F1").setFontWeight("bold").setBackground("#e0e7ff");
      sheet.setColumnWidth(2, 200); 
      SpreadsheetApp.flush(); 
    }

    const data = sheet.getDataRange().getValues();
    const rows = data.length > 1 ? data.slice(1) : []; 
    
    const formattedData = rows.map((row, index) => ({
      id: (row[0] || ('U-' + (index + 1))).toString(),
      nama: (row[1] || '-').toString(),
      username: (row[2] || '-').toString(),
      password: (row[3] || '-').toString(),
      role: (row[4] || '-').toString(),
      status: (row[5] || 'Aktif').toString()
    }));

    return { status: "success", data: formattedData };
  } catch (error) {
    Logger.log("Error di getPenggunaData: " + error.message);
    return { status: "error", message: error.message };
  }
}

function savePenggunaData(payload) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pengguna');
    const data = sheet.getDataRange().getValues();
    
    let savedRow = {};

    if (payload.id) {
      // Edit Data
      const rowIndex = data.findIndex(row => row[0].toString() === payload.id);
      if (rowIndex > -1) {
        sheet.getRange(rowIndex + 1, 2, 1, 5).setValues([[payload.nama, payload.username, payload.password, payload.role, payload.status]]);
      }
      savedRow = { id: payload.id, nama: payload.nama, username: payload.username, password: payload.password, role: payload.role, status: payload.status };
    } else {
      // Tambah Data Baru (Auto Generate ID)
      const newId = 'U-' + Utilities.formatDate(new Date(), "GMT", "yyMMddHHmmss");
      sheet.appendRow([newId, payload.nama, payload.username, payload.password, payload.role, payload.status]);
      savedRow = { id: newId, nama: payload.nama, username: payload.username, password: payload.password, role: payload.role, status: payload.status };
    }
    
    SpreadsheetApp.flush();
    // Mengirim kembali savedRow ke Frontend untuk dirender secara spesifik per baris
    return { status: "success", message: "Data pengguna berhasil disimpan!", data: savedRow };
  } catch (error) {
    return { status: "error", message: error.message };
  }
}

function togglePenggunaStatus(id) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pengguna');
    const data = sheet.getDataRange().getValues();
    const rowIndex = data.findIndex(row => row[0].toString() === id);
    
    if (rowIndex > -1) {
      const currentStatus = data[rowIndex][5];
      const newStatus = currentStatus === 'Aktif' ? 'Non-Aktif' : 'Aktif';
      sheet.getRange(rowIndex + 1, 6).setValue(newStatus);
      SpreadsheetApp.flush();
      return { status: "success", message: "Status pengguna berhasil diubah!", newStatus: newStatus };
    }
    return { status: "error", message: "Data tidak ditemukan." };
  } catch (error) {
    return { status: "error", message: error.message };
  }
}

function deletePenggunaData(id) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pengguna');
    const data = sheet.getDataRange().getValues();
    const rowIndex = data.findIndex(row => row[0].toString() === id);
    
    if (rowIndex > -1) {
      sheet.deleteRow(rowIndex + 1);
      SpreadsheetApp.flush();
      return { status: "success", message: "Pengguna berhasil dihapus!" };
    }
    return { status: "error", message: "Data tidak ditemukan." };
  } catch (error) {
    return { status: "error", message: error.message };
  }
}

function getTahunAjaranData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Tahun_Ajaran');
    
    if (!sheet) {
      sheet = ss.insertSheet('Tahun_Ajaran');
      sheet.appendRow(['ID', 'Tahun Ajaran', 'Semester', 'Status']);
      sheet.appendRow(['TA-001', '2023/2024', 'Ganjil', 'Non-Aktif']);
      sheet.appendRow(['TA-002', '2023/2024', 'Genap', 'Aktif']);
      
      sheet.getRange("A1:D1").setFontWeight("bold").setBackground("#e0e7ff");
      sheet.setColumnWidth(2, 150); 
      SpreadsheetApp.flush(); 
    }

    const data = sheet.getDataRange().getValues();
    const rows = data.length > 1 ? data.slice(1) : []; 
    
    const formattedData = rows.map((row, index) => ({
      id: (row[0] || ('TA-' + (index + 1))).toString(),
      tahun: (row[1] || '-').toString(),
      semester: (row[2] || '-').toString(),
      status: (row[3] || 'Non-Aktif').toString()
    }));

    return { status: "success", data: formattedData };
  } catch (error) {
    Logger.log("Error di getTahunAjaranData: " + error.message);
    return { status: "error", message: error.message };
  }
}

function saveTahunAjaranData(payload) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tahun_Ajaran');
    const data = sheet.getDataRange().getValues();

    // Aturan eksklusif: Hanya boleh ada 1 TA yang aktif
    if (payload.status === 'Aktif') {
      for (let i = 1; i < data.length; i++) {
        sheet.getRange(i + 1, 4).setValue('Non-Aktif');
      }
    }
    
    let savedRow = {};

    if (payload.id) {
      const rowIndex = data.findIndex(row => row[0].toString() === payload.id);
      if (rowIndex > -1) {
        sheet.getRange(rowIndex + 1, 2, 1, 3).setValues([[payload.tahun, payload.semester, payload.status]]);
      }
      savedRow = { id: payload.id, tahun: payload.tahun, semester: payload.semester, status: payload.status };
    } else {
      const newId = 'TA-' + Utilities.formatDate(new Date(), "GMT", "yyMMddHHmmss");
      sheet.appendRow([newId, payload.tahun, payload.semester, payload.status]);
      savedRow = { id: newId, tahun: payload.tahun, semester: payload.semester, status: payload.status };
    }
    
    SpreadsheetApp.flush();
    return { status: "success", message: "Tahun Ajaran berhasil disimpan!", data: savedRow };
  } catch (error) {
    return { status: "error", message: error.message };
  }
}

// ZettBOT: FUNGSI BARU UNTUK TOGGLE STATUS TAHUN AJARAN
function toggleTahunAjaranStatus(id) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tahun_Ajaran');
    const data = sheet.getDataRange().getValues();
    const rowIndex = data.findIndex(row => row[0].toString() === id);
    
    if (rowIndex > -1) {
      const currentStatus = data[rowIndex][3];
      const newStatus = currentStatus === 'Aktif' ? 'Non-Aktif' : 'Aktif';
      
      // Jika diubah jadi Aktif, jadikan yang lain Non-Aktif
      if (newStatus === 'Aktif') {
        for (let i = 1; i < data.length; i++) {
          if (i !== rowIndex) {
             sheet.getRange(i + 1, 4).setValue('Non-Aktif');
          }
        }
      }

      sheet.getRange(rowIndex + 1, 4).setValue(newStatus);
      SpreadsheetApp.flush();
      return { status: "success", message: "Status Tahun Ajaran berhasil diubah!", newStatus: newStatus };
    }
    return { status: "error", message: "Data tidak ditemukan." };
  } catch (error) {
    return { status: "error", message: error.message };
  }
}

function deleteTahunAjaranData(id) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tahun_Ajaran');
    const data = sheet.getDataRange().getValues();
    const rowIndex = data.findIndex(row => row[0].toString() === id);
    
    if (rowIndex > -1) {
      sheet.deleteRow(rowIndex + 1);
      SpreadsheetApp.flush();
      return { status: "success", message: "Tahun Ajaran berhasil dihapus!" };
    }
    return { status: "error", message: "Data tidak ditemukan." };
  } catch (error) {
    return { status: "error", message: error.message };
  }
}

function getSettingAplikasi() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Setting_Aplikasi');
    
    if (!sheet) {
      sheet = ss.insertSheet('Setting_Aplikasi');
      sheet.appendRow(['Nama_Aplikasi', 'Logo_Base64_1', 'Logo_Base64_2', 'Logo_Base64_3']); 
      sheet.appendRow(['EduZett - LMS Terpadu', '']); 
      sheet.getRange("A1:D1").setFontWeight("bold").setBackground("#e0e7ff");
      sheet.setColumnWidth(2, 300); 
      SpreadsheetApp.flush(); 
    }

    const data = sheet.getDataRange().getValues();
    const nama = data.length > 1 ? data[1][0] : 'EduZett - LMS Terpadu';
    
    let logo = '';
    if (data.length > 1) {
      for (let i = 1; i < data[1].length; i++) {
        if (data[1][i]) {
          logo += data[1][i].toString();
        }
      }
    }

    return { status: "success", data: { nama: nama.toString(), logo: logo } };
  } catch (error) {
    Logger.log("Error di getSettingAplikasi: " + error.message);
    return { status: "error", message: error.message };
  }
}

function saveSettingAplikasi(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Setting_Aplikasi');

    if (!sheet) return { status: "error", message: "Database Setting belum siap." };

    sheet.getRange(2, 1).setValue(payload.nama);

    const logoStr = payload.logo || '';
    const chunkSize = 45000; 
    const chunks = [];
    
    for (let i = 0; i < logoStr.length; i += chunkSize) {
      chunks.push(logoStr.substring(i, i + chunkSize));
    }

    const maxCol = sheet.getMaxColumns();
    if (maxCol > 1) {
      sheet.getRange(2, 2, 1, maxCol - 1).clearContent();
    }

    if (1 + chunks.length > maxCol) {
      sheet.insertColumnsAfter(maxCol, (1 + chunks.length) - maxCol);
    }

    if (chunks.length > 0) {
      sheet.getRange(2, 2, 1, chunks.length).setValues([chunks]);
    }
    
    SpreadsheetApp.flush();
    return { status: "success", message: "Setting aplikasi berhasil diperbarui!" };
  } catch (error) {
    return { status: "error", message: error.message };
  }
}

function getActiveTaId_(ss) {
  const sheet = ss.getSheetByName('Tahun_Ajaran');
  if(!sheet) return '-';
  const data = sheet.getDataRange().getValues();
  for(let i=1; i<data.length; i++) {
    if(data[i][3] === 'Aktif') return data[i][0]; 
  }
  return '-';
}

function getPengumumanData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Pengumuman');
    
    if (!sheet) {
      sheet = ss.insertSheet('Pengumuman');
      sheet.appendRow(['ID', 'ID_Tahun_Ajaran', 'Judul', 'Pengumuman', 'Tanggal_Buat', 'Status', 'Gambar']);
      sheet.getRange("A1:G1").setFontWeight("bold").setBackground("#e0e7ff");
      sheet.setColumnWidth(4, 300); 
      SpreadsheetApp.flush(); 
    }

    const data = sheet.getDataRange().getValues();
    const rows = data.length > 1 ? data.slice(1) : []; 
    
    const formattedData = rows.map((row, index) => ({
      id: (row[0] || ('P-' + (index + 1))).toString(),
      id_ta: (row[1] || '-').toString(),
      judul: (row[2] || '-').toString(),
      isi: (row[3] || '-').toString(),
      tanggal: (row[4] || '-').toString(),
      status: (row[5] || 'Aktif').toString(),
      gambar: (row[6] || '').toString()
    }));

    return { status: "success", data: formattedData };
  } catch (error) {
    Logger.log("Error di getPengumumanData: " + error.message);
    return { status: "error", message: error.message };
  }
}

function savePengumumanData(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Pengumuman');
    const data = sheet.getDataRange().getValues();
    
    const timestamp = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    let savedRow = {};

    if (payload.id) {
      // Edit Data
      const rowIndex = data.findIndex(row => row[0].toString() === payload.id);
      let existingData = { id_ta: '-', tanggal: timestamp, status: 'Aktif', gambar: '' };
      
      if (rowIndex > -1) {
        existingData.id_ta = data[rowIndex][1] || '-';
        existingData.tanggal = data[rowIndex][4] || timestamp;
        existingData.status = data[rowIndex][5] || 'Aktif';
        existingData.gambar = data[rowIndex][6] || '';
        
        sheet.getRange(rowIndex + 1, 3, 1, 2).setValues([[payload.judul, payload.isi]]);
        if(payload.gambar !== undefined) {
           sheet.getRange(rowIndex + 1, 7).setValue(payload.gambar);
        }
      }
      
      savedRow = {
        id: payload.id,
        id_ta: existingData.id_ta,
        judul: payload.judul,
        isi: payload.isi,
        tanggal: existingData.tanggal,
        status: existingData.status,
        gambar: payload.gambar !== undefined ? payload.gambar : existingData.gambar
      };
      
    } else {
      // Tambah Data Baru
      const newId = 'P-' + Utilities.formatDate(new Date(), "GMT", "yyMMddHHmmss");
      const activeTaId = getActiveTaId_(ss); 
      const newGambar = payload.gambar || '';
      
      sheet.appendRow([newId, activeTaId, payload.judul, payload.isi, timestamp, 'Aktif', newGambar]);
      
      savedRow = {
        id: newId,
        id_ta: activeTaId,
        judul: payload.judul,
        isi: payload.isi,
        tanggal: timestamp,
        status: 'Aktif',
        gambar: newGambar
      };
    }
    
    SpreadsheetApp.flush();
    return { status: "success", message: "Pengumuman berhasil disimpan!", data: savedRow };
  } catch (error) {
    return { status: "error", message: error.message };
  }
}

function deletePengumumanData(id) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pengumuman');
    const data = sheet.getDataRange().getValues();
    const rowIndex = data.findIndex(row => row[0].toString() === id);
    
    if (rowIndex > -1) {
      sheet.deleteRow(rowIndex + 1);
      SpreadsheetApp.flush();
      return { status: "success", message: "Pengumuman berhasil dihapus!" };
    }
    return { status: "error", message: "Data tidak ditemukan." };
  } catch (error) {
    return { status: "error", message: error.message };
  }
}

function togglePengumumanStatus(id) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Pengumuman');
    const data = sheet.getDataRange().getValues();
    const rowIndex = data.findIndex(row => row[0].toString() === id);
    
    if (rowIndex > -1) {
      const currentStatus = data[rowIndex][5];
      const newStatus = currentStatus === 'Aktif' ? 'Non-Aktif' : 'Aktif';
      sheet.getRange(rowIndex + 1, 6).setValue(newStatus);
      SpreadsheetApp.flush();
      return { status: "success", message: "Status pengumuman berhasil diubah!", newStatus: newStatus };
    }
    return { status: "error", message: "Data tidak ditemukan." };
  } catch (error) {
    return { status: "error", message: error.message };
  }
}
