/**
 * Modul Referensi & Profil - Backend Code By ZettBOT
 * File: modul_referensi.gs
 */

// Helper: Setup/Get Sheet (Auto-Build Database)
function zettGetOrCreateSheet(sheetName, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#e0e7ff");
    SpreadsheetApp.flush();
  }
  return sheet;
}

/** * ===================================================
 * LOGIKA PROFIL USER (DATA GURU, DATA SISWA & PENGGUNA)
 * =====================================================
 */

function getUserProfileFromSheets(clientName, sessionId) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetPengguna = ss.getSheetByName("Pengguna");
    if (!sheetPengguna) {
      return { status: 'error', message: 'Sheet Pengguna tidak ditemukan.' };
    }
    
    const dataPengguna = sheetPengguna.getDataRange().getValues();
    let loggedInUserId = null;
    
    // 1. Prioritas Utama: Cari berdasarkan Session ID / Username dari Browser Frontend
    if (sessionId) {
      for (let i = 1; i < dataPengguna.length; i++) {
        if (String(dataPengguna[i][0]) === String(sessionId) || String(dataPengguna[i][2]) === String(sessionId)) {
          loggedInUserId = dataPengguna[i][0];
          break;
        }
      }
    }
    
    // 2. Prioritas Kedua: Cari berdasarkan Nama yang tampil di Frontend UI
    if (!loggedInUserId && clientName && clientName !== 'Nama User') {
      for (let i = 1; i < dataPengguna.length; i++) {
        if (String(dataPengguna[i][1]).trim() === String(clientName).trim() || String(dataPengguna[i][2]).trim() === String(clientName).trim()) {
          loggedInUserId = dataPengguna[i][0];
          break;
        }
      }
    }
    
    // 3. Fallback: Cari menggunakan Cache Server GAS
    if (!loggedInUserId) {
      const userCache = CacheService.getUserCache();
      loggedInUserId = userCache.get("ACTIVE_USER_ID") || CacheService.getScriptCache().get("ACTIVE_USER_ID") || PropertiesService.getScriptProperties().getProperty("ACTIVE_USER_ID");
    }
    
    // 4. Failsafe Mode: Preview bypass
    if (!loggedInUserId && dataPengguna.length > 1) {
      loggedInUserId = dataPengguna[1][0]; 
    }
    
    if (!loggedInUserId) {
      return { status: 'error', message: 'Sesi aktif tidak ditemukan di server.' };
    }
    
    let userData = null;
    for (let i = 1; i < dataPengguna.length; i++) {
      if (String(dataPengguna[i][0]) === String(loggedInUserId)) {
        userData = {
          id: dataPengguna[i][0],
          nama: dataPengguna[i][1],
          username: dataPengguna[i][2],
          role: dataPengguna[i][4],
          status: dataPengguna[i][5],
          foto: dataPengguna[i].length > 6 ? dataPengguna[i][6] : ""
        };
        break;
      }
    }
    
    if (!userData) {
      return { status: 'error', message: 'Data Pengguna dengan ID ' + loggedInUserId + ' tidak ditemukan.' };
    }
    
    let resFoto = userData.foto || "";
    let extraData = { nip: "", mapel: "", nis: "", kelas: "" };
    
    // Tarik relasi foto & info di Data_Guru
    if (String(userData.role).toLowerCase() === 'guru') {
      const sheetGuru = ss.getSheetByName("Data_Guru");
      if (sheetGuru) {
        const dataGuru = sheetGuru.getDataRange().getValues();
        for (let i = 1; i < dataGuru.length; i++) {
          if (String(dataGuru[i][1]) === String(loggedInUserId)) {
            extraData.nip = dataGuru[i][2];
            extraData.mapel = dataGuru[i][4];
            if (!resFoto && dataGuru[i][5]) {
              resFoto = dataGuru[i][5]; 
            }
            break;
          }
        }
      }
    } 
    // Tarik relasi foto & info di Data_Siswa
    else if (String(userData.role).toLowerCase() === 'siswa') {
      const sheetSiswa = ss.getSheetByName("Data_Siswa");
      if (sheetSiswa) {
        const dataSiswa = sheetSiswa.getDataRange().getValues();
        for (let i = 1; i < dataSiswa.length; i++) {
          if (String(dataSiswa[i][2]) === String(loggedInUserId)) {
            extraData.nis = dataSiswa[i][3];
            extraData.kelas = dataSiswa[i][5];
            if (!resFoto && dataSiswa[i][6]) {
              resFoto = dataSiswa[i][6]; 
            }
            break;
          }
        }
      }
    }
    
    return {
      status: 'success',
      role: userData.role,
      nama: userData.nama,
      username: userData.username,
      nip: extraData.nip || '',
      mapel: extraData.mapel || '',
      nis: extraData.nis || '',
      kelas: extraData.kelas || '',
      foto: resFoto || ''
    };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

function updateUserProfil(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetPengguna = ss.getSheetByName("Pengguna");
    if (!sheetPengguna) return { status: 'error', message: 'Sheet Pengguna tidak ditemukan.' };
    
    const dataPengguna = sheetPengguna.getDataRange().getValues();
    let userRowIndex = -1;
    let role = "";
    let loggedInUserId = "";
    
    // 1. Sinkronisasi identitas dari Payload Frontend (Menjamin 100% akurat)
    if (payload.sessionId) {
       for (let i = 1; i < dataPengguna.length; i++) {
         if (String(dataPengguna[i][0]) === String(payload.sessionId) || String(dataPengguna[i][2]) === String(payload.sessionId)) {
           userRowIndex = i + 1;
           loggedInUserId = dataPengguna[i][0];
           role = String(dataPengguna[i][4]).toLowerCase();
           break;
         }
       }
    }
    
    if (userRowIndex === -1 && payload.clientName && payload.clientName !== 'Nama User') {
       for (let i = 1; i < dataPengguna.length; i++) {
         if (String(dataPengguna[i][1]).trim() === String(payload.clientName).trim() || String(dataPengguna[i][2]).trim() === String(payload.clientName).trim()) {
           userRowIndex = i + 1;
           loggedInUserId = dataPengguna[i][0];
           role = String(dataPengguna[i][4]).toLowerCase();
           break;
         }
       }
    }
    
    // 2. Fallback Cache
    if (userRowIndex === -1) {
      const userCache = CacheService.getUserCache();
      let cacheId = userCache.get("ACTIVE_USER_ID") || CacheService.getScriptCache().get("ACTIVE_USER_ID") || PropertiesService.getScriptProperties().getProperty("ACTIVE_USER_ID");
      if (cacheId) {
        for (let i = 1; i < dataPengguna.length; i++) {
          if (String(dataPengguna[i][0]) === String(cacheId)) {
            userRowIndex = i + 1;
            loggedInUserId = cacheId;
            role = String(dataPengguna[i][4]).toLowerCase();
            break;
          }
        }
      }
    }
    
    // 3. Failsafe Bypass
    if (userRowIndex === -1 && dataPengguna.length > 1) {
      userRowIndex = 2;
      loggedInUserId = dataPengguna[1][0];
      role = String(dataPengguna[1][4]).toLowerCase();
    }
    
    if (userRowIndex === -1) {
      return { status: 'error', message: 'Sesi aktif tidak ditemukan atau telah kadaluarsa.' };
    }
    
    let fotoData = "";
    if (payload.fotoBase64) {
      fotoData = "data:" + payload.fotoMimeType + ";base64," + payload.fotoBase64;
    }
    
    // --- PROSES SAVE KE SHEET PENGGUNA ---
    if (payload.username) {
      sheetPengguna.getRange(userRowIndex, 3).setValue(payload.username);
    }
    if (payload.password) {
      sheetPengguna.getRange(userRowIndex, 4).setValue(payload.password);
    }
    if (fotoData) {
      const colMax = sheetPengguna.getLastColumn();
      if (colMax < 7) {
        sheetPengguna.getRange(1, 7).setValue("Foto").setFontWeight("bold");
      }
      sheetPengguna.getRange(userRowIndex, 7).setValue(fotoData);
    }
    
    // --- PROSES SAVE KE SHEET RELASI (Data_Guru ATAU Data_Siswa) ---
    if (role === 'guru') {
      const sheetGuru = ss.getSheetByName("Data_Guru");
      if (sheetGuru) {
        const dataGuru = sheetGuru.getDataRange().getValues();
        for (let i = 1; i < dataGuru.length; i++) {
          // Cari baris Guru berdasarkan ID_User yang cocok
          if (String(dataGuru[i][1]) === String(loggedInUserId)) {
            if (fotoData) {
              sheetGuru.getRange(i + 1, 6).setValue(fotoData); // Update Kolom 6 (Foto)
            }
            break;
          }
        }
      }
    } else if (role === 'siswa') {
      const sheetSiswa = ss.getSheetByName("Data_Siswa");
      if (sheetSiswa) {
        const dataSiswa = sheetSiswa.getDataRange().getValues();
        for (let i = 1; i < dataSiswa.length; i++) {
          // Cari baris Siswa berdasarkan ID_User yang cocok
          if (String(dataSiswa[i][2]) === String(loggedInUserId)) {
            if (fotoData) {
              sheetSiswa.getRange(i + 1, 7).setValue(fotoData); // Update Kolom 7 (Foto)
            }
            break;
          }
        }
      }
    }
    
    SpreadsheetApp.flush();
    return { status: 'success', message: 'Profil dan foto berhasil diperbarui ke semua sheet.' };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

/** * ==============================
 * LOGIKA DATA GURU (MULTI MAPEL)
 * ==============================
 */

function getDataGuruZett() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetGuru = zettGetOrCreateSheet("Data_Guru", ["ID_Guru", "ID_User", "NIP", "Nama_Guru", "Mapel", "Foto", "Status"]);
    const sheetPengguna = ss.getSheetByName("Pengguna"); 
    
    const lrGuru = sheetGuru.getLastRow();
    const dataGuru = lrGuru > 0 ? sheetGuru.getRange(1, 1, lrGuru, 7).getValues() : [];
    
    const lrPengguna = sheetPengguna ? sheetPengguna.getLastRow() : 0;
    const dataPengguna = lrPengguna > 0 ? sheetPengguna.getRange(1, 1, lrPengguna, 6).getValues() : [];
    
    const userMap = {};
    if (dataPengguna.length > 1) {
      for (let i = 1; i < dataPengguna.length; i++) {
        userMap[dataPengguna[i][0]] = { 
          username: dataPengguna[i][2] || "", 
          password: dataPengguna[i][3] || "" 
        };
      }
    }

    const result = [];
    for (let i = 1; i < dataGuru.length; i++) {
      const row = dataGuru[i];
      if ((row[0] || row[3]) && row[6] !== "Dihapus") {
        const akun = userMap[row[1]] || { username: "", password: "" };
        
        let parsedMapel = [];
        if (row[4]) {
          try { 
            let tempJSON = JSON.parse(row[4]); 
            if(Array.isArray(tempJSON)) {
                parsedMapel = tempJSON.map(m => m.nama ? m.nama : m);
            }
          } catch (e) { 
            parsedMapel = String(row[4]).split(',').map(s => s.trim()).filter(s => s !== ""); 
          }
        }

        result.push({
          id_guru: String(row[0] || `MANUAL_G${i}`),
          id_user: String(row[1] || ""),
          nip: String(row[2] || "-"),
          nama_guru: String(row[3] || "Tanpa Nama"),
          mapel_data: parsedMapel, 
          foto: String(row[5] || ""),
          username: String(akun.username),
          password: String(akun.password)
        });
      }
    }
    return { status: 'success', data: result };
  } catch (error) {
    return { status: 'error', message: error.toString() };
  }
}

function simpanDataGuruZett(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetGuru = zettGetOrCreateSheet("Data_Guru", ["ID_Guru", "ID_User", "NIP", "Nama_Guru", "Mapel", "Foto", "Status"]);
    const sheetPengguna = zettGetOrCreateSheet("Pengguna", ["ID", "Nama Pengguna", "Username", "Password", "Role", "Status", "Foto"]);

    const idUser = "USR" + new Date().getTime();
    const idGuru = "GRU" + new Date().getTime();
    const mapelString = payload.mapel_data && payload.mapel_data.length > 0 ? payload.mapel_data.join(', ') : '';

    const rowPengguna = [idUser, payload.nama_guru, payload.username, payload.password, "guru", "Aktif", payload.foto || ""];
    const rowGuru = [idGuru, idUser, payload.nip, payload.nama_guru, mapelString, payload.foto, "Aktif"];

    sheetPengguna.appendRow(rowPengguna);
    sheetGuru.appendRow(rowGuru);
    SpreadsheetApp.flush(); 
    return { status: 'success', id: idGuru };
  } catch (error) { return { status: 'error', message: error.toString() }; }
}

function updateDataGuruZett(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetGuru = ss.getSheetByName("Data_Guru");
    const sheetPengguna = ss.getSheetByName("Pengguna");
    
    if (!sheetGuru) return { status: 'error', message: 'Sheet Data_Guru tidak ditemukan.' };

    const lrGuru = sheetGuru.getLastRow();
    if (lrGuru < 2) return { status: 'error', message: 'Data Guru kosong.' };
    
    const dataGuru = sheetGuru.getRange(1, 1, lrGuru, 7).getValues();
    let rowGuru = -1;
    let idUser = "";

    for (let i = 1; i < dataGuru.length; i++) {
      if (dataGuru[i][0] === payload.id_guru) {
        rowGuru = i + 1;
        idUser = dataGuru[i][1];
        break;
      }
    }

    if (rowGuru === -1) return { status: 'error', message: 'Data Guru tidak ditemukan.' };

    const mapelString = payload.mapel_data && payload.mapel_data.length > 0 ? payload.mapel_data.join(', ') : '';
    const existingStatus = dataGuru[rowGuru - 1][6] || "Aktif";
    
    const rowGuruUpdate = [[payload.id_guru, idUser, payload.nip, payload.nama_guru, mapelString, payload.foto, existingStatus]];
    sheetGuru.getRange(rowGuru, 1, 1, 7).setValues(rowGuruUpdate);

    if (idUser && sheetPengguna) {
      const lrPengguna = sheetPengguna.getLastRow();
      if(lrPengguna > 0) {
        const lcPengguna = sheetPengguna.getLastColumn();
        if (lcPengguna < 7) { sheetPengguna.getRange(1, 7).setValue("Foto").setFontWeight("bold"); }
        const dataPengguna = sheetPengguna.getDataRange().getValues();
        let rowPenggunaIdx = -1;
        for (let i = 1; i < dataPengguna.length; i++) {
          if (dataPengguna[i][0] === idUser) { rowPenggunaIdx = i + 1; break; }
        }
        if (rowPenggunaIdx !== -1) {
          const extRole = dataPengguna[rowPenggunaIdx-1][4] || "guru";
          const extStat = dataPengguna[rowPenggunaIdx-1][5] || "Aktif";
          sheetPengguna.getRange(rowPenggunaIdx, 1, 1, 7).setValues([[idUser, payload.nama_guru, payload.username, payload.password, extRole, extStat, payload.foto || ""]]);
        }
      }
    }
    SpreadsheetApp.flush(); 
    return { status: 'success' };
  } catch(e) { return {status: 'error', message: e.toString()}; }
}

function hapusDataGuruZett(idGuru) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetGuru = ss.getSheetByName("Data_Guru");
    const sheetPengguna = ss.getSheetByName("Pengguna");
    
    if(!sheetGuru) return { status: 'error', message: 'Sheet Data_Guru tidak ditemukan.' };

    const dataGuru = sheetGuru.getDataRange().getValues();
    let idUserTerkait = "";
    for (let i = 1; i < dataGuru.length; i++) {
      if (dataGuru[i][0] == idGuru) {
        idUserTerkait = dataGuru[i][1];
        sheetGuru.deleteRow(i + 1);
        break;
      }
    }
    
    if (idUserTerkait && sheetPengguna) {
      const dataPengguna = sheetPengguna.getDataRange().getValues();
      for (let i = 1; i < dataPengguna.length; i++) {
        if (dataPengguna[i][0] == idUserTerkait) {
          sheetPengguna.deleteRow(i + 1);
          break;
        }
      }
    }
    SpreadsheetApp.flush();
    return { status: 'success' };
  } catch (error) { return { status: 'error', message: error.toString() }; }
}

/** * ==============================
 * LOGIKA DATA SISWA
 * ==============================
 */

function getDataSiswaZett() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetSiswa = zettGetOrCreateSheet("Data_Siswa", ["ID_Siswa", "ID_TA", "ID_User", "NIS", "Nama_Siswa", "Kelas", "Foto", "Status"]);
    const sheetPengguna = ss.getSheetByName("Pengguna");
    
    const activeTaId = typeof getActiveTaId_ === 'function' ? getActiveTaId_(ss) : '-';

    let activeTaYearStr = "";
    const taYearMap = {};
    const sheetTA = ss.getSheetByName("Tahun_Ajaran");
    if (sheetTA) {
      const dataTA = sheetTA.getDataRange().getValues();
      for (let i = 1; i < dataTA.length; i++) {
        taYearMap[dataTA[i][0]] = String(dataTA[i][1]).trim();
        if (dataTA[i][0] === activeTaId) {
            activeTaYearStr = String(dataTA[i][1]).trim();
        }
      }
    }
    
    const lrSiswa = sheetSiswa.getLastRow();
    const dataSiswa = lrSiswa > 0 ? sheetSiswa.getRange(1, 1, lrSiswa, 8).getValues() : [];
    
    const lrPengguna = sheetPengguna ? sheetPengguna.getLastRow() : 0;
    const dataPengguna = lrPengguna > 0 ? sheetPengguna.getRange(1, 1, lrPengguna, 6).getValues() : [];
    
    const userMap = {};
    if (dataPengguna.length > 1) {
      for (let i = 1; i < dataPengguna.length; i++) {
        userMap[dataPengguna[i][0]] = { 
          username: dataPengguna[i][2] || "", 
          password: dataPengguna[i][3] || "" 
        };
      }
    }

    const result = [];
    for (let i = 1; i < dataSiswa.length; i++) {
      const row = dataSiswa[i];
      const rowTaId = row[1];
      const rowTaYearStr = taYearMap[rowTaId] || "-";

      if ((row[0] || row[4]) && row[7] !== "Dihapus" && (rowTaId === activeTaId || rowTaYearStr === activeTaYearStr)) {
        const akun = userMap[row[2]] || { username: "", password: "" };
        result.push({
          id_siswa: String(row[0] || `MANUAL_S${i}`),
          id_ta: String(row[1] || "-"),
          id_user: String(row[2] || ""),
          nis: String(row[3] || "-"),
          nama_siswa: String(row[4] || "Tanpa Nama"),
          kelas: String(row[5] || "-"),
          foto: String(row[6] || ""),
          username: String(akun.username),
          password: String(akun.password)
        });
      }
    }
    return { status: 'success', data: result };
  } catch (error) { return { status: 'error', message: error.toString() }; }
}

function simpanDataSiswaZett(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetSiswa = zettGetOrCreateSheet("Data_Siswa", ["ID_Siswa", "ID_TA", "ID_User", "NIS", "Nama_Siswa", "Kelas", "Foto", "Status"]);
    const sheetPengguna = zettGetOrCreateSheet("Pengguna", ["ID", "Nama Pengguna", "Username", "Password", "Role", "Status", "Foto"]);

    const activeTaId = typeof getActiveTaId_ === 'function' ? getActiveTaId_(ss) : '-';
    const idUser = "USR_S" + new Date().getTime();
    const idSiswa = "SIS" + new Date().getTime();

    const rowPengguna = [idUser, payload.nama_siswa, payload.username, payload.password, "siswa", "Aktif", payload.foto || ""];
    const rowSiswa = [idSiswa, activeTaId, idUser, payload.nis, payload.nama_siswa, payload.kelas, payload.foto, "Aktif"];

    sheetPengguna.appendRow(rowPengguna);
    sheetSiswa.appendRow(rowSiswa);
    SpreadsheetApp.flush();
    return { status: 'success', id: idSiswa };
  } catch (error) { return { status: 'error', message: error.toString() }; }
}

function simpanDataSiswaMassalZett(chunk) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetSiswa = zettGetOrCreateSheet("Data_Siswa", ["ID_Siswa", "ID_TA", "ID_User", "NIS", "Nama_Siswa", "Kelas", "Foto", "Status"]);
    const sheetPengguna = zettGetOrCreateSheet("Pengguna", ["ID", "Nama Pengguna", "Username", "Password", "Role", "Status", "Foto"]);
    
    if(!sheetSiswa || !sheetPengguna) return {status: 'error', message: 'Database tidak ditemukan.'};

    const activeTaId = typeof getActiveTaId_ === 'function' ? getActiveTaId_(ss) : '-';
    const rowsPengguna = [];
    const rowsSiswa = [];
    const baseTime = new Date().getTime();

    for (let i = 0; i < chunk.length; i++) {
      const nis = String(chunk[i][0]).trim();
      const nama = String(chunk[i][1]).trim();
      const kelas = String(chunk[i][2]).trim();
      
      const idUser = "USR_S" + (baseTime + i);
      const idSiswa = "SIS" + (baseTime + i);

      rowsPengguna.push([idUser, nama, nis, nis, "siswa", "Aktif", ""]);
      rowsSiswa.push([idSiswa, activeTaId, idUser, nis, nama, kelas, "", "Aktif"]);
    }

    if (rowsPengguna.length > 0) {
      const lc = sheetPengguna.getLastColumn();
      if (lc < 7) { sheetPengguna.getRange(1, 7).setValue("Foto").setFontWeight("bold"); }
      sheetPengguna.getRange(sheetPengguna.getLastRow() + 1, 1, rowsPengguna.length, 7).setValues(rowsPengguna);
    }
    if (rowsSiswa.length > 0) {
      sheetSiswa.getRange(sheetSiswa.getLastRow() + 1, 1, rowsSiswa.length, 8).setValues(rowsSiswa);
    }

    SpreadsheetApp.flush();
    return { status: 'success' };
  } catch (error) { return { status: 'error', message: error.toString() }; }
}

function updateDataSiswaZett(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetSiswa = ss.getSheetByName("Data_Siswa");
    const sheetPengguna = ss.getSheetByName("Pengguna");
    
    if (!sheetSiswa) return { status: 'error', message: 'Sheet Data_Siswa tidak ditemukan.' };

    const lrSiswa = sheetSiswa.getLastRow();
    if (lrSiswa < 2) return { status: 'error', message: 'Data Siswa kosong.' };

    const dataSiswa = sheetSiswa.getRange(1, 1, lrSiswa, 8).getValues();
    let rowSiswa = -1;
    let idUser = "";
    let idTA = "";

    for (let i = 1; i < dataSiswa.length; i++) {
      if (dataSiswa[i][0] === payload.id_siswa) {
        rowSiswa = i + 1;
        idTA = dataSiswa[i][1];
        idUser = dataSiswa[i][2];
        break;
      }
    }

    if (rowSiswa === -1) return { status: 'error', message: 'Data Siswa tidak ditemukan.' };

    const existingStatus = dataSiswa[rowSiswa - 1][7] || "Aktif";
    const rowSiswaUpdate = [[payload.id_siswa, idTA, idUser, payload.nis, payload.nama_siswa, payload.kelas, payload.foto, existingStatus]];
    sheetSiswa.getRange(rowSiswa, 1, 1, 8).setValues(rowSiswaUpdate);

    if (idUser && sheetPengguna) {
      const lrPengguna = sheetPengguna.getLastRow();
      if(lrPengguna > 0) {
        const lcPengguna = sheetPengguna.getLastColumn();
        if (lcPengguna < 7) { sheetPengguna.getRange(1, 7).setValue("Foto").setFontWeight("bold"); }
        const dataPengguna = sheetPengguna.getDataRange().getValues();
        let rowPenggunaIdx = -1;
        for (let i = 1; i < dataPengguna.length; i++) {
          if (dataPengguna[i][0] === idUser) { rowPenggunaIdx = i + 1; break; }
        }
        if (rowPenggunaIdx !== -1) {
          const extRole = dataPengguna[rowPenggunaIdx-1][4] || "siswa";
          const extStat = dataPengguna[rowPenggunaIdx-1][5] || "Aktif";
          sheetPengguna.getRange(rowPenggunaIdx, 1, 1, 7).setValues([[idUser, payload.nama_siswa, payload.username, payload.password, extRole, extStat, payload.foto || ""]]);
        }
      }
    }
    SpreadsheetApp.flush();
    return { status: 'success' };
  } catch(e) { return {status: 'error', message: e.toString()}; }
}

function hapusDataSiswaZett(idSiswa) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetSiswa = ss.getSheetByName("Data_Siswa");
    const sheetPengguna = ss.getSheetByName("Pengguna");
    
    if(!sheetSiswa) return { status: 'error', message: 'Sheet Data_Siswa tidak ditemukan.' };

    const dataSiswa = sheetSiswa.getDataRange().getValues();
    let idUserTerkait = "";
    for (let i = 1; i < dataSiswa.length; i++) {
      if (dataSiswa[i][0] == idSiswa) {
        idUserTerkait = dataSiswa[i][2]; 
        sheetSiswa.deleteRow(i + 1);
        break;
      }
    }

    if (idUserTerkait && sheetPengguna) {
      const dataPengguna = sheetPengguna.getDataRange().getValues();
      for (let i = 1; i < dataPengguna.length; i++) {
        if (dataPengguna[i][0] == idUserTerkait) {
          sheetPengguna.deleteRow(i + 1);
          break;
        }
      }
    }
    SpreadsheetApp.flush();
    return { status: 'success' };
  } catch (error) { return { status: 'error', message: error.toString() }; }
}

function getKelasPerTaZett() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetSiswa = ss.getSheetByName("Data_Siswa");
    if (!sheetSiswa) return { status: 'success', data: {} };
    
    const data = sheetSiswa.getDataRange().getValues();
    const mapKelas = {}; 
    
    for (let i = 1; i < data.length; i++) {
      const ta = data[i][1];
      const kelas = data[i][5];
      const status = data[i][7];
      if (status !== 'Dihapus' && ta && kelas && kelas !== '-') {
        if (!mapKelas[ta]) mapKelas[ta] = [];
        if (mapKelas[ta].indexOf(kelas) === -1) {
          mapKelas[ta].push(kelas);
        }
      }
    }
    return { status: 'success', data: mapKelas };
  } catch (e) {
    return { status: 'error', message: e.toString() };
  }
}

function naikKelasSiswaZett(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetSiswa = ss.getSheetByName("Data_Siswa");
    if (!sheetSiswa) return { status: 'error', message: 'Sheet Data_Siswa tidak ditemukan.' };

    const activeTaId = typeof getActiveTaId_ === 'function' ? getActiveTaId_(ss) : '-';
    
    const sheetTA = ss.getSheetByName("Tahun_Ajaran");
    if (sheetTA) {
      const dataTA = sheetTA.getDataRange().getValues();
      let taAsalYear = "";
      let activeTaYear = "";
      for (let i = 1; i < dataTA.length; i++) {
        if (dataTA[i][0] === payload.taAsal) taAsalYear = String(dataTA[i][1]).trim();
        if (dataTA[i][0] === activeTaId) activeTaYear = String(dataTA[i][1]).trim(); 
      }
      
      if (taAsalYear !== "" && taAsalYear === activeTaYear) {
         return { 
             status: 'error', 
             message: 'Tidak perlu melakukan proses Naik Kelas untuk Tahun Ajaran yang sama (' + taAsalYear + '). Data Siswa otomatis dipertahankan saat berganti semester!' 
         };
      }
    }

    const lr = sheetSiswa.getLastRow();
    if (lr < 2) return { status: 'error', message: 'Data Siswa kosong.' };

    const dataSiswa = sheetSiswa.getRange(1, 1, lr, 8).getValues();
    let rowsToAppend = [];
    let rowsToUpdate = [];

    for (let i = 1; i < dataSiswa.length; i++) {
      if (dataSiswa[i][1] === payload.taAsal && dataSiswa[i][5] === payload.kelasAsal && dataSiswa[i][7] !== 'Dihapus') {
        
        let alreadyInActiveTa = false;
        for (let j = 1; j < dataSiswa.length; j++) {
            if (dataSiswa[j][1] === activeTaId && dataSiswa[j][3] === dataSiswa[i][3] && dataSiswa[j][7] !== 'Dihapus') {
                alreadyInActiveTa = true;
                break;
            }
        }

        if (payload.taAsal === activeTaId) {
            dataSiswa[i][5] = payload.kelasTujuan;
            rowsToUpdate.push({rowIdx: i + 1, data: dataSiswa[i]});
        } else if (!alreadyInActiveTa) {
            const idSiswaNew = "SIS" + new Date().getTime() + i;
            const newRow = [
                idSiswaNew, 
                activeTaId, 
                dataSiswa[i][2], 
                dataSiswa[i][3], 
                dataSiswa[i][4], 
                payload.kelasTujuan, 
                dataSiswa[i][6], 
                "Aktif"
            ];
            rowsToAppend.push(newRow);
        }
      }
    }

    if (rowsToAppend.length > 0) {
      sheetSiswa.getRange(lr + 1, 1, rowsToAppend.length, 8).setValues(rowsToAppend);
      SpreadsheetApp.flush();
      return { status: 'success', message: `Berhasil memindahkan ${rowsToAppend.length} siswa ke TA Aktif.` };
    } else if (rowsToUpdate.length > 0) {
      rowsToUpdate.forEach(item => {
          sheetSiswa.getRange(item.rowIdx, 1, 1, 8).setValues([item.data]);
      });
      SpreadsheetApp.flush();
      return { status: 'success', message: `Berhasil merubah kelas ${rowsToUpdate.length} siswa di TA Aktif.` };
    } else {
      return { status: 'error', message: 'Tidak ada siswa baru yang dapat dipindahkan. (Siswa tersebut mungkin sudah ada di TA Aktif).' };
    }
  } catch(e) { return {status: 'error', message: e.toString()}; }
}

/** * ==============================
 * LOGIKA MAPING GURU
 * ==============================
 */

function getDataMapingZett() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetMaping = zettGetOrCreateSheet("Maping_Guru", ["ID_Maping", "ID_TA", "ID_Guru", "NIP", "Nama_Guru", "Mapel", "Kelas", "Status"]);
    
    const activeTaId = typeof getActiveTaId_ === 'function' ? getActiveTaId_(ss) : '-';

    let activeTaYearStr = "";
    const taYearMap = {};
    const sheetTA = ss.getSheetByName("Tahun_Ajaran");
    if (sheetTA) {
      const dataTA = sheetTA.getDataRange().getValues();
      for (let i = 1; i < dataTA.length; i++) {
        taYearMap[dataTA[i][0]] = String(dataTA[i][1]).trim();
        if (dataTA[i][0] === activeTaId) {
            activeTaYearStr = String(dataTA[i][1]).trim();
        }
      }
    }
    
    const lrMaping = sheetMaping.getLastRow();
    const dataMaping = lrMaping > 0 ? sheetMaping.getRange(1, 1, lrMaping, 8).getValues() : [];

    const mapingMap = {};
    for (let i = 1; i < dataMaping.length; i++) {
      const rowTaId = dataMaping[i][1];
      const rowTaYearStr = taYearMap[rowTaId] || "-";

      if (dataMaping[i][7] !== "Dihapus" && (rowTaId === activeTaId || rowTaYearStr === activeTaYearStr)) {
        const key = dataMaping[i][2] + "_" + dataMaping[i][5]; 
        
        if (!mapingMap[key] || rowTaId === activeTaId) {
            mapingMap[key] = {
              id_maping: String(dataMaping[i][0]),
              id_guru: String(dataMaping[i][2]),
              nip: String(dataMaping[i][3] || "-"),
              nama_guru: String(dataMaping[i][4] || "-"),
              mapel: String(dataMaping[i][5] || "-"),
              kelas: String(dataMaping[i][6] || "-")
            };
        }
      }
    }
    
    const result = Object.keys(mapingMap).map(function(k) { return mapingMap[k]; });
    return { status: 'success', data: result };
  } catch (error) { return { status: 'error', message: error.toString() }; }
}

function simpanDataMapingZett(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetMaping = zettGetOrCreateSheet("Maping_Guru", ["ID_Maping", "ID_TA", "ID_Guru", "NIP", "Nama_Guru", "Mapel", "Kelas", "Status"]);
    
    const idMaping = "MPG" + new Date().getTime();
    const activeTaId = typeof getActiveTaId_ === 'function' ? getActiveTaId_(ss) : '-';
    
    sheetMaping.appendRow([idMaping, activeTaId, payload.id_guru, payload.nip, payload.nama_guru, payload.mapel, payload.kelas, "Aktif"]);
    SpreadsheetApp.flush();
    return { status: 'success', id: idMaping };
  } catch (e) { return { status: 'error', message: e.toString() }; }
}

function updateDataMapingZett(payload) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetMaping = ss.getSheetByName("Maping_Guru");
    if(!sheetMaping) return { status: 'error', message: 'Sheet Maping_Guru tidak ditemukan.' };

    const lrMaping = sheetMaping.getLastRow();
    if (lrMaping < 2) return { status: 'error', message: 'Data Maping kosong.' };

    const dataMaping = sheetMaping.getRange(1, 1, lrMaping, 8).getValues();
    let rowTarget = -1;
    let idTA = "";

    for (let i = 1; i < dataMaping.length; i++) {
      if (dataMaping[i][0] === payload.id_maping) {
        rowTarget = i + 1;
        idTA = dataMaping[i][1];
        break;
      }
    }

    if (rowTarget === -1) return { status: 'error', message: 'Data Maping tidak ditemukan.' };

    const activeTaId = typeof getActiveTaId_ === 'function' ? getActiveTaId_(ss) : '-';
    const existingStatus = dataMaping[rowTarget - 1][7] || "Aktif";
    
    if (idTA !== activeTaId) {
        const idMapingBaru = "MPG" + new Date().getTime();
        sheetMaping.appendRow([idMapingBaru, activeTaId, payload.id_guru, payload.nip, payload.nama_guru, payload.mapel, payload.kelas, existingStatus]);
    } else {
        const rowUpdate = [[payload.id_maping, idTA, payload.id_guru, payload.nip, payload.nama_guru, payload.mapel, payload.kelas, existingStatus]];
        sheetMaping.getRange(rowTarget, 1, 1, 8).setValues(rowUpdate);
    }

    SpreadsheetApp.flush();
    return { status: 'success' };
  } catch(e) { return {status: 'error', message: e.toString()}; }
}

function hapusDataMapingZett(idMaping) {
  try {
    const sheetMaping = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Maping_Guru");
    if(!sheetMaping) return { status: 'error', message: 'Sheet Maping_Guru tidak ditemukan.' };

    const dataMaping = sheetMaping.getDataRange().getValues();
    for (let i = 1; i < dataMaping.length; i++) {
      if (dataMaping[i][0] == idMaping) {
        sheetMaping.deleteRow(i + 1);
        break;
      }
    }
    SpreadsheetApp.flush();
    return { status: 'success' };
  } catch (error) { return { status: 'error', message: error.toString() }; }
}
