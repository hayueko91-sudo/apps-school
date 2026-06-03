// ==========================================
// FUNGSI PANCINGAN UTAMA (WAJIB DIJALANKAN SEKALI)
// ==========================================
function forceAuthorizeDrive() {
  try {
    var fileTest = DriveApp.getFiles();
    var sheetTest = SpreadsheetApp.getActiveSpreadsheet();
    var pesan = "Otorisasi Google Drive dan Google Sheets Berhasil! Sistem siap 100%.";
    Logger.log(pesan);
    return pesan;
  } catch (e) {
    Logger.log("Gagal: " + e.message);
    throw new Error("Gagal otorisasi: " + e.message);
  }
}

function checkAndCreateSheetTugas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Tugas');
  if (!sheet) {
    sheet = ss.insertSheet('Tugas');
    const headers = [
      "ID_Tugas", "Tanggal_Buat", "ID_TA", "Nama_Guru", "Mapel", "Judul_Tugas", "Kelas",
      "Tanggal_Mulai", "Jam_Mulai", "Tanggal_Selesai", "Jam_Selesai",
      "Instruksi", "Foto_Instruksi", 
      "Req_Teks", "Req_Foto", "Req_Video", "Req_File", "Status"
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#e0e7ff");
    SpreadsheetApp.flush();
  }
  return sheet;
}

function checkAndCreateSheetKumpulTugas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Kumpul_Tugas');
  
  if (!sheet) {
    sheet = ss.insertSheet('Kumpul_Tugas');
    const headers = [
      "ID_Kumpul", "ID_Tugas", "Nama_Guru", "Mapel", "NIS_ID", "Nama_Siswa", "Kelas", 
      "Waktu_Kumpul", "Jawaban_Teks", "File_Foto", "File_Video", "File_Dokumen", 
      "Nilai", "Catatan_Guru"
    ];
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#e0e7ff");
    SpreadsheetApp.flush();
  } else {
    let lc = sheet.getLastColumn();
    if (lc > 0) {
        let existingHeaders = sheet.getRange(1, 1, 1, lc).getValues()[0].map(h => h.toString().toLowerCase().trim());
        let needFlush = false;
        
        if (existingHeaders.indexOf('nama_guru') === -1) {
            lc++;
            sheet.getRange(1, lc).setValue("Nama_Guru").setFontWeight("bold").setBackground("#e0e7ff");
            needFlush = true;
        }
        if (existingHeaders.indexOf('mapel') === -1) {
            lc++;
            sheet.getRange(1, lc).setValue("Mapel").setFontWeight("bold").setBackground("#e0e7ff");
            needFlush = true;
        }
        
        if(needFlush) SpreadsheetApp.flush();
    }
  }
  return sheet;
}

function getMapingGuruForTugasList() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let allowedTAs = [];
    
    const sheetTA = ss.getSheetByName('Tahun_Ajaran');
    if (sheetTA) {
      const dataTA = sheetTA.getDataRange().getValues();
      if (dataTA.length > 1) {
        const headersTA = dataTA[0].map(h => h.toString().toLowerCase().trim());
        let idxTahun = headersTA.indexOf('tahun');
        if (idxTahun === -1) idxTahun = headersTA.indexOf('tahun_ajaran');
        if (idxTahun === -1) idxTahun = 1; 
        
        let activeTahun = "";
        for (let i = 1; i < dataTA.length; i++) {
          let isAktif = dataTA[i].some(cell => cell.toString().toLowerCase().trim() === 'aktif');
          if (isAktif) {
            activeTahun = dataTA[i][idxTahun] ? dataTA[i][idxTahun].toString().trim() : ""; 
            break;
          }
        }

        if (activeTahun !== "") {
           for (let i = 1; i < dataTA.length; i++) {
             let loopTahun = dataTA[i][idxTahun] ? dataTA[i][idxTahun].toString().trim() : "";
             if (loopTahun.toLowerCase() === activeTahun.toLowerCase()) {
                 let idTA = dataTA[i][0] ? dataTA[i][0].toString().trim().toLowerCase() : "";
                 if(idTA !== "") allowedTAs.push(idTA);
             }
           }
        }
      }
    }

    const sheet = ss.getSheetByName('Maping_Guru');
    if(!sheet) throw new Error("Sheet Maping_Guru tidak ditemukan.");

    const data = sheet.getDataRange().getValues();
    if(data.length <= 1) return { status: 'success', data: [] };

    const headers = data[0].map(h => h.toString().toLowerCase().trim());
    let idxTA = headers.indexOf('id_ta'); if (idxTA === -1) idxTA = 1; 
    let idxGuru = headers.indexOf('nama_guru');
    if (idxGuru === -1) idxGuru = headers.indexOf('nama guru');
    if (idxGuru === -1) idxGuru = 4; 
    let idxMapel = headers.indexOf('mapel'); if (idxMapel === -1) idxMapel = 5; 
    let idxKelas = headers.indexOf('kelas'); if (idxKelas === -1) idxKelas = 6; 
    let idxStatus = headers.indexOf('status'); if (idxStatus === -1) idxStatus = 7;

    let mapingMap = new Map();

    for(let i = 1; i < data.length; i++) {
      let row = data[i];
      let taRow = row[idxTA] ? row[idxTA].toString().trim().toLowerCase() : '';
      let guruName = row[idxGuru] ? row[idxGuru].toString().trim() : '';
      let mapelName = row[idxMapel] ? row[idxMapel].toString().trim() : '-';
      let kelasRaw = row[idxKelas] ? row[idxKelas].toString().trim() : '-';
      let statusRow = row[idxStatus] ? row[idxStatus].toString().trim() : '';

      if(guruName !== '') {
        if (allowedTAs.length > 0 && taRow !== "") { if(!allowedTAs.includes(taRow)) continue; }
        if (statusRow !== "") { if(statusRow.toLowerCase() !== 'aktif') continue; }
        
        let kelasArray = kelasRaw.split(',').map(k => k.trim()).filter(k => k !== '');
        if (kelasArray.length === 0) kelasArray = ['-'];

        let uniqueKey = guruName + "_" + mapelName;
        
        if (!mapingMap.has(uniqueKey)) {
          mapingMap.set(uniqueKey, { guru: guruName, mapel: mapelName, kelasList: new Set(kelasArray) });
        } else {
          kelasArray.forEach(k => mapingMap.get(uniqueKey).kelasList.add(k));
        }
      }
    }
    
    let finalData = [];
    mapingMap.forEach(value => {
        finalData.push({ guru: value.guru, mapel: value.mapel, kelasList: Array.from(value.kelasList).sort() });
    });

    return { status: 'success', data: finalData };
  } catch (error) {
    return { status: 'error', message: error.message };
  }
}

function getTugasManajemenData() {
  try {
    const sheet = checkAndCreateSheetTugas();
    const data = sheet.getDataRange().getValues();
    if(data.length <= 1) return { status: 'success', data: [] };

    let listTugas = [];
    for(let i = 1; i < data.length; i++) {
      let row = data[i];
      if(!row[0]) continue;
      listTugas.push({
        id: row[0],
        guru: row[3], mapel: row[4], judul: row[5], kelas: row[6],
        tglMulai: row[7], jamMulai: row[8], tglSelesai: row[9], jamSelesai: row[10],
        instruksi: row[11], fotoBase64: row[12], reqTeks: row[13], reqFoto: row[14], reqVideo: row[15], reqFile: row[16],
        status: row[17] || 'Buka'
      });
    }
    return { status: 'success', data: listTugas.reverse() };
  } catch (error) {
    return { status: 'error', message: error.message };
  }
}

function simpanDataTugasBaru(payload) {
  try {
    const sheet = checkAndCreateSheetTugas();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let activeTA = "";
    const sheetTA = ss.getSheetByName('Tahun_Ajaran');
    if (sheetTA) {
      const dataTA = sheetTA.getDataRange().getValues();
      for (let i = 1; i < dataTA.length; i++) {
        let isAktif = dataTA[i].some(cell => cell.toString().toLowerCase().trim() === 'aktif');
        if (isAktif) { activeTA = dataTA[i][0] ? dataTA[i][0].toString().trim() : ""; break; }
      }
    }

    const idTugas = 'TGS-' + new Date().getTime() + '-' + Math.floor(Math.random() * 1000);
    const extractDate = (isoStr) => isoStr ? isoStr.split('T')[0] : '';
    const extractTime = (isoStr) => isoStr ? isoStr.split('T')[1] : '';

    const rowData = [
      idTugas, "'" + Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss"), activeTA, 
      payload.guru, payload.mapel, payload.judul, payload.kelas,
      "'" + extractDate(payload.mulai), "'" + extractTime(payload.mulai),
      "'" + extractDate(payload.selesai), "'" + extractTime(payload.selesai),
      payload.instruksi, payload.fotoBase64, payload.reqTeks ? true : false, payload.reqFoto ? true : false, payload.reqVideo ? true : false, payload.reqFile ? true : false,
      "Buka"
    ];

    sheet.appendRow(rowData);
    SpreadsheetApp.flush(); 
    return { 
      status: 'success', 
      message: 'Tugas berhasil disimpan.',
      data: {
        id: idTugas,
        guru: payload.guru,
        mapel: payload.mapel,
        judul: payload.judul,
        kelas: payload.kelas,
        tglMulai: extractDate(payload.mulai),
        jamMulai: extractTime(payload.mulai),
        tglSelesai: extractDate(payload.selesai),
        jamSelesai: extractTime(payload.selesai),
        instruksi: payload.instruksi,
        fotoBase64: payload.fotoBase64,
        reqTeks: payload.reqTeks,
        reqFoto: payload.reqFoto,
        reqVideo: payload.reqVideo,
        reqFile: payload.reqFile,
        status: "Buka"
      }
    };
  } catch (error) { return { status: 'error', message: error.message }; }
}

function updateDataTugasBackend(payload) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tugas');
    if(!sheet) throw new Error("Sheet Tugas tidak ditemukan.");

    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    for(let i = 1; i < data.length; i++) {
      if(data[i][0] === payload.idTugasEdit) { rowIndex = i + 1; break; }
    }
    if(rowIndex === -1) throw new Error("ID Tugas tidak ditemukan.");

    const extractDate = (isoStr) => isoStr ? isoStr.split('T')[0] : '';
    const extractTime = (isoStr) => isoStr ? isoStr.split('T')[1] : '';

    // Ambil status asli dari spreadsheet agar tidak hilang saat edit
    const existingStatus = sheet.getRange(rowIndex, 18).getValue() || "Buka";

    sheet.getRange(rowIndex, 4).setValue(payload.guru);
    sheet.getRange(rowIndex, 5).setValue(payload.mapel);
    sheet.getRange(rowIndex, 6).setValue(payload.judul);
    sheet.getRange(rowIndex, 7).setValue(payload.kelas);
    sheet.getRange(rowIndex, 8).setValue("'" + extractDate(payload.mulai));
    sheet.getRange(rowIndex, 9).setValue("'" + extractTime(payload.mulai));
    sheet.getRange(rowIndex, 10).setValue("'" + extractDate(payload.selesai));
    sheet.getRange(rowIndex, 11).setValue("'" + extractTime(payload.selesai));
    sheet.getRange(rowIndex, 12).setValue(payload.instruksi);
    
    if(payload.fotoBase64 && payload.fotoBase64.trim() !== '') { sheet.getRange(rowIndex, 13).setValue(payload.fotoBase64); }

    sheet.getRange(rowIndex, 14).setValue(payload.reqTeks ? true : false);
    sheet.getRange(rowIndex, 15).setValue(payload.reqFoto ? true : false);
    sheet.getRange(rowIndex, 16).setValue(payload.reqVideo ? true : false);
    sheet.getRange(rowIndex, 17).setValue(payload.reqFile ? true : false);

    SpreadsheetApp.flush(); 
    return { 
      status: 'success', 
      message: 'Tugas berhasil diperbarui.',
      data: {
        id: payload.idTugasEdit,
        guru: payload.guru,
        mapel: payload.mapel,
        judul: payload.judul,
        kelas: payload.kelas,
        tglMulai: extractDate(payload.mulai),
        jamMulai: extractTime(payload.mulai),
        tglSelesai: extractDate(payload.selesai),
        jamSelesai: extractTime(payload.selesai),
        instruksi: payload.instruksi,
        fotoBase64: payload.fotoBase64,
        reqTeks: payload.reqTeks,
        reqFoto: payload.reqFoto,
        reqVideo: payload.reqVideo,
        reqFile: payload.reqFile,
        status: existingStatus
      }
    };
  } catch (error) { return { status: 'error', message: error.message }; }
}

function hapusDataTugasBackend(idTugas) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tugas');
    if(!sheet) throw new Error("Sheet Tugas tidak ditemukan.");
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    for(let i = 1; i < data.length; i++) { if(data[i][0] === idTugas) { rowIndex = i + 1; break; } }
    if(rowIndex === -1) throw new Error("ID Tugas tidak ditemukan.");
    sheet.deleteRow(rowIndex);
    SpreadsheetApp.flush(); 
    return { status: 'success', message: 'Data tugas berhasil dihapus.' };
  } catch (error) { return { status: 'error', message: error.message }; }
}

function toggleStatusTugasBackend(idTugas, nextStatus) {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Tugas');
    if(!sheet) throw new Error("Sheet Tugas tidak ditemukan.");
    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    for(let i = 1; i < data.length; i++) { if(data[i][0] === idTugas) { rowIndex = i + 1; break; } }
    if(rowIndex === -1) throw new Error("ID Tugas tidak ditemukan.");
    sheet.getRange(rowIndex, 18).setValue(nextStatus); 
    SpreadsheetApp.flush(); 
    return { status: 'success', message: 'Status berhasil diperbarui.' };
  } catch (error) { return { status: 'error', message: error.message }; }
}

function getTugasSiswaList(currentUser) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let activeTA = "";
    const sheetTA = ss.getSheetByName('Tahun_Ajaran');
    if (sheetTA) {
      const dataTA = sheetTA.getDataRange().getValues();
      for (let i = 1; i < dataTA.length; i++) {
        let isAktif = dataTA[i].some(cell => cell.toString().toLowerCase().trim() === 'aktif');
        if (isAktif) { activeTA = dataTA[i][0] ? dataTA[i][0].toString().trim() : ""; break; }
      }
    }

    let siswaKelas = ""; 
    let matchedIdUser = ""; 
    let curId = (currentUser.id_user || currentUser.id_siswa || "").toString().trim().toLowerCase();
    let curUser = (currentUser.username || "").toString().trim().toLowerCase();
    let curNama = (currentUser.nama || "").toString().trim().toLowerCase();
    
    const sheetSiswa = ss.getSheetByName('Data_Siswa') || ss.getSheetByName('Siswa');
    if (sheetSiswa) {
      const dataSiswa = sheetSiswa.getDataRange().getValues();
      if (dataSiswa.length > 0) {
        const headersSiswa = dataSiswa[0].map(h => h.toString().toLowerCase().trim());
        let idxIdUser = headersSiswa.indexOf('id_user'); 
        let idxIdSiswa = headersSiswa.indexOf('id_siswa'); 
        let idxUsername = headersSiswa.indexOf('username');
        let idxNama = headersSiswa.findIndex(h => h === 'nama' || h === 'nama_siswa' || h === 'nama siswa' || h === 'nama lengkap');
        let idxIdTa = headersSiswa.indexOf('id_ta');
        let idxKelas = headersSiswa.findIndex(h => h === 'kelas' || h === 'rombel'); 
        if (idxKelas === -1) idxKelas = 3;
        
        for(let i = 1; i < dataSiswa.length; i++) {
          let matchFound = false;
          let rowIdUser = idxIdUser !== -1 && dataSiswa[i][idxIdUser] ? dataSiswa[i][idxIdUser].toString().trim().toLowerCase() : "";
          let rowIdSiswa = idxIdSiswa !== -1 && dataSiswa[i][idxIdSiswa] ? dataSiswa[i][idxIdSiswa].toString().trim().toLowerCase() : "";
          let rowUsername = idxUsername !== -1 && dataSiswa[i][idxUsername] ? dataSiswa[i][idxUsername].toString().trim().toLowerCase() : "";
          let rowNama = idxNama !== -1 && dataSiswa[i][idxNama] ? dataSiswa[i][idxNama].toString().trim().toLowerCase() : "";

          if (curId !== "" && (curId === rowIdUser || curId === rowIdSiswa)) matchFound = true;
          else if (curUser !== "" && curUser === rowUsername) matchFound = true;
          else if (curNama !== "" && curNama === rowNama) matchFound = true;

          if (matchFound) {
            if (idxIdTa !== -1) {
               let rowIdTa = dataSiswa[i][idxIdTa] ? dataSiswa[i][idxIdTa].toString().trim() : "";
               if (activeTA !== "" && rowIdTa !== "" && rowIdTa !== activeTA) continue; 
            }
            siswaKelas = dataSiswa[i][idxKelas] ? dataSiswa[i][idxKelas].toString().trim() : "";
            matchedIdUser = (rowIdUser || rowIdSiswa || rowUsername || rowNama); 
            break; 
          }
        }
      }
    }

    if(!matchedIdUser) matchedIdUser = curId || curUser || curNama;

    const sheetKumpul = checkAndCreateSheetKumpulTugas();
    const dataKumpul = sheetKumpul.getDataRange().getValues();
    const headersKumpul = dataKumpul[0].map(h => h.toString().toLowerCase().trim());
    
    const kTugasIdx = headersKumpul.indexOf('id_tugas') !== -1 ? headersKumpul.indexOf('id_tugas') : 1;
    const kNisIdx = headersKumpul.indexOf('nis_id') !== -1 ? headersKumpul.indexOf('nis_id') : 2;
    const kNamaIdx = headersKumpul.indexOf('nama_siswa') !== -1 ? headersKumpul.indexOf('nama_siswa') : headersKumpul.indexOf('nama');
    const kWaktuIdx = headersKumpul.indexOf('waktu_kumpul') !== -1 ? headersKumpul.indexOf('waktu_kumpul') : 5;
    const kTeksIdx = headersKumpul.indexOf('jawaban_teks') !== -1 ? headersKumpul.indexOf('jawaban_teks') : 6;
    const kFotoIdx = headersKumpul.indexOf('file_foto') !== -1 ? headersKumpul.indexOf('file_foto') : 7;
    const kVideoIdx = headersKumpul.indexOf('file_video') !== -1 ? headersKumpul.indexOf('file_video') : 8;
    const kFileIdx = headersKumpul.indexOf('file_dokumen') !== -1 ? headersKumpul.indexOf('file_dokumen') : 9;
    const kNilaiIdx = headersKumpul.indexOf('nilai') !== -1 ? headersKumpul.indexOf('nilai') : 10;
    const kCatatanIdx = headersKumpul.indexOf('catatan_guru') !== -1 ? headersKumpul.indexOf('catatan_guru') : 11;

    let kumpulMap = {};
    for(let j=1; j<dataKumpul.length; j++) {
      let r = dataKumpul[j];
      let rowNis = "";
      if (kNisIdx !== -1 && r[kNisIdx]) rowNis = r[kNisIdx].toString().toLowerCase().trim();
      else if (kNamaIdx !== -1 && r[kNamaIdx]) rowNis = r[kNamaIdx].toString().toLowerCase().trim();
      
      if (rowNis === matchedIdUser.toLowerCase() || rowNis === curId || rowNis === curUser || rowNis === curNama) {
        kumpulMap[r[kTugasIdx]] = { 
          tglKumpul: r[kWaktuIdx],
          teks: r[kTeksIdx],
          foto: r[kFotoIdx],
          video: r[kVideoIdx],
          file: r[kFileIdx],
          nilai: r[kNilaiIdx],
          catatan: r[kCatatanIdx]
        };
      }
    }

    const sheetTugas = checkAndCreateSheetTugas();
    const dataTugas = sheetTugas.getDataRange().getValues();
    let tugasList = [];
    
    for (let i = 1; i < dataTugas.length; i++) {
      let row = dataTugas[i];
      if(!row[0]) continue;
      
      let taTugas = row[2] ? row[2].toString().trim() : "";
      let kelasTugas = row[6] ? row[6].toString().trim() : "";
      let statusTugas = row[17] ? row[17].toString().trim() : "Buka";
      
      if (taTugas === activeTA && kelasTugas === siswaKelas && statusTugas === "Buka") {
        let jns = [];
        let reqTeks = (row[13] === true || row[13] === "TRUE");
        let reqFoto = (row[14] === true || row[14] === "TRUE");
        let reqVideo = (row[15] === true || row[15] === "TRUE");
        let reqFile = (row[16] === true || row[16] === "TRUE");
        
        if(reqTeks) jns.push("teks");
        if(reqFoto) jns.push("foto");
        if(reqVideo) jns.push("video");
        if(reqFile) jns.push("file");

        let tglSelRaw = row[9];
        let jamSelRaw = row[10];
        let tglSel = "";
        let jamSel = "23:59";
        
        if (tglSelRaw instanceof Date) {
            let yyyy = tglSelRaw.getFullYear();
            let mm = String(tglSelRaw.getMonth() + 1).padStart(2, '0');
            let dd = String(tglSelRaw.getDate()).padStart(2, '0');
            tglSel = `${yyyy}-${mm}-${dd}`;
        } else {
            tglSel = tglSelRaw ? tglSelRaw.toString().replace(/'/g, "").trim() : "";
        }
        
        if (jamSelRaw instanceof Date) {
            let hh = String(jamSelRaw.getHours()).padStart(2, '0');
            let mmin = String(jamSelRaw.getMinutes()).padStart(2, '0');
            jamSel = `${hh}:${mmin}`;
        } else if (jamSelRaw) {
            jamSel = jamSelRaw.toString().replace(/'/g, "").trim();
        }
        
        let kData = kumpulMap[row[0]];
        let isLengkap = true;
        let infoKekurangan = [];
        
        if(kData) {
            if(reqTeks && !kData.teks) { isLengkap = false; infoKekurangan.push("Teks"); }
            if(reqFoto && !kData.foto) { isLengkap = false; infoKekurangan.push("Foto"); }
            if(reqVideo && !kData.video) { isLengkap = false; infoKekurangan.push("Video"); }
            if(reqFile && !kData.file) { isLengkap = false; infoKekurangan.push("File"); }
        } else {
            isLengkap = false;
        }

        tugasList.push({
          id: row[0], guru: row[3], mapel: row[4], judul: row[5], kelas: row[6],
          instruksi: row[11], fotoInstruksi: row[12], jenisTugas: jns.join(","),
          selesai: tglSel ? `${tglSel}T${jamSel}` : null,
          progress: kData ? 100 : 0, 
          kumpul: kData || null, 
          isLengkap: isLengkap,
          infoKekurangan: infoKekurangan.join(", ")
        });
      }
    }

    return { 
      status: 'success', 
      data: {
        nis: matchedIdUser, 
        kelas: siswaKelas,
        tugas: tugasList.reverse()
      } 
    };
  } catch(error) {
    return { status: 'error', message: error.message };
  }
}

function saveToDriveHelper(base64Data, mimeType, fileName, folderName) {
  try {
    let folder;
    let folders = DriveApp.getFoldersByName(folderName);
    if (folders.hasNext()) {
        folder = folders.next();
    } else {
        folder = DriveApp.createFolder(folderName);
    }
    
    let cleanBase64 = base64Data.replace(/^data:.*,/, '');
    let decodedData = Utilities.base64Decode(cleanBase64);
    let blob = Utilities.newBlob(decodedData, mimeType, fileName);
    let file = folder.createFile(blob);
    
    try {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch(shareError) {
        Logger.log("Peringatan: Gagal set sharing link publik. " + shareError.message);
    }
    
    return String(file.getUrl()); 
  } catch(e) {
    throw new Error("Gagal menyimpan file ke Google Drive. Detail: " + e.message);
  }
}

function submitJawabanTugasSiswa(payload, currentUser) {
  try {
    const sheetKumpul = checkAndCreateSheetKumpulTugas();
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    let namaGuruTarget = "";
    let mapelTarget = "";
    const sheetTugas = ss.getSheetByName('Tugas');
    if (sheetTugas) {
      const dataTugas = sheetTugas.getDataRange().getValues();
      for (let i = 1; i < dataTugas.length; i++) {
        if (dataTugas[i][0] === payload.idTugas) {
          namaGuruTarget = dataTugas[i][3] ? dataTugas[i][3].toString() : ""; 
          mapelTarget = dataTugas[i][4] ? dataTugas[i][4].toString() : "";    
          break;
        }
      }
    }

    const idKumpul = 'KPL-' + new Date().getTime();
    const waktuKumpul = "'" + Utilities.formatDate(new Date(), "Asia/Jakarta", "yyyy-MM-dd HH:mm:ss");
    
    let videoUrl = "";
    if (payload.videoBase64 && payload.videoBase64 !== "") {
        let vName = payload.videoName || `Video_${payload.nis}_${payload.idTugas}`;
        videoUrl = saveToDriveHelper(payload.videoBase64, payload.videoMime, vName, "LMS_Video_Siswa");
    }
    
    let fileUrl = "";
    if (payload.fileBase64 && payload.fileBase64 !== "") {
        let fName = payload.fileName || `Dokumen_${payload.nis}_${payload.idTugas}`;
        fileUrl = saveToDriveHelper(payload.fileBase64, payload.fileMime, fName, "LMS_Dokumen_Siswa");
    }

    const lc = sheetKumpul.getLastColumn();
    const headersRaw = lc > 0 ? sheetKumpul.getRange(1, 1, 1, lc).getValues()[0] : [];
    const headers = headersRaw.map(h => h.toString().toLowerCase().trim());
    
    let rowData = new Array(headers.length > 0 ? headers.length : 14).fill("");
    
    const setVal = (colName, val) => {
        let idx = headers.indexOf(colName.toLowerCase());
        if (idx !== -1) {
            rowData[idx] = val; 
        }
    };

    setVal('id_kumpul', idKumpul);
    setVal('id_tugas', payload.idTugas);
    setVal('nama_guru', namaGuruTarget); 
    setVal('mapel', mapelTarget);        
    setVal('nis_id', payload.nis);
    setVal('nama_siswa', currentUser.nama); 
    setVal('kelas', payload.kelas);
    setVal('waktu_kumpul', waktuKumpul);
    setVal('jawaban_teks', payload.teks || "");
    setVal('file_foto', payload.fotoBase64 || ""); 
    setVal('file_video', videoUrl); 
    setVal('file_dokumen', fileUrl); 
    setVal('nilai', "");
    setVal('catatan_guru', "");
    
    sheetKumpul.appendRow(rowData);
    SpreadsheetApp.flush(); 
    
    return { status: 'success', message: 'Tugas berhasil dikumpulkan.' };
  } catch(error) {
    return { status: 'error', message: error.message };
  }
}

function getDaftarTugasUntukDinilai(currentUser) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    const sheetTugas = ss.getSheetByName('Tugas');
    if(!sheetTugas) throw new Error("Sheet Tugas tidak ditemukan");
    const dataTugas = sheetTugas.getDataRange().getValues();
    
    const sheetSiswa = ss.getSheetByName('Data_Siswa') || ss.getSheetByName('Siswa');
    let mapKelasCount = {};
    if(sheetSiswa) {
        const dataSiswa = sheetSiswa.getDataRange().getValues();
        let headers = dataSiswa[0].map(h => h.toString().toLowerCase().trim());
        let idxKelas = headers.findIndex(h => h === 'kelas' || h === 'rombel');
        if(idxKelas === -1) idxKelas = 3;
        
        for(let i=1; i<dataSiswa.length; i++) {
            let kls = dataSiswa[i][idxKelas] ? dataSiswa[i][idxKelas].toString().trim() : "";
            if(kls !== "") {
                mapKelasCount[kls] = (mapKelasCount[kls] || 0) + 1;
            }
        }
    }

    const sheetKumpul = checkAndCreateSheetKumpulTugas();
    const dataKumpul = sheetKumpul.getDataRange().getValues();
    const headersKumpul = dataKumpul[0].map(h => h.toString().toLowerCase().trim());
    const kTugasIdx = headersKumpul.indexOf('id_tugas') !== -1 ? headersKumpul.indexOf('id_tugas') : 1;
    const kNilaiIdx = headersKumpul.indexOf('nilai') !== -1 ? headersKumpul.indexOf('nilai') : 10;

    let mapKumpulCount = {}; 
    let mapKoreksiCount = {}; 
    
    for(let j=1; j<dataKumpul.length; j++) {
        let idTugas = dataKumpul[j][kTugasIdx];
        let nilai = dataKumpul[j][kNilaiIdx];
        if(idTugas) {
            mapKumpulCount[idTugas] = (mapKumpulCount[idTugas] || 0) + 1;
            if(nilai === "") mapKoreksiCount[idTugas] = (mapKoreksiCount[idTugas] || 0) + 1;
        }
    }

    let isGuru = currentUser.role.toLowerCase() === 'guru';
    let hasil = [];

    for(let i=1; i<dataTugas.length; i++) {
        let row = dataTugas[i];
        let t_id = row[0]; if(!t_id) continue;
        let t_guru = row[3];
        
        if(isGuru && t_guru !== currentUser.nama) continue; 

        let kls = row[6];
        let totSiswa = mapKelasCount[kls] || 0;
        let totKumpul = mapKumpulCount[t_id] || 0;
        let totBelum = Math.max(0, totSiswa - totKumpul);
        let totKoreksi = mapKoreksiCount[t_id] || 0;

        hasil.push({
            id: t_id,
            judul: row[5],
            mapel: row[4],
            kelas: kls,
            tglBuat: row[1] ? row[1].toString().replace(/'/g, "") : "-",
            tglMulai: (row[7] || "") + " " + (row[8] || ""),
            tglSelesai: (row[9] || "") + " " + (row[10] || ""),
            totSiswa: totSiswa,
            totKumpul: totKumpul,
            totBelum: totBelum,
            totKoreksi: totKoreksi
        });
    }

    return { status: 'success', data: hasil.reverse() };
  } catch(error) {
    return { status: 'error', message: error.message };
  }
}

function getDetailKumpulanTugas(idTugas) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetTugas = ss.getSheetByName('Tugas');
    const dataTugas = sheetTugas.getDataRange().getValues();
    let tugasInfo = null;
    let reqTeks = false, reqFoto = false, reqVideo = false, reqFile = false;
    let tglDibuat = "";

    for(let i=1; i<dataTugas.length; i++) {
        if(dataTugas[i][0] === idTugas) {
            tugasInfo = dataTugas[i];
            tglDibuat = (tugasInfo[1] || "").toString().replace(/'/g,"");
            reqTeks = (tugasInfo[13] === true || tugasInfo[13] === "TRUE");
            reqFoto = (tugasInfo[14] === true || tugasInfo[14] === "TRUE");
            reqVideo = (tugasInfo[15] === true || tugasInfo[15] === "TRUE");
            reqFile = (tugasInfo[16] === true || tugasInfo[16] === "TRUE");
            break;
        }
    }
    if(!tugasInfo) throw new Error("Tugas tidak ditemukan");

    let targetKelas = tugasInfo[6];
    const sheetSiswa = ss.getSheetByName('Data_Siswa') || ss.getSheetByName('Siswa');
    let siswaList = [];
    if(sheetSiswa) {
        const dataSiswa = sheetSiswa.getDataRange().getValues();
        let headers = dataSiswa[0].map(h => h.toString().toLowerCase().trim());
        let idxKelas = headers.findIndex(h => h === 'kelas' || h === 'rombel'); if(idxKelas===-1) idxKelas=3;
        let idxNama = headers.findIndex(h => h === 'nama' || h === 'nama_siswa' || h === 'nama siswa');
        let idxIdUser = headers.indexOf('id_user'); if(idxIdUser===-1) idxIdUser=0;
        let idxNis = headers.findIndex(h => h === 'nis' || h === 'nisn' || h === 'no_induk');

        for(let i=1; i<dataSiswa.length; i++) {
            let kls = dataSiswa[i][idxKelas] ? dataSiswa[i][idxKelas].toString().trim() : "";
            if(kls === targetKelas) {
                siswaList.push({
                    id_user: dataSiswa[i][idxIdUser],
                    nis_display: idxNis !== -1 && dataSiswa[i][idxNis] ? dataSiswa[i][idxNis] : (dataSiswa[i][idxIdUser] || "-"),
                    nama: dataSiswa[i][idxNama],
                    kelas: kls
                });
            }
        }
    }

    const sheetKumpul = checkAndCreateSheetKumpulTugas();
    const dataKumpul = sheetKumpul.getDataRange().getValues();
    const headersKumpul = dataKumpul[0].map(h => h.toString().toLowerCase().trim());
    
    const kIdIdx = headersKumpul.indexOf('id_kumpul') !== -1 ? headersKumpul.indexOf('id_kumpul') : 0;
    const kTugasIdx = headersKumpul.indexOf('id_tugas') !== -1 ? headersKumpul.indexOf('id_tugas') : 1;
    const kNisIdx = headersKumpul.indexOf('nis_id') !== -1 ? headersKumpul.indexOf('nis_id') : 2;
    const kNamaIdx = headersKumpul.indexOf('nama_siswa') !== -1 ? headersKumpul.indexOf('nama_siswa') : headersKumpul.indexOf('nama');
    const kWaktuIdx = headersKumpul.indexOf('waktu_kumpul') !== -1 ? headersKumpul.indexOf('waktu_kumpul') : 5;
    const kTeksIdx = headersKumpul.indexOf('jawaban_teks') !== -1 ? headersKumpul.indexOf('jawaban_teks') : 6;
    const kFotoIdx = headersKumpul.indexOf('file_foto') !== -1 ? headersKumpul.indexOf('file_foto') : 7;
    const kVideoIdx = headersKumpul.indexOf('file_video') !== -1 ? headersKumpul.indexOf('file_video') : 8;
    const kFileIdx = headersKumpul.indexOf('file_dokumen') !== -1 ? headersKumpul.indexOf('file_dokumen') : 9;
    const kNilaiIdx = headersKumpul.indexOf('nilai') !== -1 ? headersKumpul.indexOf('nilai') : 10;
    const kCatatanIdx = headersKumpul.indexOf('catatan_guru') !== -1 ? headersKumpul.indexOf('catatan_guru') : 11;

    let kumpulMap = {};
    for(let j=1; j<dataKumpul.length; j++) {
        if(dataKumpul[j][kTugasIdx] === idTugas) {
            let namaRaw = dataKumpul[j][kNamaIdx] ? dataKumpul[j][kNamaIdx].toString().toLowerCase() : "";
            let nisRaw = dataKumpul[j][kNisIdx] ? dataKumpul[j][kNisIdx].toString().toLowerCase() : namaRaw;
            
            kumpulMap[nisRaw] = {
                idKumpul: dataKumpul[j][kIdIdx],
                waktu: dataKumpul[j][kWaktuIdx] ? dataKumpul[j][kWaktuIdx].toString().replace(/'/g,"") : "-",
                teks: dataKumpul[j][kTeksIdx], foto: dataKumpul[j][kFotoIdx], video: dataKumpul[j][kVideoIdx], file: dataKumpul[j][kFileIdx],
                nilai: dataKumpul[j][kNilaiIdx], catatan: dataKumpul[j][kCatatanIdx]
            };
        }
    }

    let finalData = [];
    siswaList.forEach(s => {
        let kData = kumpulMap[s.id_user.toString().toLowerCase()] || kumpulMap[s.nama.toString().toLowerCase()];
        
        let statusKumpul = "Belum Mengumpulkan";
        let isLengkap = true;
        if(kData) {
            statusKumpul = "Sudah Kumpul";
            if(reqTeks && !kData.teks) isLengkap = false;
            if(reqFoto && !kData.foto) isLengkap = false;
            if(reqVideo && !kData.video) isLengkap = false;
            if(reqFile && !kData.file) isLengkap = false;
            if(!isLengkap) statusKumpul = "Kurang Lengkap";
        }

        finalData.push({
            idKumpul: kData ? kData.idKumpul : "",
            id_user: s.id_user,
            nis: s.nis_display,
            nama: s.nama,
            kelas: s.kelas,
            tugasJudul: tugasInfo[5],
            waktuBuat: tglDibuat,
            waktuKumpul: kData ? kData.waktu : "-",
            status: statusKumpul,
            nilai: kData ? kData.nilai : "",
            catatan: kData ? kData.catatan : "",
            jawabanTeks: kData ? kData.teks : "",
            jawabanFoto: kData ? kData.foto : "",
            jawabanVideo: kData ? kData.video : "",
            jawabanFile: kData ? kData.file : ""
        });
    });

    return { status: 'success', data: finalData };
  } catch(error) {
    return { status: 'error', message: error.message };
  }
}

function simpanNilaiTugas(payload) {
    try {
        const sheetKumpul = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Kumpul_Tugas');
        if(!sheetKumpul) throw new Error("Sheet Kumpul_Tugas tidak ditemukan");
        
        const data = sheetKumpul.getDataRange().getValues();
        const headers = data[0].map(h => h.toString().toLowerCase().trim());
        const idxId = headers.indexOf('id_kumpul') !== -1 ? headers.indexOf('id_kumpul') : 0;
        
        const idxNilai = headers.indexOf('nilai') !== -1 ? headers.indexOf('nilai') + 1 : 11;
        const idxCatatan = headers.indexOf('catatan_guru') !== -1 ? headers.indexOf('catatan_guru') + 1 : 12;

        let rowIndex = -1;
        for(let i=1; i<data.length; i++) {
            if(data[i][idxId] === payload.idKumpul) { rowIndex = i + 1; break; }
        }
        if(rowIndex === -1) throw new Error("Data pengumpulan tidak ditemukan.");

        sheetKumpul.getRange(rowIndex, idxNilai).setValue(payload.nilai); 
        sheetKumpul.getRange(rowIndex, idxCatatan).setValue(payload.catatan); 
        SpreadsheetApp.flush();
        return { status: 'success', message: 'Nilai berhasil disimpan.' };
    } catch(error) {
        return { status: 'error', message: error.message };
    }
}

// FUNGSI BARU: SIMPAN KOREKSI MASSAL SECARA BATCH (SUPER CEPAT)
function simpanNilaiTugasMassal(payload) {
  try {
    const sheetKumpul = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Kumpul_Tugas');
    if(!sheetKumpul) throw new Error("Sheet Kumpul_Tugas tidak ditemukan");
    
    // Ambil semua data sekaligus ke RAM (Batch Processing)
    const data = sheetKumpul.getDataRange().getValues();
    const headers = data[0].map(h => h.toString().toLowerCase().trim());
    
    const idxId = headers.indexOf('id_kumpul');
    const idxNilai = headers.indexOf('nilai');
    const idxCatatan = headers.indexOf('catatan_guru');

    if (idxId === -1 || idxNilai === -1 || idxCatatan === -1) {
        throw new Error("Struktur kolom sheet tidak valid.");
    }

    // Menggunakan Set untuk pencarian ID sangat cepat O(1)
    const idSet = new Set(payload.idKumpulArray);
    let updateCount = 0;

    // Modifikasi data di dalam RAM
    for(let i=1; i<data.length; i++) {
        if(idSet.has(data[i][idxId])) {
            data[i][idxNilai] = payload.nilai;
            data[i][idxCatatan] = payload.catatan;
            updateCount++;
        }
    }

    // Tulis kembali ke Sheet HANYA JIKA ada perubahan (dalam 1 kali aksi Batch Processing)
    if(updateCount > 0) {
        sheetKumpul.getRange(1, 1, data.length, headers.length).setValues(data);
        SpreadsheetApp.flush(); // Eksekusi real-time instan
    }

    return { status: 'success', message: 'Nilai massal berhasil disimpan.' };
  } catch(error) {
    return { status: 'error', message: error.message };
  }
}
