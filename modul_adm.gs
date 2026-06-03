/**
 * File: modul_adm.gs
 * Deskripsi: Modul backend untuk fitur administratif (Jadwal, Absensi, & Jurnal)
 * Dioptimalkan: Pencocokan Data Absen & Jurnal yang Fleksibel dan Tahan Bug (ZettBOT Fix)
 */

function getSheetWithHeaders_(sheetName, requiredHeaders) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(requiredHeaders);
    sheet.getRange(1, 1, 1, requiredHeaders.length).setFontWeight("bold").setBackground("#e0e0e0");
  } else {
    let currentLastCol = Math.max(1, sheet.getLastColumn());
    let headers = sheet.getRange(1, 1, 1, currentLastCol).getValues()[0];
    let needsUpdate = false;
    
    requiredHeaders.forEach(req => {
      let found = headers.some(h => String(h).trim().toLowerCase() === String(req).trim().toLowerCase());
      if (!found) {
        currentLastCol++;
        sheet.getRange(1, currentLastCol).setValue(req);
        needsUpdate = true;
      }
    });
    if (needsUpdate) {
      sheet.getRange(1, 1, 1, currentLastCol).setFontWeight("bold").setBackground("#e0e0e0");
    }
  }
  return sheet;
}

function getHeaderMap_(sheet) {
  let map = {};
  if (sheet.getLastColumn() < 1) return map;
  let headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  headers.forEach((h, i) => { 
      if (h !== null && h !== undefined && String(h).trim() !== '') {
          map[String(h).trim().toLowerCase()] = i; 
      }
  });
  return map;
}

function getActiveTAInternal_() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetTA = ss.getSheetByName('Tahun_Ajaran') || ss.getSheetByName('TahunAjaran') || ss.getSheetByName('TA');
    if(!sheetTA) return null;
    const data = sheetTA.getDataRange().getValues();
    if(!data || data.length === 0) return null;
    let statusCol = -1, idCol = -1;
    for(let c=0; c<data[0].length; c++) {
        let h = String(data[0][c]).toLowerCase();
        if(h.includes('status')) statusCol = c;
        if(h === 'id' || h === 'id_ta') idCol = c;
    }
    if(idCol === -1) {
        for(let c=0; c<data[0].length; c++) {
            let h = String(data[0][c]).toLowerCase();
            if(h.includes('tahun') || h === 'ta') idCol = c;
        }
    }
    if(statusCol !== -1 && idCol !== -1) {
        for(let i=1; i<data.length; i++) {
            if(String(data[i][statusCol]).trim().toLowerCase() === 'aktif') {
                return String(data[i][idCol]).trim();
            }
        }
    }
  } catch(e) {}
  return null;
}

function getIdsByGuruName_(namaTarget) {
  if (!namaTarget) return { idFound: "" };
  try {
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      
      const searchInSheet = (sheetName) => {
          let sheet = ss.getSheetByName(sheetName);
          if (!sheet) return "";
          let data = sheet.getDataRange().getValues();
          if (data.length <= 1) return "";
          
          let headers = data[0].map(h => String(h).toLowerCase().trim());
          
          let idCol = headers.findIndex(h => h === 'id' || h === 'id user' || h === 'id pengguna' || h === 'id guru' || h === 'id_user');
          if(idCol === -1) idCol = 0; 
          
          let namaCol = headers.findIndex(h => h === 'nama' || h === 'nama lengkap' || h === 'nama pengguna' || h === 'nama guru' || h === 'nama_lengkap');
          if(namaCol === -1) {
              let maybeNama = headers.findIndex(h => h.includes('nama'));
              namaCol = maybeNama !== -1 ? maybeNama : 2; 
          }
          
          for(let i=1; i<data.length; i++) {
              if (String(data[i][namaCol]).trim().toLowerCase() === String(namaTarget).trim().toLowerCase()) {
                  return String(data[i][idCol]).trim();
              }
          }
          return "";
      };

      let foundId = searchInSheet('DataPengguna') || searchInSheet('Data_Pengguna') || searchInSheet('Pengguna');
      if (foundId) return { idFound: foundId };
      
      foundId = searchInSheet('DataGuru') || searchInSheet('Data_Guru') || searchInSheet('Guru');
      if (foundId) return { idFound: foundId };
      
  } catch(e) {
      console.error("Error getIdsByGuruName_:", e);
  }
  return { idFound: "" };
}

function getMapelFromMaping_(guruName, kelasTarget, taAktif) {
    if(!guruName) return "";
    try {
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Maping_Guru');
        if(!sheet) return "";
        const data = sheet.getDataRange().getValues();
        let foundMapels = new Set();
        
        for(let i=1; i<data.length; i++) {
            let rowTA = String(data[i][1] || '').trim();
            let rowGuru = String(data[i][4] || '').trim();
            let rowMapel = String(data[i][5] || '').trim();
            let rowKelasStr = String(data[i][6] || '').trim();
            
            if ((!taAktif || rowTA === taAktif) && rowGuru.toLowerCase() === guruName.toLowerCase()) {
                if (kelasTarget) {
                    let kelasArr = rowKelasStr.split(',').map(k => k.trim());
                    if(kelasArr.includes(kelasTarget) && rowMapel) {
                        foundMapels.add(rowMapel);
                    }
                } else {
                    if(rowMapel) foundMapels.add(rowMapel);
                }
            }
        }
        return Array.from(foundMapels).join(', ');
    } catch(e) {
        return "";
    }
}

function getMapingDataHandler() {
   try {
       const ss = SpreadsheetApp.getActiveSpreadsheet();
       let activeTAId = getActiveTAInternal_();
       let sheetMaping = ss.getSheetByName('Maping_Guru');
       let mapingRecords = [];
       if(sheetMaping) {
           const data = sheetMaping.getDataRange().getValues();
           for(let i=1; i<data.length; i++) {
               let rowTA = String(data[i][1] || '').trim();
               let guruName = String(data[i][4] || '').trim();
               if (activeTAId && rowTA && rowTA !== activeTAId) continue;
               mapingRecords.push({
                   guru: guruName,
                   mapel: String(data[i][5] || '').trim(),
                   kelas: String(data[i][6] || '').trim()
               });
           }
       }
       return { status: 'success', data: mapingRecords };
   } catch(e) { 
       return { status: 'error', message: e.message }; 
   }
}

function formatTgl_(tglRaw) {
  if (tglRaw === null || tglRaw === undefined || tglRaw === "") return "";
  if (tglRaw instanceof Date) {
      if(isNaN(tglRaw.getTime())) return "";
      return Utilities.formatDate(tglRaw, "Asia/Jakarta", "yyyy-MM-dd");
  }
  let tglStr = String(tglRaw).trim();
  if (tglStr.includes("GMT") || tglStr.includes("Waktu") || tglStr.length > 15) {
    let d = new Date(tglStr);
    if (!isNaN(d.getTime())) return Utilities.formatDate(d, "Asia/Jakarta", "yyyy-MM-dd");
  }
  if (tglStr.includes("/")) {
     let p = tglStr.split("/");
     if (p.length === 3) {
         if (p[2].length === 4) return `${p[2]}-${p[1].padStart(2,'0')}-${p[0].padStart(2,'0')}`;
         if (p[0].length === 4) return `${p[0]}-${p[1].padStart(2,'0')}-${p[2].padStart(2,'0')}`;
     }
  }
  return tglStr;
}

function getGuruDashboardData(currentUser) {
  try {
    if (!currentUser) return { status: 'error', message: 'Sesi tidak valid.' };
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let activeTAId = getActiveTAInternal_();

    const now = new Date(new Date().toLocaleString("en-US", {timeZone: "Asia/Jakarta"}));
    const days = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
    const todayName = days[now.getDay()];
    const todayStr = Utilities.formatDate(now, "Asia/Jakarta", "yyyy-MM-dd");

    const sheetJadwal = ss.getSheetByName('DataJadwal');
    if (!sheetJadwal) return { status: 'success', data: { hari: todayName, jadwalHariIni: [] } };

    const mapJadwal = getHeaderMap_(sheetJadwal);
    const dataJadwal = sheetJadwal.getDataRange().getValues();
    let jadwalHariIni = [];

    if(dataJadwal.length > 1) {
      for(let i=1; i<dataJadwal.length; i++) {
        let row = dataJadwal[i];
        let rowTA = String(mapJadwal['tahun ajaran'] !== undefined ? row[mapJadwal['tahun ajaran']] : '').trim();
        if (activeTAId && rowTA && rowTA !== activeTAId) continue;

        let guruName = String(mapJadwal['nama guru'] !== undefined ? row[mapJadwal['nama guru']] : '').trim();
        let idUser = String(mapJadwal['id yang menginput'] !== undefined ? row[mapJadwal['id yang menginput']] : '').trim();
        
        let isMatch = (guruName.toLowerCase() === String(currentUser.nama).trim().toLowerCase()) || (idUser === String(currentUser.id).trim());
        if (!isMatch) continue;

        let hari = String(mapJadwal['hari'] !== undefined ? row[mapJadwal['hari']] : '').trim();
        if (hari.toLowerCase() !== todayName.toLowerCase()) continue;

        jadwalHariIni.push({
          id: String(mapJadwal['id'] !== undefined ? row[mapJadwal['id']] : '').trim(),
          hari: hari,
          ta: rowTA,
          kelas: String(mapJadwal['kelas'] !== undefined ? row[mapJadwal['kelas']] : '').trim(),
          mapel: String(mapJadwal['mapel'] !== undefined ? row[mapJadwal['mapel']] : '').trim(),
          jamMulai: String(mapJadwal['jam mulai'] !== undefined ? row[mapJadwal['jam mulai']] : '').trim(),
          jamSelesai: String(mapJadwal['jam selesai'] !== undefined ? row[mapJadwal['jam selesai']] : '').trim(),
          guru: guruName,
          absenDone: false,
          jurnalDone: false,
          ruangan: "", 
          ruang: "",   
          status: ""   
        });
      }
    }

    // PENGECEKAN ABSENSI (Lebih Tahan Banting)
    const sheetAbsen = ss.getSheetByName('DataAbsensi');
    if (sheetAbsen) {
      const mapAbsen = getHeaderMap_(sheetAbsen);
      const dataAbsen = sheetAbsen.getDataRange().getValues();
      if(dataAbsen.length > 1) {
        for(let i=1; i<dataAbsen.length; i++) {
          let rowTA = String(mapAbsen['tahun ajaran'] !== undefined ? dataAbsen[i][mapAbsen['tahun ajaran']] : '').trim();
          if (activeTAId && rowTA && rowTA !== activeTAId) continue;

          let tglStr = formatTgl_(mapAbsen['tanggal'] !== undefined ? dataAbsen[i][mapAbsen['tanggal']] : '');
          if (tglStr === todayStr) {
            let kls = String(mapAbsen['kelas'] !== undefined ? dataAbsen[i][mapAbsen['kelas']] : '').trim().toLowerCase();
            
            jadwalHariIni.forEach(j => { 
                if(j.kelas.toLowerCase() === kls) {
                    j.absenDone = true; 
                }
            });
          }
        }
      }
    }

    // PENGECEKAN JURNAL (Diperbaiki: Hanya Cek Kesesuaian Kelas & Mapel, Mengabaikan Format Jam)
    const sheetJurnal = ss.getSheetByName('DataJurnal');
    if (sheetJurnal) {
      const mapJurnal = getHeaderMap_(sheetJurnal);
      const dataJurnal = sheetJurnal.getDataRange().getValues();
      if(dataJurnal.length > 1) {
        for(let i=1; i<dataJurnal.length; i++) {
          let rowTA = String(mapJurnal['tahun ajaran'] !== undefined ? dataJurnal[i][mapJurnal['tahun ajaran']] : '').trim();
          if (activeTAId && rowTA && rowTA !== activeTAId) continue;

          let tglStr = formatTgl_(mapJurnal['tanggal'] !== undefined ? dataJurnal[i][mapJurnal['tanggal']] : '');
          if (tglStr === todayStr) {
            let guruName = String(mapJurnal['nama guru'] !== undefined ? dataJurnal[i][mapJurnal['nama guru']] : '').trim();
            let idUser = String(mapJurnal['id yang menginput'] !== undefined ? dataJurnal[i][mapJurnal['id yang menginput']] : '').trim();
            let kls = String(mapJurnal['kelas'] !== undefined ? dataJurnal[i][mapJurnal['kelas']] : '').trim().toLowerCase();
            let mapel = String(mapJurnal['mapel'] !== undefined ? dataJurnal[i][mapJurnal['mapel']] : '').trim().toLowerCase();

            let isMatch = (guruName.toLowerCase() === String(currentUser.nama).trim().toLowerCase()) || (idUser && idUser === String(currentUser.id).trim());
            if (isMatch) {
              jadwalHariIni.forEach(j => {
                // REVISI ZETTBOT: Jika kelas cocok dan mapel cocok (atau mapel kosong di DB lama), anggap Jurnal selesai
                if(j.kelas.toLowerCase() === kls && (j.mapel.toLowerCase() === mapel || mapel === '')) {
                    j.jurnalDone = true;
                }
              });
            }
          }
        }
      }
    }

    // LOGIKA PENONAKTIFAN KARTU JADWAL JIKA ABSEN & JURNAL SELESAI
    jadwalHariIni.forEach(j => {
      j.isDisabled = j.absenDone && j.jurnalDone;
      if (j.isDisabled) {
        j.status = "Selesai";
      }
    });

    return { status: 'success', data: { hari: todayName, jadwalHariIni: jadwalHariIni } };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

function getJadwalDataHandler(currentUser) {
  try {
    if (!currentUser) return { status: 'error', message: 'Sesi tidak valid.' };
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let activeTAId = getActiveTAInternal_();

    const sheetMaping = ss.getSheetByName('Maping_Guru');
    const mapingRecords = [];
    if (sheetMaping) {
        const dataMapingRaw = sheetMaping.getDataRange().getValues();
        if(dataMapingRaw.length > 1) {
          for(let i=1; i<dataMapingRaw.length; i++) {
            let rowTA = String(dataMapingRaw[i][1] || '').trim();
            let guruName = String(dataMapingRaw[i][4] || '').trim();
            if (activeTAId && rowTA && rowTA !== activeTAId) continue; 
            if(currentUser.role && currentUser.role.toLowerCase() === 'guru' && guruName.toLowerCase() !== String(currentUser.nama).trim().toLowerCase()) continue;
            
            mapingRecords.push({
              guru: guruName, 
              mapel: String(dataMapingRaw[i][5] || '').trim(),  
              kelas: String(dataMapingRaw[i][6] || '').trim(), 
              ta: rowTA      
            });
          }
        }
    }
    
    const sheetJadwal = ss.getSheetByName('DataJadwal');
    const jadwalRecords = [];
    if (sheetJadwal) {
        const map = getHeaderMap_(sheetJadwal);
        const dataJadwalRaw = sheetJadwal.getDataRange().getValues();
        
        let cId = map['id']; let cGuru = map['nama guru']; let cHari = map['hari'];
        let cKelas = map['kelas']; let cMapel = map['mapel']; let cMulai = map['jam mulai'];
        let cSelesai = map['jam selesai']; let cTA = map['tahun ajaran']; let cIdUser = map['id yang menginput'];

        if(dataJadwalRaw.length > 1) {
          for(let i=1; i<dataJadwalRaw.length; i++) {
            let row = dataJadwalRaw[i];
            let rowIdUser = cIdUser !== undefined ? String(row[cIdUser] || '').trim() : '';
            let guruName = cGuru !== undefined ? String(row[cGuru] || '').trim() : '';
            let rowTA = cTA !== undefined ? String(row[cTA] || '').trim() : ''; 
            
            if (activeTAId && rowTA && rowTA !== activeTAId) continue;
            
            if(currentUser.role && currentUser.role.toLowerCase() === 'guru') {
                let isMatch = (guruName.toLowerCase() === String(currentUser.nama).trim().toLowerCase()) || (rowIdUser && rowIdUser === String(currentUser.id).trim());
                if (!isMatch) continue;
            }
            
            jadwalRecords.push({
              id: String(cId !== undefined ? row[cId] : ''), 
              guru: guruName, 
              hari: String(cHari !== undefined ? row[cHari] : ''), 
              kelas: String(cKelas !== undefined ? row[cKelas] : ''), 
              mapel: String(cMapel !== undefined ? row[cMapel] : ''), 
              jamMulai: String(cMulai !== undefined ? row[cMulai] : ''), 
              jamSelesai: String(cSelesai !== undefined ? row[cSelesai] : ''), 
              ta: rowTA,
              ruangan: "", 
              ruang: "",   
              status: ""   
            });
          }
        }
    }
    return { status: 'success', data: { jadwal: jadwalRecords, maping: mapingRecords } };
  } catch (e) {
    return { status: 'error', message: 'Gagal memuat jadwal: ' + e.message };
  }
}

function saveJadwalRecord(payload, currentUser) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); 
    const requiredHeaders = ['ID', 'Nama Guru', 'Hari', 'Kelas', 'Mapel', 'Jam Mulai', 'Jam Selesai', 'Tahun Ajaran', 'Timestamp', 'ID yang Menginput'];
    const sheet = getSheetWithHeaders_('DataJadwal', requiredHeaders);
    const map = getHeaderMap_(sheet);
    const timestamp = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    let taAktif = getActiveTAInternal_() || "TA Aktif";
    
    let saveGuru = (currentUser.role.toLowerCase() === 'guru') ? currentUser.nama : payload.guru;
    
    let idUserInput = "";
    if (currentUser && currentUser.nama) {
        let fetched = getIdsByGuruName_(currentUser.nama);
        if (fetched.idFound) idUserInput = fetched.idFound;
    }
    if (!idUserInput && currentUser && currentUser.id) idUserInput = currentUser.id;
    if (!idUserInput) idUserInput = (currentUser && currentUser.nama) ? "USER-" + currentUser.nama : "SYSTEM-ADMIN"; 

    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    let idCol = map['id'] !== undefined ? map['id'] : 0;
    
    if(payload.id) {
      for(let i=1; i<data.length; i++) { 
        if(String(data[i][idCol]).trim() === String(payload.id).trim()) { 
          rowIndex = i + 1; break; 
        } 
      }
      if(rowIndex === -1) return { status: 'error', message: 'Data tidak ditemukan.' };
    }
    
    let newRow = new Array(sheet.getLastColumn()).fill("");
    if(rowIndex !== -1) newRow = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    newRow[map['id']] = payload.id || ('JDW-' + new Date().getTime());
    newRow[map['nama guru']] = saveGuru; 
    newRow[map['hari']] = payload.hari;
    newRow[map['kelas']] = payload.kelas;
    newRow[map['mapel']] = payload.mapel; 
    newRow[map['jam mulai']] = payload.jamMulai;
    newRow[map['jam selesai']] = payload.jamSelesai;
    newRow[map['tahun ajaran']] = taAktif;
    newRow[map['timestamp']] = timestamp;
    newRow[map['id yang menginput']] = idUserInput;

    if(rowIndex !== -1) sheet.getRange(rowIndex, 1, 1, newRow.length).setValues([newRow]);
    else sheet.appendRow(newRow);
    
    SpreadsheetApp.flush(); 
    return { status: 'success', message: 'Jadwal berhasil disimpan!' };
  } catch (e) { 
    return { status: 'error', message: e.message.includes('timeout') ? 'Sistem sibuk. Silakan coba lagi.' : e.message }; 
  } finally { lock.releaseLock(); }
}

function deleteJadwalRecord(id) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DataJadwal');
    if(!sheet) return { status: 'error', message: 'Sheet tidak ditemukan.' };
    const data = sheet.getDataRange().getValues();
    let map = getHeaderMap_(sheet);
    let idCol = map['id'] !== undefined ? map['id'] : 0;
    let rowIndex = -1;
    
    for(let i=1; i<data.length; i++) { 
      if(String(data[i][idCol]).trim() === String(id).trim()) { 
        rowIndex = i + 1; break; 
      } 
    }
    
    if(rowIndex !== -1) {
      sheet.deleteRow(rowIndex); 
      SpreadsheetApp.flush(); 
      return { status: 'success', message: 'Jadwal berhasil dihapus.' };
    } else {
      return { status: 'error', message: 'Data tidak ditemukan di database.' };
    }
  } catch (e) { 
    return { status: 'error', message: e.message }; 
  } finally { lock.releaseLock(); }
}

function getJurnalDataHandler(guruFilter, mapelFilter, bulanFilter, currentUser) {
  try {
    if (!currentUser || !currentUser.role) return { status: 'error', message: 'Sesi login tidak valid. Silakan relogin.' };

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let activeTAId = getActiveTAInternal_();

    const sheetJurnal = ss.getSheetByName('DataJurnal');
    if (!sheetJurnal) return { status: 'success', data: [] };

    let lastCol = sheetJurnal.getLastColumn();
    if (lastCol < 1) return { status: 'success', data: [] };

    const mapJur = getHeaderMap_(sheetJurnal);
    const dataJurnal = sheetJurnal.getDataRange().getValues();
    const records = [];
    
    if(dataJurnal.length > 1) {
      let cId = mapJur['id']; let cTgl = mapJur['tanggal']; let cKls = mapJur['kelas'];
      let cGuru = mapJur['nama guru']; let cJam = mapJur['jam ke']; let cMateri = mapJur['uraian materi'];
      let cMapel = mapJur['mapel']; 
      let cTA = mapJur['tahun ajaran']; let cIdUser = mapJur['id yang menginput'];
      let cH = mapJur['hadir']; let cS = mapJur['sakit']; let cI = mapJur['izin']; let cA = mapJur['alpa'];

      for(let i=1; i<dataJurnal.length; i++) {
        let row = dataJurnal[i];
        if (!row || row.length === 0) continue;
        
        let tglStr = formatTgl_(cTgl !== undefined ? row[cTgl] : '');
        let guruName = String(cGuru !== undefined ? row[cGuru] : '').trim();
        let rowMapel = String(cMapel !== undefined ? row[cMapel] : '').trim();
        let rowTA = String(cTA !== undefined ? row[cTA] : '').trim();
        let rowIdUser = String(cIdUser !== undefined ? row[cIdUser] : '').trim();
        
        if (activeTAId && rowTA && rowTA !== activeTAId) continue;
        
        // Pengecekan Filter Role / Parameter Filter
        if(currentUser.role.toLowerCase() === 'guru') {
            let isMatch = (guruName.toLowerCase() === String(currentUser.nama).trim().toLowerCase()) || (rowIdUser && rowIdUser === String(currentUser.id).trim());
            if (!isMatch) continue;
        } else {
            // Admin bisa filter guru
            if (guruFilter && guruName.toLowerCase() !== guruFilter.toLowerCase()) continue;
        }

        if (mapelFilter && rowMapel.toLowerCase() !== mapelFilter.toLowerCase()) continue;
        
        // Filter bulan format "YYYY-MM"
        if (bulanFilter && !tglStr.startsWith(bulanFilter)) continue;
        
        let vJamRaw = cJam !== undefined ? row[cJam] : '';
        if(vJamRaw instanceof Date) {
             let dd = vJamRaw.getDate(); let mm = vJamRaw.getMonth()+1;
             vJamRaw = dd + "-" + mm; 
        }
        let vJam = String(vJamRaw).replace(/^'/, '').trim();
        
        let vH = parseInt(cH !== undefined ? row[cH] : 0); if (isNaN(vH)) vH = 0;
        let vS = parseInt(cS !== undefined ? row[cS] : 0); if (isNaN(vS)) vS = 0;
        let vI = parseInt(cI !== undefined ? row[cI] : 0); if (isNaN(vI)) vI = 0;
        let vA = parseInt(cA !== undefined ? row[cA] : 0); if (isNaN(vA)) vA = 0;

        records.push({
          id: String(cId !== undefined ? row[cId] : ''),
          tanggal: String(tglStr),
          kelas: String(cKls !== undefined ? row[cKls] : '').trim(),
          guru: String(guruName),
          mapel: String(rowMapel), 
          jamKe: vJam,
          materi: String(cMateri !== undefined ? row[cMateri] : ''),
          ta: String(rowTA),
          kehadiran: { H: vH, S: vS, I: vI, A: vA }
        });
      }
    }
    
    records.reverse(); 
    return { status: 'success', data: records };
  } catch (e) {
    return { status: 'error', message: 'Gagal memuat jurnal: ' + e.message };
  }
}

function getKehadiranForJurnal(tanggal, kelas) {
    try {
        let taAktif = getActiveTAInternal_();
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName('DataAbsensi');
        let counts = {H: 0, S: 0, I: 0, A: 0};
        if (!sheet) return { status: 'success', data: counts };

        const data = sheet.getDataRange().getValues();
        let map = getHeaderMap_(sheet);
        
        for(let i=1; i<data.length; i++) {
            let rowTA = map['tahun ajaran'] !== undefined ? String(data[i][map['tahun ajaran']] || '').trim() : '';
            if (taAktif && rowTA && rowTA !== taAktif) continue;
            let rowKelas = map['kelas'] !== undefined ? String(data[i][map['kelas']] || '').trim() : '';
            if (rowKelas !== kelas) continue;
            
            let rowTglStr = formatTgl_(data[i][map['tanggal']]);
            if (rowTglStr === tanggal) {
                let stat = String(data[i][map['status']] || '').trim().toUpperCase();
                if (counts[stat] !== undefined) counts[stat]++;
            }
        }
        return { status: 'success', data: counts };
    } catch(e) { return { status: 'error', message: e.message }; }
}

function saveJurnalRecord(payload, currentUser) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000); 
    const requiredHeaders = ['ID', 'Tanggal', 'Kelas', 'Nama Guru', 'Mapel', 'Jam Ke', 'Uraian Materi', 'Hadir', 'Sakit', 'Izin', 'Alpa', 'Tahun Ajaran', 'Timestamp', 'ID yang Menginput'];
    const sheet = getSheetWithHeaders_('DataJurnal', requiredHeaders);
    const map = getHeaderMap_(sheet);
    const timestamp = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
    let taAktif = getActiveTAInternal_() || "TA Aktif";
    
    let saveGuru = (currentUser.role.toLowerCase() === 'guru') ? currentUser.nama : payload.guru;
    let saveMapel = getMapelFromMaping_(saveGuru, payload.kelas, taAktif);
    
    let idUserInput = "";
    if (currentUser && currentUser.nama) {
        let fetched = getIdsByGuruName_(currentUser.nama);
        if (fetched.idFound) idUserInput = fetched.idFound;
    }
    if (!idUserInput && currentUser && currentUser.id) idUserInput = currentUser.id;
    if (!idUserInput) idUserInput = (currentUser && currentUser.nama) ? "USER-" + currentUser.nama : "SYSTEM-ADMIN";

    const data = sheet.getDataRange().getValues();
    let rowIndex = -1;
    let idCol = map['id'] !== undefined ? map['id'] : 0;
    
    if(payload.id) {
      for(let i=1; i<data.length; i++) { 
        if(String(data[i][idCol]).trim() === String(payload.id).trim()) { 
          rowIndex = i + 1; break; 
        } 
      }
      if(rowIndex === -1) return { status: 'error', message: 'Data tidak ditemukan.' };
    }
    
    let newRow = new Array(sheet.getLastColumn()).fill("");
    if(rowIndex !== -1) newRow = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    newRow[map['id']] = payload.id || ('JRN-' + new Date().getTime());
    newRow[map['tanggal']] = payload.tanggal; 
    newRow[map['kelas']] = payload.kelas;
    newRow[map['nama guru']] = saveGuru;
    newRow[map['mapel']] = saveMapel; 
    newRow[map['jam ke']] = "'" + payload.jamKe; 
    newRow[map['uraian materi']] = payload.materi;
    
    if (payload.kehadiran) {
        newRow[map['hadir']] = parseInt(payload.kehadiran.H) || 0;
        newRow[map['sakit']] = parseInt(payload.kehadiran.S) || 0;
        newRow[map['izin']] = parseInt(payload.kehadiran.I) || 0;
        newRow[map['alpa']] = parseInt(payload.kehadiran.A) || 0;
    } else {
        newRow[map['hadir']] = 0; newRow[map['sakit']] = 0; newRow[map['izin']] = 0; newRow[map['alpa']] = 0;
    }
    
    newRow[map['tahun ajaran']] = taAktif;
    newRow[map['timestamp']] = timestamp;
    newRow[map['id yang menginput']] = idUserInput;

    if(rowIndex !== -1) sheet.getRange(rowIndex, 1, 1, newRow.length).setValues([newRow]);
    else sheet.appendRow(newRow);
    
    SpreadsheetApp.flush(); 
    return { status: 'success', message: 'Jurnal berhasil disimpan!' };
  } catch (e) { 
    return { status: 'error', message: e.message.includes('timeout') ? 'Sistem sibuk. Silakan coba lagi.' : e.message }; 
  } finally { lock.releaseLock(); }
}

function deleteJurnalRecord(id) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DataJurnal');
    if(!sheet) return { status: 'error', message: 'Sheet tidak ditemukan.' };
    const data = sheet.getDataRange().getValues();
    let map = getHeaderMap_(sheet);
    let idCol = map['id'] !== undefined ? map['id'] : 0;
    let rowIndex = -1;
    
    for(let i=1; i<data.length; i++) { 
      if(String(data[i][idCol]).trim() === String(id).trim()) { 
        rowIndex = i + 1; break; 
      } 
    }
    
    if(rowIndex !== -1) {
      sheet.deleteRow(rowIndex); 
      SpreadsheetApp.flush(); 
      return { status: 'success', message: 'Jurnal berhasil dihapus.' };
    } else {
      return { status: 'error', message: 'Data tidak ditemukan di database.' };
    }
  } catch (e) { 
    return { status: 'error', message: e.message }; 
  } finally { lock.releaseLock(); }
}

function getKelasListHandler(currentUser) {
  try {
      let kelasSet = new Set();
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      let taAktif = getActiveTAInternal_();

      if (currentUser && currentUser.role.toLowerCase() === 'guru') {
          let sheetMaping = ss.getSheetByName('Maping_Guru');
          if(sheetMaping) {
              const data = sheetMaping.getDataRange().getValues();
              for(let i=1; i<data.length; i++) {
                  let rowTA = String(data[i][1] || '').trim();
                  let guruName = String(data[i][4] || '').trim();
                  if ((!taAktif || rowTA === taAktif) && guruName.toLowerCase() === String(currentUser.nama).trim().toLowerCase()) {
                      let kelasStr = String(data[i][6]).trim(); 
                      if(kelasStr) kelasStr.split(',').forEach(k => { if(k.trim()) kelasSet.add(k.trim()); });
                  }
              }
          }
      } else {
          let sheet = ss.getSheetByName('DataSiswa') || ss.getSheetByName('Data_Siswa');
          if(sheet) {
              const data = sheet.getDataRange().getValues();
              let m = getHeaderMap_(sheet);
              let kelasCol = m['kelas'] !== undefined ? m['kelas'] : 3; 
              for(let i=1; i<data.length; i++) { let k = String(data[i][kelasCol]).trim(); if(k) kelasSet.add(k); }
          }
      }
      return { status: 'success', data: Array.from(kelasSet).sort() };
  } catch(e) { return { status: 'error', message: e.message }; }
}

function getSiswaByKelas_(kelasFilter) {
   const ss = SpreadsheetApp.getActiveSpreadsheet();
   const sheet = ss.getSheetByName('DataSiswa') || ss.getSheetByName('Data_Siswa');
   if (!sheet) return []; 

   const data = sheet.getDataRange().getValues();
   let m = getHeaderMap_(sheet);
   let nisCol = m['nis'] !== undefined ? m['nis'] : (m['nisn'] !== undefined ? m['nisn'] : 1);
   
   let namaCol = -1;
   if (m['nama_siswa'] !== undefined) namaCol = m['nama_siswa'];
   else if (m['nama lengkap'] !== undefined) namaCol = m['nama lengkap'];
   else if (m['nama'] !== undefined) namaCol = m['nama'];
   else namaCol = 2; 

   let kelasCol = m['kelas'] !== undefined ? m['kelas'] : 3;

   let result = [];
   for(let i=1; i<data.length; i++) {
       let k = String(data[i][kelasCol]).trim();
       if (!kelasFilter || k === kelasFilter) {
           let nis = String(data[i][nisCol]).trim();
           let nama = String(data[i][namaCol]).trim();
           if(nis && nama) result.push({ nis: nis, nama: nama, kelas: k });
       }
   }
   return result;
}

function getAbsensiDataHandler(kelasFilter, guruFilter, mapelFilter, currentUser) {
  try {
      let activeTAId = getActiveTAInternal_();
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      
      let allowedKelas = null;
      if (currentUser && currentUser.role.toLowerCase() === 'guru') {
          allowedKelas = new Set();
          let sheetMaping = ss.getSheetByName('Maping_Guru');
          if(sheetMaping) {
              const data = sheetMaping.getDataRange().getValues();
              for(let i=1; i<data.length; i++) {
                  let rowTA = String(data[i][1] || '').trim();
                  let guruName = String(data[i][4] || '').trim();
                  if ((!activeTAId || rowTA === activeTAId) && guruName.toLowerCase() === String(currentUser.nama).trim().toLowerCase()) {
                      let kelasStr = String(data[i][6]).trim(); 
                      if(kelasStr) kelasStr.split(',').forEach(k => { if(k.trim()) allowedKelas.add(k.trim()); });
                  }
              }
          }
      }

      // 1. Bulk Fetch Data Absensi
      const sheetAbsen = ss.getSheetByName('DataAbsensi');
      let datesSet = new Set();
      let records = [];
      
      if(sheetAbsen) {
          const map = getHeaderMap_(sheetAbsen);
          const dataAbsen = sheetAbsen.getDataRange().getValues();
          
          let cTA = map['tahun ajaran']; let cKelas = map['kelas'];
          let cTgl = map['tanggal']; let cNis = map['nis']; 
          let cNama = map['nama siswa']; let cStatus = map['status'];
          let cGuru = map['nama guru']; let cMapel = map['mapel'];

          for(let i=1; i<dataAbsen.length; i++) {
              let rowTA = cTA !== undefined ? String(dataAbsen[i][cTA]).trim() : '';
              if(activeTAId && rowTA && rowTA !== activeTAId) continue;

              let k = cKelas !== undefined ? String(dataAbsen[i][cKelas]).trim() : '';
              
              if(allowedKelas && !allowedKelas.has(k)) continue;
              if(kelasFilter && k !== kelasFilter) continue;

              // Ambil Guru & Mapel dari Absensi Data
              let rowGuru = cGuru !== undefined ? String(dataAbsen[i][cGuru]).trim() : '';
              let rowMapel = cMapel !== undefined ? String(dataAbsen[i][cMapel]).trim() : '';

              // Terapkan Filter Tambahan dari Admin
              if (currentUser.role.toLowerCase() !== 'guru' && guruFilter && rowGuru.toLowerCase() !== guruFilter.toLowerCase()) continue;
              if (currentUser.role.toLowerCase() === 'guru' && rowGuru.toLowerCase() !== currentUser.nama.toLowerCase()) continue; 
              
              if (mapelFilter && rowMapel.toLowerCase() !== mapelFilter.toLowerCase()) continue;

              let tglStr = formatTgl_(cTgl !== undefined ? dataAbsen[i][cTgl] : '');
              if(tglStr) datesSet.add(tglStr);

              records.push({
                  nis: cNis !== undefined ? String(dataAbsen[i][cNis]).trim() : '', 
                  nama: cNama !== undefined ? String(dataAbsen[i][cNama]).trim() : '',
                  kelas: k, 
                  tgl: tglStr, 
                  status: cStatus !== undefined ? String(dataAbsen[i][cStatus]).trim() : '',
                  guru: rowGuru,
                  mapel: rowMapel
              });
          }
      }

      let dates = Array.from(datesSet).sort();

      // 2. Bulk Fetch Data Master Siswa
      let masterSiswa = [];
      const sheetSiswa = ss.getSheetByName('DataSiswa') || ss.getSheetByName('Data_Siswa');
      if (sheetSiswa) {
          const dataSiswa = sheetSiswa.getDataRange().getValues();
          let mSiswa = getHeaderMap_(sheetSiswa);
          let nisCol = mSiswa['nis'] !== undefined ? mSiswa['nis'] : (mSiswa['nisn'] !== undefined ? mSiswa['nisn'] : 1);
          let namaCol = mSiswa['nama_siswa'] !== undefined ? mSiswa['nama_siswa'] : (mSiswa['nama lengkap'] !== undefined ? mSiswa['nama lengkap'] : (mSiswa['nama'] !== undefined ? mSiswa['nama'] : 2));
          let kelasCol = mSiswa['kelas'] !== undefined ? mSiswa['kelas'] : 3;

          for(let i=1; i<dataSiswa.length; i++) {
              let k = String(dataSiswa[i][kelasCol]).trim();
              if (allowedKelas && !allowedKelas.has(k)) continue;
              if (kelasFilter && k !== kelasFilter) continue;

              let nis = String(dataSiswa[i][nisCol]).trim();
              let nama = String(dataSiswa[i][namaCol]).trim();
              if(nis && nama) masterSiswa.push({nis: nis, nama: nama, kelas: k});
          }
      }
      
      if (allowedKelas && kelasFilter && !allowedKelas.has(kelasFilter)) masterSiswa = [];

      // 3. Mapping Cepat & Grouping berdasarkan NIS dan MAPEL
      let grouped = {};
      let studentHasRecord = {};
      
      let masterMap = {};
      masterSiswa.forEach(s => masterMap[s.nis] = s);

      records.forEach(r => {
          let mapelKey = r.mapel ? r.mapel.trim() : "-";
          let key = r.nis + "_" + mapelKey;
          studentHasRecord[r.nis] = true;

          if(!grouped[key]) {
              grouped[key] = { 
                  nis: r.nis, 
                  nama: masterMap[r.nis] ? masterMap[r.nis].nama : r.nama, 
                  kelas: masterMap[r.nis] ? masterMap[r.nis].kelas : r.kelas, 
                  mapel: mapelKey,
                  attendance: {}, 
                  guruSet: new Set()
              };
          }
          
          grouped[key].attendance[r.tgl] = r.status;
          if (r.guru) grouped[key].guruSet.add(r.guru);
      });

      // Jika tidak ada filter guru/mapel spesifik, tampilkan juga siswa yang belum ada rekam absensi
      if (!mapelFilter && !guruFilter) {
          masterSiswa.forEach(s => {
              if (!studentHasRecord[s.nis]) {
                  let key = s.nis + "_-";
                  grouped[key] = {
                      nis: s.nis,
                      nama: s.nama,
                      kelas: s.kelas,
                      mapel: '-',
                      attendance: {},
                      guruSet: new Set()
                  };
              }
          });
      }

      let finalData = Object.values(grouped).map(g => {
          g.guru = Array.from(g.guruSet).filter(Boolean).join(', ');
          delete g.guruSet;
          return g;
      }).sort((a, b) => {
          let cmpKls = a.kelas.localeCompare(b.kelas); 
          if (cmpKls !== 0) return cmpKls;
          let cmpNama = a.nama.localeCompare(b.nama); 
          if (cmpNama !== 0) return cmpNama;
          return a.mapel.localeCompare(b.mapel); // Sort by Mapel as final tie-breaker
      });
      
      return { status: 'success', data: { dates: dates, students: finalData } };
  } catch(e) { return { status: 'error', message: 'Gagal memuat absensi: ' + e.message }; }
}

function getSiswaForInputAbsen(kelas, tanggal) {
  try {
       let taAktif = getActiveTAInternal_();
       let masterSiswa = getSiswaByKelas_(kelas);
       
       if(masterSiswa.length === 0) return { status: 'success', data: [] };
       
       const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DataAbsensi');
       let existingData = {};
       
       if(sheet) {
           const data = sheet.getDataRange().getValues();
           let map = getHeaderMap_(sheet);
           let cTgl = map['tanggal']; let cTA = map['tahun ajaran'];
           let cKls = map['kelas']; let cNis = map['nis']; let cStat = map['status'];

           for(let i=1; i<data.length; i++) {
               let rowTglStr = formatTgl_(cTgl !== undefined ? data[i][cTgl] : '');
               let rowTA = cTA !== undefined ? String(data[i][cTA]).trim() : '';
               let rowKelas = cKls !== undefined ? String(data[i][cKls]).trim() : '';
               
               if(rowTA === taAktif && rowTglStr === tanggal && rowKelas === kelas) {
                    let nis = cNis !== undefined ? String(data[i][cNis]).trim() : '';
                    let status = cStat !== undefined ? String(data[i][cStat]).trim() : '';
                    if(nis) existingData[nis] = status;
               }
           }
       }
       let result = masterSiswa.map(s => ({ ...s, status: existingData[s.nis] || 'H' }));
       return { status: 'success', data: result };
  } catch(e) { return { status: 'error', message: e.message }; }
}

function saveAbsensiMassal(payload, currentUser) {
   const lock = LockService.getScriptLock();
   try {
      lock.waitLock(15000); 
      let taAktif = getActiveTAInternal_() || "TA Aktif";
      const requiredHeaders = ['ID', 'NIS', 'Nama Siswa', 'Kelas', 'Tanggal', 'Status', 'Nama Guru', 'Mapel', 'Tahun Ajaran', 'Timestamp', 'ID yang Menginput'];
      const sheet = getSheetWithHeaders_('DataAbsensi', requiredHeaders);
      const map = getHeaderMap_(sheet);
      
      // MENGAMBIL SEMUA DATA DALAM 1 PERINTAH (SANGAT CEPAT)
      let data = sheet.getDataRange().getValues(); 
      const timestamp = Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss');
      
      let idUserInput = "";
      if (currentUser && currentUser.nama) {
          let fetched = getIdsByGuruName_(currentUser.nama);
          if (fetched.idFound) idUserInput = fetched.idFound;
      }
      if (!idUserInput && currentUser && currentUser.id) idUserInput = currentUser.id;
      if (!idUserInput) idUserInput = (currentUser && currentUser.nama) ? "USER-" + currentUser.nama : "SYSTEM-ADMIN";
      
      let saveGuru = payload.guru || ((currentUser && currentUser.nama) ? currentUser.nama : 'Admin');
      let saveMapel = payload.mapel || "";

      // Membuat Mapping Index untuk Pencarian Super Cepat (O(1))
      let existingMap = {}; 
      if (data.length > 1) {
          for(let i=1; i<data.length; i++) {
              let rowTglStr = formatTgl_(map['tanggal'] !== undefined ? data[i][map['tanggal']] : '');
              let rowTA = map['tahun ajaran'] !== undefined ? String(data[i][map['tahun ajaran']]).trim() : '';
              let rowKelas = map['kelas'] !== undefined ? String(data[i][map['kelas']]).trim() : '';
              
              if(rowTA === taAktif && rowTglStr === payload.tanggal && rowKelas === payload.kelas) {
                   let nis = map['nis'] !== undefined ? String(data[i][map['nis']]).trim() : '';
                   if(nis) existingMap[nis] = i; // Simpan indeks array langsung
              }
          }
      }

      let newRows = [];
      let isDataModified = false;

      payload.records.forEach(r => {
          if(existingMap[r.nis] !== undefined) {
              // UBAH DATA DI DALAM MEMORI (ARRAY) BUKAN DI SHEET
              let rowIndex = existingMap[r.nis];
              data[rowIndex][map['status']] = r.status;
              data[rowIndex][map['timestamp']] = timestamp;
              
              if(map['id yang menginput'] !== undefined) data[rowIndex][map['id yang menginput']] = idUserInput;
              if(map['nama guru'] !== undefined) data[rowIndex][map['nama guru']] = saveGuru;
              if(map['mapel'] !== undefined) data[rowIndex][map['mapel']] = saveMapel;
              
              isDataModified = true;
          } else {
              // SIAPKAN ROW BARU DI ARRAY TERPISAH
              let newRow = new Array(data[0].length).fill("");
              // Tambahkan random di ID untuk pastikan uniqueness bila loop berjalan sangat cepat
              newRow[map['id']] = 'ABS-' + new Date().getTime() + '-' + r.nis + '-' + Math.floor(Math.random() * 1000);
              newRow[map['nis']] = r.nis;
              newRow[map['nama siswa']] = r.nama;
              newRow[map['kelas']] = payload.kelas;
              newRow[map['tanggal']] = payload.tanggal;
              newRow[map['status']] = r.status;
              newRow[map['nama guru']] = saveGuru;
              newRow[map['mapel']] = saveMapel;
              newRow[map['tahun ajaran']] = taAktif;
              newRow[map['timestamp']] = timestamp;
              newRow[map['id yang menginput']] = idUserInput;
              newRows.push(newRow);
          }
      });

      // EKSEKUSI PENYIMPANAN DALAM 1 ATAU 2 PERINTAH (BULK WRITE)
      if (isDataModified) {
          // Menulis ulang data lama yang sudah diperbarui dengan 1 perintah
          sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
      }
      
      if(newRows.length > 0) {
          // Menulis semua data baru dengan 1 perintah
          sheet.getRange(data.length + 1, 1, newRows.length, newRows[0].length).setValues(newRows);
      }
      
      SpreadsheetApp.flush(); // Sinkronisasi Real-Time
      return { status: 'success', message: 'Absensi harian berhasil disimpan dengan Cepat & Aman!' };
   } catch(e) { 
      return { status: 'error', message: e.message.includes('timeout') ? 'Sistem sedang memproses antrean. Silakan coba lagi.' : e.message }; 
   } finally { 
      lock.releaseLock(); 
   }
}

function getJamDariJadwal(guru, kelas) {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        let taAktif = getActiveTAInternal_();
        let sheet = ss.getSheetByName('DataJadwal');
        if (!sheet) return "";

        const data = sheet.getDataRange().getValues();
        let map = getHeaderMap_(sheet);
        
        let targetGuru = String(guru).trim().toLowerCase();
        let targetKelas = String(kelas).trim().toLowerCase();

        for (let i = 1; i < data.length; i++) {
            let rowTA = map['tahun ajaran'] !== undefined ? String(data[i][map['tahun ajaran']]).trim() : '';
            if (taAktif && rowTA && rowTA !== taAktif) continue;

            let rGuru = map['nama guru'] !== undefined ? String(data[i][map['nama guru']]).trim().toLowerCase() : '';
            let rKelas = map['kelas'] !== undefined ? String(data[i][map['kelas']]).trim().toLowerCase() : '';

            if (rGuru === targetGuru && rKelas === targetKelas) {
                let jamMulai = map['jam mulai'] !== undefined ? String(data[i][map['jam mulai']]).trim() : '';
                let jamSelesai = map['jam selesai'] !== undefined ? String(data[i][map['jam selesai']]).trim() : '';
                
                if(jamMulai && jamSelesai) {
                    return `${jamMulai}-${jamSelesai}`;
                } else if (jamMulai) {
                    return jamMulai;
                }
            }
        }
        return ""; 
    } catch(e) {
        return "";
    }
}
/**
 * ZettBOT Enhancement - Backend Function untuk Modul Absensi
 * Pastikan untuk menaruh file ini di bagian bawah Code.gs atau di file gs mana saja
 */

function hapusAbsensiKelasCepat(tanggalFilter, kelasFilter) {
    const lock = LockService.getScriptLock();
    try {
        lock.waitLock(15000);
        let taAktif = getActiveTAInternal_(); 
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const sheet = ss.getSheetByName('DataAbsensi');
        
        if (!sheet) return { status: 'error', message: 'Database Absensi tidak ditemukan.' };
        
        const data = sheet.getDataRange().getValues();
        if (data.length <= 1) return { status: 'error', message: 'Data absensi masih kosong.' };
  
        let m = getHeaderMap_(sheet);
        let colTgl = m['tanggal'];
        let colKelas = m['kelas'];
        let colTA = m['tahun ajaran'];
  
        let newData = [data[0]]; // Simpan header
        let deletedCount = 0;
  
        for (let i = 1; i < data.length; i++) {
            let tglRow = formatTgl_(colTgl !== undefined ? data[i][colTgl] : '');
            let kelasRow = colKelas !== undefined ? String(data[i][colKelas]).trim() : '';
            let taRow = colTA !== undefined ? String(data[i][colTA]).trim() : '';
  
            // Jika kondisi cocok, kita SKIP (berarti dihapus dari array memori)
            if ((taAktif === null || taRow === taAktif) && tglRow === tanggalFilter && kelasRow === kelasFilter) {
                deletedCount++;
            } else {
                newData.push(data[i]);
            }
        }
  
        if (deletedCount === 0) {
            return { status: 'error', message: 'Tidak ada data presensi yang ditemukan untuk Kelas dan Tanggal tersebut.' };
        }
  
        // PERBAIKAN BUG TYPO: Menggunakan s (clearContents) -> API Google yang benar
        sheet.clearContents();
        
        // Tulis Ulang array hasil seleksi dalam 1 instruksi tunggal (Super Cepat)
        sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
        SpreadsheetApp.flush();
  
        return { 
            status: 'success', 
            message: `Berhasil menghapus ${deletedCount} riwayat presensi dari database.` 
        };
  
    } catch(e) {
        return { status: 'error', message: e.message };
    } finally {
        lock.releaseLock();
    }
}
  
function updateAbsensiTunggal(nis, tanggal, kelas, newStatus) {
    const lock = LockService.getScriptLock();
    try {
        lock.waitLock(15000);
        let taAktif = getActiveTAInternal_(); 
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DataAbsensi');
        if(!sheet) throw new Error('Sheet DataAbsensi tidak ditemukan.');
        
        const data = sheet.getDataRange().getValues();
        let m = getHeaderMap_(sheet);
        
        for(let i=1; i<data.length; i++) {
            let rowTA = String(data[i][m['tahun ajaran']] || '').trim();
            let rowTgl = formatTgl_(data[i][m['tanggal']]);
            let rowNis = String(data[i][m['nis']] || '').trim();
            let rowKelas = String(data[i][m['kelas']] || '').trim();
            
            if((taAktif === null || rowTA === taAktif) && rowTgl === tanggal && rowNis === nis && rowKelas === kelas) {
                sheet.getRange(i+1, m['status'] + 1).setValue(newStatus);
                sheet.getRange(i+1, m['timestamp'] + 1).setValue(Utilities.formatDate(new Date(), 'Asia/Jakarta', 'dd/MM/yyyy HH:mm:ss'));
                SpreadsheetApp.flush();
                return {status: 'success', message: 'Status Diupdate'};
            }
        }
        return {status: 'error', message: 'Data presensi tersebut tidak ditemukan di sistem.'};
    } catch(e) {
        return {status: 'error', message: e.message};
    } finally { lock.releaseLock(); }
}
  
function deleteAbsensiTunggal(nis, tanggal, kelas) {
    const lock = LockService.getScriptLock();
    try {
        lock.waitLock(15000);
        let taAktif = getActiveTAInternal_(); 
        const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('DataAbsensi');
        if(!sheet) throw new Error('Sheet DataAbsensi tidak ditemukan.');
        
        const data = sheet.getDataRange().getValues();
        let m = getHeaderMap_(sheet);
        
        // Loop mundur (dari bawah ke atas) aman untuk fungsi Hapus Baris Manual (deleteRow)
        for(let i = data.length - 1; i >= 1; i--) {
            let rowTA = String(data[i][m['tahun ajaran']] || '').trim();
            let rowTgl = formatTgl_(data[i][m['tanggal']]);
            let rowNis = String(data[i][m['nis']] || '').trim();
            let rowKelas = String(data[i][m['kelas']] || '').trim();
            
            if((taAktif === null || rowTA === taAktif) && rowTgl === tanggal && rowNis === nis && rowKelas === kelas) {
                sheet.deleteRow(i+1);
                SpreadsheetApp.flush();
                return {status: 'success', message: 'Data spesifik berhasil dihapus'};
            }
        }
        return {status: 'error', message: 'Data presensi tersebut tidak ditemukan di sistem.'};
    } catch(e) {
        return {status: 'error', message: e.message};
    } finally { lock.releaseLock(); }
}
