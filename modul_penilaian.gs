// =======================================================================
// BACKEND OPTIMIZED: MODUL PENILAIAN (BOBOT, PRAKTEK, & STATS DETAIL)
// Architected for High Performance & Low API Calls
// =======================================================================

/**
 * HELPER: Universal Sheet Getter & Creator
 * Mengurangi redundansi kode pengecekan dan pembuatan sheet
 */
function _getDbSheet(sheetName, headers) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#e0e7ff");
  } else if (sheetName === 'Nilai_Praktek') {
    // Auto-Upgrade untuk kolom TA & Semester khusus Nilai_Praktek
    let lastCol = sheet.getLastColumn();
    if (lastCol < 11) {
      let currentHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
      if (!currentHeaders.includes("TA")) {
         sheet.getRange(1, 10).setValue("TA").setFontWeight("bold").setBackground("#e0e7ff");
         sheet.getRange(1, 11).setValue("Semester").setFontWeight("bold").setBackground("#e0e7ff");
      }
    }
  }
  return sheet;
}

function getMapingGuruData() {
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        let sheet = ss.getSheetByName('Maping_Guru');
        if(!sheet) return { status: 'success', data: { guruList: [], maping: {}, mapingMapel: {} } };
        
        let maping = {};
        let mapingMapel = {};
        let guruSet = new Set();
        const data = sheet.getDataRange().getValues();
        
        for (let i = 1; i < data.length; i++) {
            let guru = data[i][4];     
            let mapelRaw = data[i][5]; 
            let kelasRaw = data[i][6]; 

            if (guru) {
                guruSet.add(guru);
                if (kelasRaw) {
                    kelasRaw.toString().split(',').forEach(k => {
                        let kStr = k.trim();
                        if (!maping[guru]) maping[guru] = new Set();
                        maping[guru].add(kStr);
                    });
                }
                if (mapelRaw) {
                    mapelRaw.toString().split(',').forEach(m => {
                        let mStr = m.trim();
                        if (!mapingMapel[guru]) mapingMapel[guru] = new Set();
                        mapingMapel[guru].add(mStr);
                    });
                }
            }
        }
        
        // Transform Set to Array
        for (let g in maping) maping[g] = Array.from(maping[g]);
        for (let g in mapingMapel) mapingMapel[g] = Array.from(mapingMapel[g]);
        
        return { status: 'success', data: { guruList: Array.from(guruSet), maping: maping, mapingMapel: mapingMapel } };
    } catch (e) {
        return { status: 'error', message: e.message };
    }
}

function getBobotPenilaian(role, nama) {
    try {
        let sheet = _getDbSheet('Bobot_Penilaian', ["ID", "Nama Materi", "Mapel", "Kategori", "Bobot", "TA", "Semester", "Nama Guru"]);
        let data = sheet.getDataRange().getValues();
        let results = [];
        let isGuru = role && role.toLowerCase() === 'guru';
        
        for(let i = 1; i < data.length; i++) {
            let r = data[i];
            if(!r[0]) continue;
            if (isGuru && r[7] !== nama) continue;
            
            results.push({
                'ID': r[0], 'Nama Materi': r[1], 'Mapel': r[2], 'Kategori': r[3],
                'Bobot': r[4], 'TA': r[5], 'Semester': r[6], 'Guru': r[7]
            });
        }
        return { status: 'success', data: results };
    } catch(e) {
        return { status: 'error', message: e.message };
    }
}

function simpanBobotPenilaian(payload) {
    try {
        let sheet = _getDbSheet('Bobot_Penilaian', ["ID", "Nama Materi", "Mapel", "Kategori", "Bobot", "TA", "Semester", "Nama Guru"]);
        
        // OPTIMASI: Batch update in memory, hindari N+1 query setValue()
        if(payload.id) {
            let data = sheet.getDataRange().getValues();
            let isUpdated = false;
            for(let i = 1; i < data.length; i++) {
                if(data[i][0] === payload.id) {
                    data[i] = [payload.id, payload.namaMateri, payload.mapel, payload.kategori, payload.bobot, payload.ta, payload.semester, payload.guru];
                    isUpdated = true;
                    break;
                }
            }
            if(isUpdated) {
                sheet.getRange(1, 1, data.length, data[0].length).setValues(data);
                return {status: 'success'};
            }
        }
        
        let newId = 'BBT-' + new Date().getTime();
        sheet.appendRow([newId, payload.namaMateri, payload.mapel, payload.kategori, payload.bobot, payload.ta, payload.semester, payload.guru]);
        return { status: 'success' };
    } catch(e) { 
        return { status: 'error', message: e.message }; 
    }
}

function hapusBobotPenilaian(id) {
    try {
        let sheet = _getDbSheet('Bobot_Penilaian', ["ID", "Nama Materi", "Mapel", "Kategori", "Bobot", "TA", "Semester", "Nama Guru"]);
        let data = sheet.getDataRange().getValues();
        for(let i = 1; i < data.length; i++) {
            if(data[i][0] === id) {
                sheet.deleteRow(i+1);
                return {status: 'success'}; // Flush dihapus karena memperlambat, UI sudah menghandle optimistically
            }
        }
        throw new Error("ID Bobot tidak ditemukan.");
    } catch(e) { 
        return { status: 'error', message: e.message }; 
    }
}

function getSiswaByKelasMaping(role, nama) {
     try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        let sheetSiswa = ss.getSheetByName('Data_Siswa') || ss.getSheetByName('Siswa');
        if(!sheetSiswa) throw new Error("Sheet Data_Siswa tidak ditemukan.");
        
        let maping = getMapingGuruData().data.maping;
        let allowedKelas = maping[nama] || [];
        let isGuru = role.toLowerCase() === 'guru';
        
        let data = sheetSiswa.getDataRange().getValues();
        let headers = data[0].map(h => h.toString().toLowerCase().trim());
        
        let idxNis = headers.findIndex(h => h.includes('nis'));
        let idxNama = headers.findIndex(h => h.includes('nama'));
        let idxKelas = headers.findIndex(h => h.includes('kelas') || h.includes('rombel'));
        
        if(idxNis === -1) idxNis = 0;
        if(idxNama === -1) idxNama = 1;
        if(idxKelas === -1) idxKelas = 2;
        
        let siswa = [];
        let kelasSet = new Set();
        
        for(let i = 1; i < data.length; i++) {
            let kls = data[i][idxKelas] ? data[i][idxKelas].toString().trim() : "";
            if(isGuru && !allowedKelas.includes(kls)) continue;
            
            if(kls) {
                kelasSet.add(kls);
                siswa.push({
                    nis: data[i][idxNis] ? data[i][idxNis].toString() : "", 
                    nama: data[i][idxNama],
                    kelas: kls
                });
            }
        }
        
        // Sorting di Frontend lebih disarankan, namun jika harus di backend, dioptimasi:
        siswa.sort((a, b) => a.kelas.localeCompare(b.kelas) || a.nama.localeCompare(b.nama));
        return { status: 'success', data: { siswa: siswa, kelas: Array.from(kelasSet) } };
     } catch(e) { 
        return { status: 'error', message: e.message }; 
     }
}

function getMateriDetailStats(idMateri, role, nama) {
  try {
    const listSiswa = getSiswaByKelasMaping(role, nama).data.siswa || []; 
    const sheetNilai = _getDbSheet('Nilai_Praktek', ["ID_Nilai", "ID_Materi", "NIS", "Nama Siswa", "Kelas", "Mapel", "Nilai", "Tanggal", "Nama Guru", "TA", "Semester"]);
    const dNilai = sheetNilai.getDataRange().getValues();
    const setSudahMateri = new Set();
    
    for (let i = 1; i < dNilai.length; i++) {
      if (dNilai[i][1] === idMateri && dNilai[i][2]) {
        setSudahMateri.add(dNilai[i][2].toString());
      }
    }
    
    const classMap = {};
    listSiswa.forEach(s => {
      const kls = s.kelas;
      if (!classMap[kls]) classMap[kls] = { kelas: kls, sudah: 0, belum: 0 };
      setSudahMateri.has(s.nis.toString()) ? classMap[kls].sudah++ : classMap[kls].belum++;
    });
    
    const classStats = Object.values(classMap).sort((a, b) => a.kelas.toString().localeCompare(b.kelas.toString()));
    return { status: 'success', data: { classStats: classStats } };
  } catch (e) {
    return { status: 'error', message: e.message };
  }
}

function getNilaiPraktek(idMateri, kelas, tanggal, role, guruPenilai) {
     try {
        let sheet = _getDbSheet('Nilai_Praktek', ["ID_Nilai", "ID_Materi", "NIS", "Nama Siswa", "Kelas", "Mapel", "Nilai", "Tanggal", "Nama Guru", "TA", "Semester"]);
        let data = sheet.getDataRange().getValues();
        let map = {};
        
        for(let i = 1; i < data.length; i++) {
            if(data[i][1] === idMateri && data[i][2]) {
                map[data[i][2].toString()] = { nilai: data[i][6], status: 'Sudah Dinilai' }; 
            }
        }
        return { status: 'success', data: map };
     } catch(e) { 
        return { status: 'error', message: e.message }; 
     }
}

function simpanNilaiPraktekMasal(payloads) {
    const lock = LockService.getScriptLock();
    try {
        lock.waitLock(30000);
    } catch(e) {
        return { status: 'error', message: 'Koneksi sibuk. Mohon coba sesaat lagi.' };
    }

    try {
        const sheet = _getDbSheet('Nilai_Praktek', ["ID_Nilai", "ID_Materi", "NIS", "Nama Siswa", "Kelas", "Mapel", "Nilai", "Tanggal", "Nama Guru", "TA", "Semester"]);
        let lastRow = sheet.getLastRow();
        let lastCol = Math.max(sheet.getLastColumn(), 11);

        let rawData = sheet.getRange(1, 1, lastRow > 0 ? lastRow : 1, lastCol).getValues();
        let headers = rawData[0];
        let rows = rawData.slice(1); 

        let existingMap = {};
        for (let i = 0; i < rows.length; i++) {
            if (rows[i][1] && rows[i][2]) existingMap[`${rows[i][1]}_${rows[i][2]}`] = i;
        }

        payloads.forEach(p => {
            let key = `${p.idMateri}_${p.nis}`;
            let rowData = [
                '', p.idMateri, p.nis.toString(), p.namaSiswa, p.kelas, p.mapel, p.nilai, `'${p.tanggal}`, p.guru, p.ta, p.semester
            ];

            if (key in existingMap) {
                let idx = existingMap[key];
                rowData[0] = rows[idx][0]; 
                rows[idx] = rowData;
            } else {
                rowData[0] = 'NP-' + new Date().getTime() + Math.floor(Math.random() * 1000);
                rows.push(rowData);
                existingMap[key] = rows.length - 1; 
            }
        });

        let finalDataArray = [headers, ...rows];
        sheet.getRange(1, 1, finalDataArray.length, 11).setValues(finalDataArray);
        SpreadsheetApp.flush(); // Tetap dipertahankan untuk keamanan data batch skala besar

        return { status: 'success' };
    } catch(e) { 
        return { status: 'error', message: e.message }; 
    } finally {
        lock.releaseLock();
    }
}

function simpanSatuNilaiPraktek(payload) {
     return simpanNilaiPraktekMasal([payload]); 
}

function hapusSatuNilaiPraktek(payload) {
    try {
        let sheet = _getDbSheet('Nilai_Praktek', ["ID_Nilai", "ID_Materi", "NIS", "Nama Siswa", "Kelas", "Mapel", "Nilai", "Tanggal", "Nama Guru", "TA", "Semester"]);
        let data = sheet.getDataRange().getValues();
        
        for(let i = data.length - 1; i >= 1; i--) {
            if(data[i][1] === payload.idMateri && data[i][2].toString() === payload.nis.toString()) { 
                sheet.deleteRow(i+1);
                break; // Stop jika sudah ketemu, optimasi
            }
        }
        return { status: 'success' };
    } catch(e) { 
        return { status: 'error', message: e.message }; 
    }
}

// Fungsi Baru: Hapus secara masal langsung dari backend untuk optimasi performa
function hapusNilaiPraktekKelas(payload) {
    try {
        let sheet = _getDbSheet('Nilai_Praktek', ["ID_Nilai", "ID_Materi", "NIS", "Nama Siswa", "Kelas", "Mapel", "Nilai", "Tanggal", "Nama Guru", "TA", "Semester"]);
        let data = sheet.getDataRange().getValues();
        
        // Loop dari bawah agar index tidak bergeser saat baris dihapus
        for(let i = data.length - 1; i >= 1; i--) {
            if(data[i][1] === payload.idMateri && data[i][4].toString().toLowerCase() === payload.kelas.toString().toLowerCase()) { 
                sheet.deleteRow(i+1);
            }
        }
        return { status: 'success' };
    } catch(e) { 
        return { status: 'error', message: e.message }; 
    }
}

function getLaporanPenilaianRelasi(role, nama) {
    try {
        // Semua fungsi di bawah sudah dioptimasi, kita memanggilnya secara terpusat
        let bobotRes = getBobotPenilaian(role, nama);
        let siswaRes = getSiswaByKelasMaping(role, nama);
        let mapingRes = getMapingGuruData();
        
        let sheetNilai = _getDbSheet('Nilai_Praktek', ["ID_Nilai", "ID_Materi", "NIS", "Nama Siswa", "Kelas", "Mapel", "Nilai", "Tanggal", "Nama Guru", "TA", "Semester"]);
        let dNilai = sheetNilai.getDataRange().getValues();
        
        // Optimasi Array Mapping
        let arrNilai = dNilai.slice(1).map(r => ({
            idMateri: r[1],
            nis: r[2] ? r[2].toString() : '', 
            nilai: r[6],   
            tanggal: r[7] ? r[7].toString().replace(/'/g, '') : '', 
            ta: r[9] ? r[9].toString() : '',        
            semester: r[10] ? r[10].toString() : '' 
        }));
        
        return { status: 'success', data: {
            bobot: bobotRes.data || [],
            siswa: siswaRes.data ? siswaRes.data.siswa : [],
            nilai: arrNilai,
            mapingGuru: mapingRes.data
        }};
    } catch(e) { 
        return { status: 'error', message: e.message }; 
    }
}
