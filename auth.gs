/**
 * Auth.gs - Modul Autentikasi Pengguna
 * Diperbarui oleh ZettBOT: Dynamic Header Detection (Anti-Error)
 */

function authenticateUser(username, password) {
  try {
    const data = getSheetData_('Pengguna'); 
    
    if (data.length <= 1) {
      return { status: 'error', message: 'Database pengguna kosong atau belum diatur.' };
    }

    // 1. Membaca Header (Baris pertama) untuk mendeteksi posisi kolom otomatis
    const headers = data[0].map(h => h.toString().trim().toLowerCase());
    const users = data.slice(1);
    
    // 2. Pencarian Index Kolom Dinamis (Mencegah error jika kolom bergeser)
    const userIdx = headers.indexOf('username');
    const passIdx = headers.indexOf('password');
    const roleIdx = headers.indexOf('role');
    const statusIdx = headers.indexOf('status');
    const namaIdx = headers.indexOf('nama pengguna');

    // Fallback index (jika nama header diubah manual oleh user di Spreadsheet)
    const uIdx = userIdx > -1 ? userIdx : 2;
    const pIdx = passIdx > -1 ? passIdx : 3;
    const rIdx = roleIdx > -1 ? roleIdx : 4;
    const sIdx = statusIdx > -1 ? statusIdx : 5;
    const nIdx = namaIdx > -1 ? namaIdx : 1;

    // 3. Pencarian data pengguna
    const matchedUser = users.find(row => row[uIdx] === username && row[pIdx] === password);

    // Jika tidak ditemukan kecocokan username dan password
    if (!matchedUser) {
      return { status: 'error', message: 'Username atau Password salah!' };
    }

    // 4. Validasi Status Pengguna
    const statusPengguna = matchedUser[sIdx];
    if (statusPengguna && statusPengguna.toString().toLowerCase() !== 'aktif') {
      return { status: 'error', message: 'Akun Anda berstatus Non-Aktif. Silakan hubungi Admin.' };
    }

    // 5. Penentuan Role (Dinamis membaca dari kolom Role, atau tebak dari username jika kolom tidak ada)
    let userRole = 'admin'; // Default
    if (roleIdx > -1 && matchedUser[rIdx]) {
      userRole = matchedUser[rIdx].toString().toLowerCase();
    } else {
      const uname = matchedUser[uIdx].toString().toLowerCase();
      if (uname.includes('guru')) userRole = 'guru';
      else if (uname.includes('siswa')) userRole = 'siswa';
    }

    // Kembalikan response sukses ke Frontend
    return {
      status: 'success',
      data: { 
        username: matchedUser[uIdx], 
        nama: matchedUser[nIdx],
        role: userRole 
      }
    };
    
  } catch (error) {
    Logger.log("Error di authenticateUser: " + error.message);
    return { status: 'error', message: "Terjadi kesalahan sistem: " + error.message };
  }
}
