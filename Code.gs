// ==================== KONFIGURASI ====================
const SPREADSHEET_ID = 'YOUR_SPREADSHEET_ID_HERE'; // GANTI
const sheetPegawai = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('PEGAWAI');
const sheetMobil = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('MOBIL');
const sheetPeminjaman = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName('PEMINJAMAN');

// ==================== ENTRY POINT ====================
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('SiMoDi - Sistem Mobil Dinas')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ==================== SESSION ====================
function getSession(token) {
  if (!token) return null;
  try {
    const cache = CacheService.getScriptCache();
    const data = cache.get(token);
    if (data) return JSON.parse(data);
  } catch(e) {}
  return null;
}

// ==================== LOGIN ====================
function login(nip, password) {
  try {
    const data = sheetPegawai.getDataRange().getValues();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const nipSheet = String(row[1]).trim();
      const passSheet = String(row[2]).trim();
      
      if (nipSheet === nip.trim() && passSheet === password.trim()) {
        const userData = {
          id: Number(row[0]),  // PASTIKAN ID NUMERIC
          nip: nipSheet,
          nama: String(row[3]).trim(),
          unit: String(row[4]).trim(),
          seksi: String(row[5]).trim(),
          role: String(row[6]).trim().toLowerCase()  // LOWERCASE
        };
        
        const token = Utilities.getUuid();
        CacheService.getScriptCache().put(token, JSON.stringify(userData), 21600);
        
        return { 
          success: true, 
          token: token, 
          role: userData.role,
          nama: userData.nama,
          unit: userData.unit
        };
      }
    }
    
    return { success: false, message: "NIP atau Password salah" };
  } catch(e) {
    return { success: false, message: "Error: " + e.toString() };
  }
}

function logout(token) {
  if (token) {
    try { CacheService.getScriptCache().remove(token); } catch(e) {}
  }
  return { success: true };
}

// ==================== API MOBIL ====================
function getMobilTersedia() {
  try {
    const data = sheetMobil.getDataRange().getValues();
    const mobil = [];
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const status = String(row[3]).trim().toLowerCase();
      
      if (status === 'tersedia') {
        mobil.push({
          id: Number(row[0]),
          nama: String(row[1]).trim(),
          plat: String(row[2]).trim()
        });
      }
    }
    
    return mobil;
  } catch(e) {
    return [];
  }
}

// ==================== API PEMINJAMAN PEGAWAI ====================
function ajukanPeminjaman(token, idMobil, tglPinjam, jamPinjam, tglKembali, jamKembali, jenisTujuan, detailTujuan) {
  const user = getSession(token);
  if (!user) return { success: false, message: "Session tidak valid." };
  
  try {
    // Validasi bentrok
    const data = sheetPeminjaman.getDataRange().getValues();
    const pinjamBaru = new Date(`${tglPinjam}T${jamPinjam}`);
    const kembaliBaru = new Date(`${tglKembali}T${jamKembali}`);
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const status = String(row[9]).trim();
      if (Number(row[2]) === Number(idMobil) && (status === 'disetujui' || status === 'menunggu')) {
        const pinjamExist = new Date(`${row[3]}T${row[4]}`);
        const kembaliExist = new Date(`${row[5]}T${row[6]}`);
        if (pinjamBaru < kembaliExist && kembaliBaru > pinjamExist) {
          return { success: false, message: "Mobil sudah dipesan." };
        }
      }
    }
    
    // Insert data
    const idBaru = sheetPeminjaman.getLastRow();
    sheetPeminjaman.appendRow([
      idBaru,
      user.id,
      Number(idMobil),
      tglPinjam,
      jamPinjam,
      tglKembali,
      jamKembali,
      jenisTujuan,
      detailTujuan,
      'menunggu',
      new Date().toLocaleString('id-ID')
    ]);
    
    return { success: true, message: "Pengajuan berhasil." };
  } catch(e) {
    return { success: false, message: "Error: " + e.toString() };
  }
}

function getRiwayatPeminjaman(token) {
  const user = getSession(token);
  if (!user) {
    Logger.log('getRiwayatPeminjaman: User tidak ditemukan');
    return [];
  }
  
  Logger.log('getRiwayatPeminjaman: User ID = ' + user.id + ', Nama = ' + user.nama);
  
  try {
    const dataP = sheetPeminjaman.getDataRange().getValues();
    const dataM = sheetMobil.getDataRange().getValues();
    const riwayat = [];
    
    Logger.log('Total baris peminjaman: ' + dataP.length);
    
    for (let i = 1; i < dataP.length; i++) {
      const row = dataP[i];
      const idPegawaiDiSheet = Number(row[1]);
      
      Logger.log('Baris ' + i + ': id_pegawai=' + idPegawaiDiSheet + ' user.id=' + user.id + ' MATCH=' + (idPegawaiDiSheet === user.id));
      
      if (idPegawaiDiSheet === user.id) {
        const mobil = dataM.find(m => Number(m[0]) === Number(row[2]));
        
        riwayat.push({
          id: row[0],
          namaMobil: mobil ? String(mobil[1]).trim() : '?',
          platMobil: mobil ? String(mobil[2]).trim() : '?',
          tglPinjam: row[3],
          jamPinjam: row[4],
          tglKembali: row[5],
          jamKembali: row[6],
          jenisTujuan: row[7],
          detailTujuan: row[8],
          status: String(row[9]).trim()
        });
      }
    }
    
    Logger.log('Riwayat ditemukan: ' + riwayat.length);
    return riwayat;
  } catch(e) {
    Logger.log('ERROR: ' + e.toString());
    return [];
  }
}

// ==================== API ADMIN ====================
function getPengajuanMenunggu(token) {
  const user = getSession(token);
  Logger.log('getPengajuanMenunggu: User = ' + JSON.stringify(user));
  
  if (!user || user.role !== 'admin') {
    Logger.log('Akses ditolak: role=' + (user ? user.role : 'null'));
    return [];
  }
  
  try {
    const dataP = sheetPeminjaman.getDataRange().getValues();
    const dataPg = sheetPegawai.getDataRange().getValues();
    const dataM = sheetMobil.getDataRange().getValues();
    const pengajuan = [];
    
    Logger.log('Total baris peminjaman: ' + dataP.length);
    
    for (let i = 1; i < dataP.length; i++) {
      const row = dataP[i];
      const status = String(row[9]).trim();
      
      Logger.log('Baris ' + i + ': status=' + status);
      
      if (status === 'menunggu') {
        const peg = dataPg.find(r => Number(r[0]) === Number(row[1]));
        const mob = dataM.find(r => Number(r[0]) === Number(row[2]));
        
        pengajuan.push({
          id: row[0],
          idPegawai: row[1],
          namaPegawai: peg ? String(peg[3]).trim() : '?',
          unitPegawai: peg ? String(peg[4]).trim() : '?',
          seksiPegawai: peg ? String(peg[5]).trim() : '?',
          namaMobil: mob ? String(mob[1]).trim() : '?',
          platMobil: mob ? String(mob[2]).trim() : '?',
          tglPinjam: row[3],
          jamPinjam: row[4],
          tglKembali: row[5],
          jamKembali: row[6],
          jenisTujuan: row[7],
          detailTujuan: row[8]
        });
      }
    }
    
    Logger.log('Pengajuan menunggu: ' + pengajuan.length);
    return pengajuan;
  } catch(e) {
    Logger.log('ERROR: ' + e.toString());
    return [];
  }
}

function setujuiPeminjaman(token, idPeminjaman) {
  const user = getSession(token);
  if (!user || user.role !== 'admin') return { success: false, message: "Akses ditolak" };
  
  try {
    const data = sheetPeminjaman.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (Number(data[i][0]) === Number(idPeminjaman)) {
        sheetPeminjaman.getRange(i + 1, 10).setValue('disetujui');
        
        const idMobil = Number(data[i][2]);
        const dataMobil = sheetMobil.getDataRange().getValues();
        for (let j = 1; j < dataMobil.length; j++) {
          if (Number(dataMobil[j][0]) === idMobil) {
            sheetMobil.getRange(j + 1, 4).setValue('dipakai');
            break;
          }
        }
        return { success: true, message: "Disetujui." };
      }
    }
    return { success: false, message: "Data tidak ditemukan." };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function tolakPeminjaman(token, idPeminjaman, alasan) {
  const user = getSession(token);
  if (!user || user.role !== 'admin') return { success: false, message: "Akses ditolak" };
  
  try {
    const data = sheetPeminjaman.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (Number(data[i][0]) === Number(idPeminjaman)) {
        sheetPeminjaman.getRange(i + 1, 10).setValue('ditolak');
        return { success: true, message: "Ditolak." };
      }
    }
    return { success: false, message: "Data tidak ditemukan." };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function konfirmasiSelesai(token, idPeminjaman) {
  const user = getSession(token);
  if (!user || user.role !== 'admin') return { success: false, message: "Akses ditolak" };
  
  try {
    const data = sheetPeminjaman.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      if (Number(data[i][0]) === Number(idPeminjaman)) {
        sheetPeminjaman.getRange(i + 1, 10).setValue('selesai');
        
        const idMobil = Number(data[i][2]);
        const dataMobil = sheetMobil.getDataRange().getValues();
        for (let j = 1; j < dataMobil.length; j++) {
          if (Number(dataMobil[j][0]) === idMobil) {
            sheetMobil.getRange(j + 1, 4).setValue('tersedia');
            break;
          }
        }
        return { success: true, message: "Selesai." };
      }
    }
    return { success: false, message: "Data tidak ditemukan." };
  } catch(e) {
    return { success: false, message: e.toString() };
  }
}

function getRingkasanMobil(token) {
  const user = getSession(token);
  if (!user || user.role !== 'admin') return [];
  
  try {
    const data = sheetMobil.getDataRange().getValues();
    const ringkasan = [];
    for (let i = 1; i < data.length; i++) {
      ringkasan.push({
        id: Number(data[i][0]),
        nama: String(data[i][1]).trim(),
        plat: String(data[i][2]).trim(),
        status: String(data[i][3]).trim().toLowerCase()
      });
    }
    return ringkasan;
  } catch(e) {
    return [];
  }
}

function getAllPeminjaman(token) {
  const user = getSession(token);
  if (!user || user.role !== 'admin') return [];
  
  try {
    const dataP = sheetPeminjaman.getDataRange().getValues();
    const dataPg = sheetPegawai.getDataRange().getValues();
    const dataM = sheetMobil.getDataRange().getValues();
    const semua = [];
    
    for (let i = dataP.length - 1; i >= 1; i--) {
      const peg = dataPg.find(r => Number(r[0]) === Number(dataP[i][1]));
      const mob = dataM.find(r => Number(r[0]) === Number(dataP[i][2]));
      semua.push({
        id: dataP[i][0],
        namaPegawai: peg ? String(peg[3]).trim() : '?',
        unitPegawai: peg ? String(peg[4]).trim() : '?',
        namaMobil: mob ? String(mob[1]).trim() : '?',
        platMobil: mob ? String(mob[2]).trim() : '?',
        tglPinjam: dataP[i][3],
        jamPinjam: dataP[i][4],
        tglKembali: dataP[i][5],
        jamKembali: dataP[i][6],
        detailTujuan: dataP[i][8],
        status: String(dataP[i][9]).trim()
      });
      if (semua.length >= 50) break;
    }
    return semua;
  } catch(e) {
    return [];
  }
}
