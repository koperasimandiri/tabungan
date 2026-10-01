/**
 * Backend E-Tabungan Pelajar untuk Google Apps Script.
 *
 * Setup:
 * 1. Buat project Apps Script dan tempel file ini.
 * 2. Jalankan setupSheet() sekali dan izinkan akses Google Sheets.
 * 3. Jalankan createUser('ADM001', 'ADMIN-NIS', 'Administrator', 'admin', '-').
 * 4. Deploy sebagai Web app: execute as pemilik, access sesuai kebutuhan aplikasi.
 */

const CONFIG = {
  spreadsheetId: '1iD7EYFYQF7rklgyaG0150LL6VPObZqLRZOFPmzjtKoA',
  sheets: {
    users: 'Users',
    transactions: 'Transactions',
    dropdowns: 'Dropdowns'
  },
  tokenTtlSeconds: 8 * 60 * 60,
  maxTransactionAmount: 100000000
};

function setupSheet() {
  const properties = PropertiesService.getScriptProperties();
  const spreadsheet = SpreadsheetApp.openById(CONFIG.spreadsheetId);
  properties.setProperty('SPREADSHEET_ID', spreadsheet.getId());
  if (!properties.getProperty('TOKEN_SECRET')) {
    properties.setProperty('TOKEN_SECRET', Utilities.getUuid() + Utilities.getUuid());
  }
  const users = getOrCreateSheet_(spreadsheet, CONFIG.sheets.users,
    ['norek', 'nis', 'nama', 'kelas', 'passwordHash', 'salt', 'role', 'active', 'createdAt']);
  const transactions = getOrCreateSheet_(spreadsheet, CONFIG.sheets.transactions,
    ['id', 'createdAt', 'tanggal', 'nis', 'jenis', 'jumlah', 'keterangan', 'createdBy']);
  getOrCreateSheet_(spreadsheet, CONFIG.sheets.dropdowns, ['type', 'value']);
  migrateUsersSheet_(users);

  [users, transactions].forEach(sheet => sheet.setFrozenRows(1));
  const result = { success: true, spreadsheetId: spreadsheet.getId(), url: spreadsheet.getUrl() };
  Logger.log(JSON.stringify(result));
  return result;
}

function createUser(norek, nis, nama, role, kelas) {
  if (!norek || !nis || !nama) throw new Error('norek, nis, dan nama wajib diisi.');
  role = role || 'siswa';
  kelas = kelas || '-';
  if (!['admin', 'siswa'].includes(role)) throw new Error('Role tidak valid.');

  const sheet = getSheet_();
  const values = sheet.getDataRange().getValues();
  const exists = values.slice(1).some(row =>
    String(row[0]).trim() === String(norek).trim() || String(row[1]).trim() === String(nis).trim()
  );
  if (exists) throw new Error('Nomor rekening atau NIS sudah terdaftar.');

  const salt = Utilities.getUuid();
  sheet.appendRow([
    String(norek).trim(), String(nis).trim(), nama, kelas,
    hashPassword_(nis, salt), salt, role, true, new Date()
  ]);
  clearUsersCache_();
  return { success: true, message: 'User berhasil dibuat.' };
}

// Jalankan fungsi ini dari tombol Run untuk membuat admin awal.
function createInitialAdmin() {
  const norek = 'ADM001';
  const nis = 'ADMIN-NIS';
  const nama = 'Administrator';
  const result = createUser(norek, nis, nama, 'admin', '-');
  Logger.log(result.message + ' Login: ' + norek + ' / ' + nis);
  return result;
}

function resetUserPassword(norek, nis) {
  const sheet = getSheet_();
  const rows = sheet.getDataRange().getValues();
  const rowIndex = rows.findIndex((row, index) => index > 0 && String(row[0]).trim() === String(norek).trim());
  if (rowIndex < 1) throw new Error('Nomor rekening tidak ditemukan.');
  const salt = Utilities.getUuid();
  sheet.getRange(rowIndex + 1, 2).setValue(String(nis).trim());
  sheet.getRange(rowIndex + 1, 5).setValue(hashPassword_(nis, salt));
  sheet.getRange(rowIndex + 1, 6).setValue(salt);
  clearUsersCache_();
  return { success: true, message: 'Password berhasil diubah menjadi NIS.' };
}

function updateAdminLogin(oldNorek, newNorek, newNis) {
  if (!oldNorek || !newNorek || !newNis) throw new Error('Semua data login wajib diisi.');
  const sheet = getSheet_();
  const rows = sheet.getDataRange().getValues();
  const rowIndex = rows.findIndex((row, index) => index > 0 && String(row[0]).trim() === String(oldNorek).trim());
  if (rowIndex < 1) throw new Error('Admin lama tidak ditemukan.');
  if (String(rows[rowIndex][6]).toLowerCase() !== 'admin' && rows[rowIndex][6] !== true) {
    throw new Error('User tersebut bukan admin.');
  }

  const duplicate = rows.some((row, index) => index > 0 && index !== rowIndex &&
    (String(row[0]).trim() === String(newNorek).trim() || String(row[1]).trim() === String(newNis).trim()));
  if (duplicate) throw new Error('Nomor rekening atau NIS baru sudah digunakan.');

  const oldNis = String(rows[rowIndex][1]).trim();
  const salt = Utilities.getUuid();
  sheet.getRange(rowIndex + 1, 1, 1, 6).setValues([[
    String(newNorek).trim(), String(newNis).trim(), rows[rowIndex][2], rows[rowIndex][3],
    hashPassword_(newNis, salt), salt
  ]]);

  const transactionSheet = getSheet_(CONFIG.sheets.transactions);
  const transactionRows = transactionSheet.getDataRange().getValues();
  transactionRows.forEach((row, index) => {
    if (index > 0 && String(row[7]).trim() === oldNis) transactionSheet.getRange(index + 1, 8).setValue(String(newNis).trim());
  });
  clearUsersCache_();
  return { success: true, message: 'Login admin berhasil diperbarui.' };
}

function bulkCreateUsers_(actor, users) {
  if (!Array.isArray(users) || !users.length) throw new Error('Data anggota kosong.');
  if (users.length > 500) throw new Error('Maksimal 500 anggota per impor.');

  const sheet = getSheet_();
  const existing = getUsers_();
  const noreks = new Set(existing.map(user => user.norek));
  const nises = new Set(existing.map(user => user.nis));
  const batchNoreks = new Set();
  const batchNises = new Set();
  const rows = [];
  const errors = [];

  users.forEach((item, index) => {
    const norek = String(item.norek || '').trim();
    const nis = String(item.nis || '').trim();
    const nama = String(item.nama || '').trim();
    const kelas = String(item.kelas || '-').trim();
    const role = String(item.role || 'siswa').toLowerCase().trim();
    const active = item.active !== false;
    if (!norek || !nis || !nama || role !== 'siswa') {
      errors.push(`Baris ${index + 2}: data wajib tidak valid.`);
      return;
    }
    if (noreks.has(norek) || nises.has(nis) || batchNoreks.has(norek) || batchNises.has(nis)) {
      errors.push(`Baris ${index + 2}: norek atau NIS sudah terdaftar/duplikat.`);
      return;
    }
    const salt = Utilities.getUuid();
    rows.push([norek, nis, nama, kelas, hashPassword_(nis, salt), salt, role, active, new Date()]);
    batchNoreks.add(norek);
    batchNises.add(nis);
  });

  if (rows.length) {
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
    } finally {
      lock.releaseLock();
    }
  }
  clearUsersCache_();
  return { success: true, created: rows.length, skipped: errors.length, errors };
}

function bulkCreateTransactions_(actor, transactions) {
  if (!Array.isArray(transactions) || !transactions.length) throw new Error('Data transaksi kosong.');
  if (transactions.length > 500) throw new Error('Maksimal 500 transaksi per upload.');

  const existing = getAllTransactions_();
  const fingerprints = new Set(existing.map(item => transactionFingerprint_(item.nis, item.tanggal, item.jenis, item.jumlah, item.keterangan)));
  const users = getUsers_();
  const byNorek = {};
  const balances = {};
  users.forEach(user => byNorek[user.norek] = user);
  users.forEach(user => balances[user.nis] = calculateBalance_(existing.filter(item => item.nis === user.nis)));
  const rows = [];
  const errors = [];

  transactions.forEach((item, index) => {
    const norek = String(item.norek || '').trim();
    const tanggal = String(item.tanggal || '').trim();
    const jenis = String(item.jenis || '').toUpperCase().trim();
    const jumlah = Number(item.jumlah);
    const keterangan = String(item.keterangan || '').trim().slice(0, 250);
    const user = byNorek[norek];
    if (!user || !/^\d{4}-\d{2}-\d{2}$/.test(tanggal) || !['SIMPAN', 'PAKAI'].includes(jenis) ||
        !Number.isInteger(jumlah) || jumlah <= 0 || jumlah > CONFIG.maxTransactionAmount) {
      errors.push(`Baris ${index + 2}: data transaksi tidak valid.`);
      return;
    }
    const fingerprint = transactionFingerprint_(user.nis, tanggal, jenis, jumlah, keterangan);
    if (fingerprints.has(fingerprint)) {
      errors.push(`Baris ${index + 2}: transaksi duplikat, dilewati.`);
      return;
    }
    if (jenis === 'PAKAI' && jumlah > balances[user.nis]) {
      errors.push(`Baris ${index + 2}: saldo tidak mencukupi, dilewati.`);
      return;
    }
    fingerprints.add(fingerprint);
    balances[user.nis] += jenis === 'SIMPAN' ? jumlah : -jumlah;
    rows.push([Utilities.getUuid(), new Date(), tanggal, user.nis, jenis, jumlah, keterangan, actor.nis]);
  });

  if (rows.length) {
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    try {
      getSheet_(CONFIG.sheets.transactions).getRange(
        getSheet_(CONFIG.sheets.transactions).getLastRow() + 1, 1, rows.length, rows[0].length
      ).setValues(rows);
    } finally {
      lock.releaseLock();
    }
  }
  return { success: true, created: rows.length, skipped: errors.length, errors };
}

function transactionFingerprint_(nis, tanggal, jenis, jumlah, keterangan) {
  return [nis, tanggal, jenis, jumlah, keterangan].join('|');
}

function doGet() {
  return json_({ success: true, service: 'E-Tabungan Pelajar API' });
}

function doPost(event) {
  try {
    const payload = JSON.parse(event.postData.contents || '{}');
    return json_(route_(payload));
  } catch (error) {
    return json_({ success: false, message: error.message || 'Request tidak valid.' });
  }
}

function route_(payload) {
  switch (payload.action) {
    case 'login': return login_(payload);
    case 'getDashboard': return getDashboard_(requireUser_(payload));
    case 'getDashboardData': return getDashboardData_(requireUser_(payload));
    case 'getTransaksi': return getTransactions_(requireUser_(payload));
    case 'getAllRekap': return getAdminRecap_(requireAdmin_(payload));
    case 'bulkCreateUsers': return bulkCreateUsers_(requireAdmin_(payload), payload.users);
    case 'bulkCreateTransactions': return bulkCreateTransactions_(requireAdmin_(payload), payload.transactions);
    case 'createTransaksi': return createTransaction_(requireUser_(payload), payload);
    case 'getDropdownData': return getDropdownData_(requireUser_(payload));
    default: throw new Error('Action tidak dikenal.');
  }
}

function login_(payload) {
  const norek = String(payload.norek || '').trim();
  const password = String(payload.password || '');
  if (!norek || !password) throw new Error('Nomor rekening dan password wajib diisi.');

  const user = findUserByNorek_(norek);
  if (!user || !isActive_(user.active) || hashPassword_(password, user.salt) !== user.passwordHash) {
    throw new Error('Username atau password salah.');
  }

  const safeUser = {
    norek: user.norek,
    nis: user.nis,
    nama: user.nama,
    kelas: user.kelas,
    role: user.role
  };
  return { success: true, user: safeUser, token: issueToken_(safeUser) };
}

function requireUser_(payload) {
  const user = verifyToken_(payload.token);
  if (!user) throw new Error('Sesi tidak valid atau sudah kedaluwarsa.');
  return user;
}

function requireAdmin_(payload) {
  const user = requireUser_(payload);
  if (user.role !== 'admin') throw new Error('Akses admin diperlukan.');
  return user;
}

function getDashboard_(user) {
  const transactions = getTransactionsForNis_(user.nis);
  return { success: true, saldo: calculateBalance_(transactions) };
}

function getDashboardData_(user) {
  const transactions = getTransactionsForNis_(user.nis);
  return { success: true, saldo: calculateBalance_(transactions), data: transactions };
}

function getTransactions_(user) {
  return { success: true, data: getTransactionsForNis_(user.nis) };
}

function createTransaction_(actor, payload) {
  const targetNorek = actor.role === 'admin' ? String(payload.norek || '').trim() : actor.norek;
  const jenis = String(payload.jenis || '').toUpperCase();
  const tanggal = String(payload.tanggal || '').trim();
  const jumlah = Number(payload.jumlah);
  const keterangan = String(payload.keterangan || '').trim().slice(0, 250);

  if (!targetNorek || !/^\d{4}-\d{2}-\d{2}$/.test(tanggal)) throw new Error('Tanggal atau nomor rekening tidak valid.');
  if (!['SIMPAN', 'PAKAI'].includes(jenis)) throw new Error('Jenis transaksi tidak valid.');
  if (!Number.isInteger(jumlah) || jumlah <= 0 || jumlah > CONFIG.maxTransactionAmount) {
    throw new Error('Nominal transaksi tidak valid.');
  }

  const target = findUserByNorek_(targetNorek);
  if (!target || !isActive_(target.active)) throw new Error('Siswa tidak ditemukan atau tidak aktif.');
  if (actor.role !== 'admin' && jenis !== 'PAKAI') throw new Error('Siswa hanya boleh memakai saldo.');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const currentBalance = calculateBalance_(getTransactionsForNis_(target.nis));
    if (jenis === 'PAKAI' && jumlah > currentBalance) throw new Error('Saldo tidak mencukupi.');

    getSheet_(CONFIG.sheets.transactions).appendRow([
      Utilities.getUuid(), new Date(), tanggal, target.nis, jenis, jumlah, keterangan, actor.nis
    ]);
  } finally {
    lock.releaseLock();
  }
  return { success: true, message: 'Transaksi berhasil disimpan.' };
}

function getAdminRecap_(actor) {
  const users = getUsers_();
  const byNis = {};
  users.forEach(user => byNis[user.nis] = user);
  const data = getAllTransactions_().map(item => ({
    ...item,
    nama: byNis[item.nis] ? byNis[item.nis].nama : item.nis,
    norek: byNis[item.nis] ? byNis[item.nis].norek : '',
    kelas: byNis[item.nis] ? byNis[item.nis].kelas : '-'
  }));
  const totalSimpan = data.filter(item => item.jenis === 'SIMPAN')
    .reduce((sum, item) => sum + item.jumlah, 0);
  const totalPakai = data.filter(item => item.jenis === 'PAKAI')
    .reduce((sum, item) => sum + item.jumlah, 0);
  return {
    success: true,
    data,
    totalSimpan,
    totalPakai,
    totalSaldo: totalSimpan - totalPakai
  };
}

function getDropdownData_() {
  const sheet = getSheet_(CONFIG.sheets.dropdowns);
  const rows = sheet.getDataRange().getValues().slice(1);
  return {
    success: true,
    jumlah: rows.filter(row => row[0] === 'jumlah').map(row => row[1]),
    keterangan: rows.filter(row => row[0] === 'keterangan').map(row => row[1])
  };
}

function getTransactionsForNis_(nis) {
  return getAllTransactions_().filter(item => String(item.nis) === String(nis));
}

function getAllTransactions_() {
  const rows = getSheet_(CONFIG.sheets.transactions).getDataRange().getValues().slice(1);
  return rows.filter(row => row[0]).map(row => ({
    id: String(row[0]),
    tanggal: formatDate_(row[2]),
    nis: String(row[3]),
    jenis: String(row[4]).toUpperCase(),
    jumlah: Number(row[5]) || 0,
    keterangan: String(row[6] || '')
  }));
}

function calculateBalance_(transactions) {
  return transactions.reduce((balance, item) =>
    balance + (item.jenis === 'SIMPAN' ? item.jumlah : -item.jumlah), 0);
}

function getUsers_() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('users_cache');
  if (cached) return JSON.parse(cached);

  const rows = getSheet_(CONFIG.sheets.users).getDataRange().getValues().slice(1);
  const users = rows.filter(row => row[0]).map(row => ({
    norek: String(row[0]).trim(), nis: String(row[1]).trim(), nama: String(row[2] || ''), kelas: String(row[3] || '-'),
    passwordHash: String(row[4] || ''), salt: String(row[5] || ''), role: String(row[6] || 'siswa').toLowerCase(), active: row[7]
  }));
  cache.put('users_cache', JSON.stringify(users), 60);
  return users;
}

function clearUsersCache_() {
  CacheService.getScriptCache().remove('users_cache');
}

function findUserByNorek_(norek) {
  return getUsers_().find(user => user.norek === String(norek).trim());
}

function issueToken_(user) {
  const payload = { ...user, exp: Math.floor(Date.now() / 1000) + CONFIG.tokenTtlSeconds };
  const encoded = encode_(JSON.stringify(payload));
  return encoded + '.' + sign_(encoded);
}

function verifyToken_(token) {
  try {
    const parts = String(token || '').split('.');
    if (parts.length !== 2 || sign_(parts[0]) !== parts[1]) return null;
    const payload = JSON.parse(decode_(parts[0]));
    if (!payload.exp || payload.exp < Math.floor(Date.now() / 1000)) return null;
    const user = findUserByNorek_(payload.norek);
    if (!user || !isActive_(user.active) || user.role !== payload.role) return null;
    return { norek: user.norek, nis: user.nis, nama: user.nama, kelas: user.kelas, role: user.role };
  } catch (error) {
    return null;
  }
}

function sign_(value) {
  const secret = PropertiesService.getScriptProperties().getProperty('TOKEN_SECRET');
  if (!secret) throw new Error('TOKEN_SECRET belum dibuat. Jalankan setupSheet().');
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(value, secret));
}

function hashPassword_(password, salt) {
  return Utilities.base64Encode(Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, String(salt) + String(password), Utilities.Charset.UTF_8
  ));
}

function encode_(value) { return Utilities.base64EncodeWebSafe(Utilities.newBlob(value).getBytes()); }
function decode_(value) { return Utilities.newBlob(Utilities.base64DecodeWebSafe(value)).getDataAsString(); }
function isActive_(value) { return value === true || String(value).toLowerCase() === 'true'; }
function formatDate_(value) { return value instanceof Date ? Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd') : String(value || ''); }

function getSheet_(name) {
  const properties = PropertiesService.getScriptProperties();
  const spreadsheetId = CONFIG.spreadsheetId || properties.getProperty('SPREADSHEET_ID');
  const spreadsheet = SpreadsheetApp.openById(spreadsheetId || CONFIG.spreadsheetId);
  if (!spreadsheet) throw new Error('Spreadsheet belum dikonfigurasi. Jalankan setupSheet().');
  const sheet = spreadsheet.getSheetByName(name || CONFIG.sheets.users);
  if (!sheet) throw new Error('Sheet belum siap. Jalankan setupSheet().');
  return sheet;
}

function getOrCreateSheet_(spreadsheet, name, headers) {
  const sheet = spreadsheet.getSheetByName(name) || spreadsheet.insertSheet(name);
  if (sheet.getLastRow() === 0) sheet.appendRow(headers);
  return sheet;
}

function migrateUsersSheet_(sheet) {
  const headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0]
    .map(value => String(value).trim().toLowerCase());
  if (headers[0] === 'norek') return;
  if (headers[0] !== 'nis') throw new Error('Format sheet Users tidak dikenali.');

  sheet.insertColumnBefore(1);
  sheet.getRange(1, 1).setValue('norek');
  const rowCount = sheet.getLastRow();
  if (rowCount > 1) {
    const oldNis = sheet.getRange(2, 2, rowCount - 1, 1).getValues();
    sheet.getRange(2, 1, rowCount - 1, 1).setValues(oldNis);
  }
}

function json_(value) {
  return ContentService.createTextOutput(JSON.stringify(value)).setMimeType(ContentService.MimeType.JSON);
}
