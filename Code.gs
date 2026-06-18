/* ===============================
   BACKEND SISTEM ADMIN GYM
   Google Apps Script + Google Sheet
   PATCH FINAL:
   - Fix bug: status kunci tidak update ke DATA_KUNCI (web selalu "Kosong")
   - Fix bug: sync bulanan bikin transaksi timeout
   - Cowo = HITAM, Cewe = MERAH
   =============================== */

// WAJIB: Isi cuma ID Google Sheet, bukan link full dan bukan link Apps Script.
const SPREADSHEET_ID = '1J4YOaOLhwZ2fb4CpWUSMct8TnGX6oxfU5EO5tyG3AZw';

/* ===============================
   FUNGSI TES / DIAGNOSA
   Jalankan ini DULUAN dari dropdown editor untuk cek koneksi & izin.
   Error di sini TIDAK dibungkus, jadi pesannya jelas.
   =============================== */

function tesKoneksi() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  Logger.log('OK. Nama file: ' + ss.getName());
  Logger.log('Jumlah sheet: ' + ss.getSheets().length);
  ss.getSheets().forEach(function (s) {
    Logger.log(' - ' + s.getName());
  });
  return 'Koneksi & izin OK: ' + ss.getName();
}

// cekVersi: pastikan file di editor benar-benar versi baru.
function cekVersi() {
  Logger.log('VERSI AKTIF: 2026-06-18-fix-merge-v3');
  return '2026-06-18-fix-merge-v3';
}

// scanMerge: nge-scan SEMUA sheet, laporkan sheet mana yang punya merged cells.
// Jalankan dari editor, lalu lihat Log eksekusi. Ini menunjuk biang merge.
function scanMerge() {
  const ss = getSpreadsheet_();
  let total = 0;

  ss.getSheets().forEach(function (sheet) {
    let merges = [];
    try {
      merges = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
        .getMergedRanges();
    } catch (error) {
      Logger.log('! ' + sheet.getName() + ' gagal discan: ' + error.message);
      return;
    }

    if (merges.length) {
      total += merges.length;
      Logger.log('MERGE di "' + sheet.getName() + '": ' + merges.length + ' range');
      merges.slice(0, 5).forEach(function (r) {
        Logger.log('    ' + r.getA1Notation());
      });
    } else {
      Logger.log('bersih: ' + sheet.getName());
    }
  });

  Logger.log('=== TOTAL merge ditemukan: ' + total + ' ===');
  return 'Total merge: ' + total;
}

// bersihkanMergeOperasional: hapus SEMUA merge di sheet operasional utama
// (DATA_KUNCI, LOG_GYM, REKAP_HARIAN, MEMBER_LIFETIME). Tidak menyentuh
// sheet bulanan. Pakai kalau scanMerge nemu merge di sheet operasional.
function bersihkanMergeOperasional() {
  const ss = getSpreadsheet_();
  const names = [SHEET_KEYS, SHEET_LOG, SHEET_DAILY, SHEET_MEMBERS];
  let dibersihkan = 0;

  names.forEach(function (name) {
    const sheet = ss.getSheetByName(name);
    if (!sheet) return;

    try {
      const merges = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
        .getMergedRanges();
      merges.forEach(function (r) {
        try { r.breakApart(); dibersihkan++; } catch (e) {}
      });
    } catch (error) {
      Logger.log('Gagal bersihkan ' + name + ': ' + error.message);
    }
  });

  return 'Selesai. Merge dibuka: ' + dibersihkan;
}

// bersihkanSemuaMerge: HAPUS SEMUA merge di SETIAP sheet di spreadsheet,
// termasuk semua sheet bulanan (12 bulan) dan DASHBOARD.
// Ini perkakas penyelamat pamungkas. Jalankan SEKALI dari editor.
// Data tidak dihapus, hanya merge yang dibuka.
function bersihkanSemuaMerge() {
  const ss = getSpreadsheet_();
  let totalDibuka = 0;
  let sheetTersentuh = 0;

  ss.getSheets().forEach(function (sheet) {
    let merges = [];
    try {
      merges = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
        .getMergedRanges();
    } catch (error) {
      Logger.log('! ' + sheet.getName() + ' gagal: ' + error.message);
      return;
    }

    if (!merges.length) return;

    sheetTersentuh++;
    merges.forEach(function (r) {
      try { r.breakApart(); totalDibuka++; } catch (e) {}
    });
    Logger.log('dibersihkan: ' + sheet.getName() + ' (' + merges.length + ' merge)');
  });

  Logger.log('=== Selesai. ' + totalDibuka + ' merge dibuka di ' + sheetTersentuh + ' sheet ===');
  return 'Selesai. ' + totalDibuka + ' merge dibuka di ' + sheetTersentuh + ' sheet.';
}

// perbaikiStrukturHarian: rapikan kolom REKAP_HARIAN supaya cocok 100% dengan
// DAILY_HEADERS yang dipakai kode. Aman dijalankan berkali-kali (idempotent).
//
// Struktur yang BENAR (12 kolom):
//   A Tanggal | B No | C Nama | D No Kunci | E (kosong) | F Jam Masuk
//   G Admin Masuk | H Jam Keluar | I Sudah Keluar | J Waktu Keluar Lengkap
//   K Admin Keluar | L Jenis Kunci
//
// Masalah yang diperbaiki: struktur geser karena ada kolom ekstra
// "Waktu Masuk Lengkap" yang nyelip di antara Jam Masuk dan Admin Masuk.
// Fungsi ini mendeteksi kolom itu lewat HEADER (bukan posisi), lalu menghapusnya,
// memastikan ada kolom kosong di E, dan menulis ulang header yang benar.
function perbaikiStrukturHarian() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(SHEET_DAILY);
  if (!sheet) throw new Error('Sheet REKAP_HARIAN tidak ada.');

  breakAllMerges_(sheet);

  const langkah = [];
  let lastRow = sheet.getLastRow();
  let lastCol = Math.max(sheet.getLastColumn(), DAILY_HEADERS.length);
  let header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(cleanHeader_);

  // Deteksi kolom-kolom kunci BERDASARKAN HEADER (bukan posisi).
  const colJamMasuk = findHeaderIndex_(header, ['jam masuk']);          // index 0-based
  const colWaktuMasukLengkap = findHeaderIndex_(header, ['waktu masuk lengkap']);

  // LANGKAH A: kalau kolom "Waktu Masuk Lengkap" ada DAN ada kolom "Jam Masuk"
  // terpisah, pindahkan jam dari "Waktu Masuk Lengkap" ke "Jam Masuk" untuk
  // baris yang Jam Masuk-nya kosong. Ini menyelamatkan data lama.
  if (colWaktuMasukLengkap !== -1 && colJamMasuk !== -1 && lastRow >= 2) {
    const jamRange = sheet.getRange(2, colJamMasuk + 1, lastRow - 1, 1);
    const wmlRange = sheet.getRange(2, colWaktuMasukLengkap + 1, lastRow - 1, 1);

    const jamVals = jamRange.getValues();
    const wmlVals = wmlRange.getValues();

    let dipindah = 0;
    for (let r = 0; r < jamVals.length; r++) {
      const jamKosong = String(jamVals[r][0] || '').trim() === '';
      const wmlAda = String(wmlVals[r][0] || '').trim() !== '';

      if (jamKosong && wmlAda) {
        // Ambil bagian jam dari "Waktu Masuk Lengkap" (mis. "11:42:53" atau "18/06/2026 11:42:53").
        jamVals[r][0] = stringifyTimeOnly_(wmlVals[r][0]);
        dipindah++;
      }
    }

    if (dipindah > 0) {
      jamRange.setValues(jamVals);
      langkah.push('pindah jam masuk dari "Waktu Masuk Lengkap" ke "Jam Masuk" untuk ' + dipindah + ' baris');
    }
  }

  // LANGKAH B: hapus kolom "Waktu Masuk Lengkap" (kolom nyasar). Dari kanan.
  header = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), DAILY_HEADERS.length)).getValues()[0].map(cleanHeader_);
  for (let c = header.length - 1; c >= 0; c--) {
    if (header[c] === 'waktu masuk lengkap') {
      sheet.deleteColumn(c + 1);
      langkah.push('hapus kolom "Waktu Masuk Lengkap" di posisi ' + (c + 1));
    }
  }

  // Refresh header.
  lastCol = Math.max(sheet.getLastColumn(), DAILY_HEADERS.length);
  header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(cleanHeader_);

  // LANGKAH C: struktur target butuh kolom KOSONG di posisi E (index 4),
  // dengan "Jam Masuk" di F (index 5). Kalau sekarang "Jam Masuk" ada di E,
  // sisipkan satu kolom kosong sebelum E untuk menggeser semuanya ke kanan.
  if (header[4] === 'jam masuk') {
    sheet.insertColumnBefore(5);
    langkah.push('sisip kolom kosong di posisi E (5)');
    lastCol = Math.max(sheet.getLastColumn(), DAILY_HEADERS.length);
    header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(cleanHeader_);
  }

  // LANGKAH D: pastikan jumlah kolom cukup, tulis ulang header yang benar.
  if (sheet.getMaxColumns() < DAILY_HEADERS.length) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), DAILY_HEADERS.length - sheet.getMaxColumns());
  }
  sheet.getRange(1, 1, 1, DAILY_HEADERS.length).setValues([DAILY_HEADERS]);
  styleHeader_(sheet, DAILY_HEADERS.length);
  sheet.setFrozenRows(1);

  // LANGKAH E: pasang ulang checkbox di kolom I (Sudah Keluar).
  lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    try {
      sheet.getRange(2, 9, lastRow - 1, 1).insertCheckboxes();
    } catch (error) {
      Logger.log('checkbox dilewati: ' + error.message);
    }
  }

  // LANGKAH F: warnai ulang kolom Jenis Kunci (L = 12).
  try { styleKeyTypeColumn_(sheet, 12); } catch (error) {}

  if (!langkah.length) {
    Logger.log('REKAP_HARIAN sudah rapi, tidak ada yang diubah.');
    return 'REKAP_HARIAN sudah rapi.';
  }

  langkah.forEach(function (s) { Logger.log('- ' + s); });
  Logger.log('=== Struktur REKAP_HARIAN dirapikan. ===');
  return 'Struktur REKAP_HARIAN dirapikan: ' + langkah.join('; ');
}

// perbaikiSemua: SATU TOMBOL untuk beresin semuanya sekaligus.
// Jalankan ini dari editor kalau masih ada toast merah / data berantakan.
// Urutan: buka semua merge -> rapikan struktur harian -> reset rekap bulanan.
function perbaikiSemua() {
  const hasil = [];

  try { hasil.push('Merge: ' + bersihkanSemuaMerge()); }
  catch (e) { hasil.push('Merge GAGAL: ' + e.message); }

  try { hasil.push('Struktur: ' + perbaikiStrukturHarian()); }
  catch (e) { hasil.push('Struktur GAGAL: ' + e.message); }

  try { hasil.push('Bulanan: ' + resetRekapBulanan()); }
  catch (e) { hasil.push('Bulanan GAGAL: ' + e.message); }

  hasil.forEach(function (h) { Logger.log(h); });
  Logger.log('=== SEMUA PERBAIKAN SELESAI ===');
  return hasil.join(' || ');
}



const TIMEZONE = 'Asia/Jakarta';

const MAX_KEY_NUMBER = 100;

// KEY_TYPES: urutan Cowo dulu, lalu Cewe
const KEY_TYPES = ['Cowo', 'Cewe'];

// === WARNA JENIS KUNCI ===
// Cowo = hitam, Cewe = merah
const COLOR_COWO = '#111827';
const COLOR_CEWE = '#dc2626';

const SHEET_LOG = 'LOG_GYM';
const SHEET_KEYS = 'DATA_KUNCI';
const SHEET_MEMBERS = 'MEMBER_LIFETIME';
const SHEET_DAILY = 'REKAP_HARIAN';

const LOG_HEADERS = [
  'No',
  'Waktu Lengkap',
  'Tanggal',
  'Jam',
  'Nama',
  'No Kunci',
  'Status',
  'Admin',
  'Jenis Kunci'
];

const KEY_HEADERS = [
  'No Kunci',
  'Status',
  'Dipakai Oleh',
  'Jam Masuk',
  'Update Terakhir',
  'Jenis Kunci'
];

const MEMBER_HEADERS = [
  'No',
  'Nama Member',
  'Status',
  'Tanggal Daftar',
  'Diinput Oleh',
  'Update Terakhir'
];

// Struktur REKAP_HARIAN:
// A Tanggal | B No | C Nama | D No Kunci | E kosong | F Jam Masuk
// G Admin Masuk | H Jam Keluar | I Sudah Keluar | J Waktu Keluar Lengkap
// K Admin Keluar | L Jenis Kunci
const DAILY_HEADERS = [
  'Tanggal',
  'No',
  'Nama',
  'No Kunci',
  '',
  'Jam Masuk',
  'Admin Masuk',
  'Jam Keluar',
  'Sudah Keluar',
  'Waktu Keluar Lengkap',
  'Admin Keluar',
  'Jenis Kunci'
];

/* ===============================
   KONFIG REKAP BULANAN
   =============================== */

const MONTHLY_BLOCK_WIDTH = 12;
const MONTHLY_GAP_WIDTH = 1;
const MONTHLY_START_ROW = 3;
const MONTHLY_VISIBLE_ROWS = 45;

const MONTH_ID = [
  'JANUARI', 'FEBRUARI', 'MARET', 'APRIL', 'MEI', 'JUNI',
  'JULI', 'AGUSTUS', 'SEPTEMBER', 'OKTOBER', 'NOVEMBER', 'DESEMBER'
];

const MONTH_DISPLAY = [
  'Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni',
  'Juli', 'Agustus', 'September', 'Oktober', 'November', 'Desember'
];

const TARGET_MONTH_INDEXES = [5, 6, 7, 8, 9];

// Property key untuk menandai bahwa sync bulanan perlu dijalankan ulang.
const PROP_NEEDS_SYNC = 'GYM_NEEDS_MONTHLY_SYNC';

/* ===============================
   SETUP SHEETS
   =============================== */

function setupGymSheets() {
  ensureReady_();

  const ss = getSpreadsheet_();

  const logSheet = getOrCreateSheet_(ss, SHEET_LOG);
  const keySheet = getOrCreateSheet_(ss, SHEET_KEYS);
  const memberSheet = getOrCreateSheet_(ss, SHEET_MEMBERS);
  const dailySheet = getOrCreateSheet_(ss, SHEET_DAILY);

  setupHeader_(logSheet, LOG_HEADERS);
  setupHeader_(keySheet, KEY_HEADERS);
  setupHeader_(memberSheet, MEMBER_HEADERS);
  setupDailyHeader_(dailySheet);

  seedKeys_(keySheet, MAX_KEY_NUMBER);
  setupDailyCheckboxes_(dailySheet);

  logSheet.setFrozenRows(1);
  keySheet.setFrozenRows(1);
  memberSheet.setFrozenRows(1);
  dailySheet.setFrozenRows(1);

  // autoResize_ kadang lambat / error di sheet besar. Bungkus supaya tidak
  // menggagalkan setup utama yang sudah berhasil.
  try {
    autoResize_(logSheet, LOG_HEADERS.length);
    autoResize_(keySheet, KEY_HEADERS.length);
    autoResize_(memberSheet, MEMBER_HEADERS.length);
    autoResize_(dailySheet, DAILY_HEADERS.length);
  } catch (error) {
    Logger.log('autoResize dilewati: ' + error.message);
  }

  return 'Setup selesai. DATA_KUNCI sekarang 200 slot (Cowo 01-100 + Cewe 01-100).';
}

/* ===============================
   API GET
   =============================== */

function doGet(e) {
  const params = e && e.parameter ? e.parameter : {};

  try {
    ensureReady_();

    const action = String(params.action || 'ping').trim().toLowerCase();

    if (action === 'ping') {
      return respondJson_(params.callback, {
        ok: true,
        message: 'Backend Sistem Admin Gym aktif.',
        data: { app: 'Sistem Admin Gym', serverTime: formatDateTime_(new Date()), version: '2026-06-18-fix-merge-v3' }
      });
    }

    if (action === 'cekversi') {
      return respondJson_(params.callback, {
        ok: true,
        message: 'Versi backend aktif.',
        data: { version: '2026-06-18-fix-merge-v3' }
      });
    }

    if (action === 'keys') {
      return respondJson_(params.callback, safeRead_('Data kunci', getKeys_));
    }

    if (action === 'members') {
      return respondJson_(params.callback, safeRead_('Data member lifetime', getMembers_));
    }

    if (action === 'logs') {
      return respondJson_(params.callback, safeRead_('Data audit', getLogs_));
    }

    if (action === 'daily') {
      return respondJson_(params.callback, safeRead_('Data rekap harian', getDailyRecap_));
    }

    if (action === 'setupmonthly') {
      const result = setupRekapBulananDariHarian();
      return respondJson_(params.callback, { ok: result.success, message: result.message });
    }

    if (action === 'syncmonthly') {
      const result = syncSemuaRekapHarianKeBulanan();
      return respondJson_(params.callback, { ok: result.success, message: result.message });
    }

    if (action === 'savelog') {
      const result = saveLogFromParams_(params);
      // Sync bulanan dijalankan di luar jalur kritis. Kalau gagal, transaksi
      // utama TETAP sukses dan tidak memunculkan error ke pengguna.
      try { flushPendingMonthlySync_(); } catch (e) { Logger.log('flush diabaikan: ' + e.message); }
      return respondJson_(params.callback, { ok: true, message: result.message });
    }

    return respondJson_(params.callback, { ok: false, message: 'Action tidak dikenal: ' + action });

  } catch (error) {
    return respondJson_(params.callback, { ok: false, message: error.message || String(error) });
  }
}

/* ===============================
   API POST
   =============================== */

function doPost(e) {
  const params = e && e.parameter ? e.parameter : {};

  try {
    const result = saveLogFromParams_(params);
    try { flushPendingMonthlySync_(); } catch (e) { Logger.log('flush diabaikan: ' + e.message); }
    return respondPostMessage_({ ok: true, message: result.message });
  } catch (error) {
    return respondPostMessage_({ ok: false, message: error.message || String(error) });
  }
}

/* ===============================
   SAVE LOG MASUK / KELUAR
   PERBAIKAN UTAMA ADA DI SINI:
   1. updateKey_ dipanggil DULUAN (sebelum recap & log) supaya status
      DATA_KUNCI pasti tercatat walau eksekusi berikutnya gagal.
   2. Sync bulanan TIDAK dipanggil di dalam lock. Hanya ditandai,
      lalu dijalankan setelah lock dilepas.
   =============================== */

function saveLogFromParams_(params) {
  const lock = LockService.getScriptLock();
  let locked = false;

  try {
    lock.waitLock(20000);
    locked = true;

    ensureReady_();

    const payload = normalizePayload_(params);
    const ss = getSpreadsheet_();

    const logSheet = getOrCreateSheet_(ss, SHEET_LOG);
    const keySheet = getOrCreateSheet_(ss, SHEET_KEYS);
    const dailySheet = getOrCreateSheet_(ss, SHEET_DAILY);

    // Jaring pengaman: kalau ada merge nyangkut di sheet operasional, buka
    // dulu supaya setValues/appendRow tidak gagal dengan error
    // "Anda harus memilih semua sel dalam rentang penggabungan".
    safeBreakMerges_(logSheet);
    safeBreakMerges_(keySheet);
    safeBreakMerges_(dailySheet);

    setupHeader_(logSheet, LOG_HEADERS);
    setupHeader_(keySheet, KEY_HEADERS);
    setupDailyHeader_(dailySheet);

    const currentKey = getKeyRecord_(keySheet, payload.keyNumber, payload.keyType);
    const currentKeyStatus = cleanText_(currentKey.status).toLowerCase();

    if (payload.status === 'Masuk') {
      if (currentKeyStatus === 'dipakai') {
        throw new Error(
          'Kunci ' + payload.keyType + ' ' + payload.keyNumber +
          ' sedang dipakai oleh ' + (currentKey.customerName || 'pelanggan lain') + '.'
        );
      }

      // (1) Update DATA_KUNCI lebih dulu -> ini yang dibaca web.
      updateKey_(keySheet, payload);
      // (2) Baru catat log & rekap harian.
      appendLog_(logSheet, payload);
      updateDailyRecap_(dailySheet, payload);

      markNeedsMonthlySync_();

      return {
        ok: true,
        message: 'Data masuk berhasil disimpan untuk kunci ' + payload.keyType + ' ' + payload.keyNumber + '.'
      };
    }

    if (payload.status === 'Keluar') {
      if (currentKeyStatus !== 'dipakai') {
        throw new Error(
          'Kunci ' + payload.keyType + ' ' + payload.keyNumber + ' belum tercatat sedang dipakai.'
        );
      }

      const checkoutPayload = {
        admin: payload.admin,
        customerName: currentKey.customerName || payload.customerName || 'Tanpa Nama',
        keyNumber: payload.keyNumber,
        keyType: payload.keyType,
        status: 'Keluar',
        timestamp: payload.timestamp,
        previousCheckInTime: currentKey.checkInTime || ''
      };

      // (1) Bebaskan kunci di DATA_KUNCI lebih dulu.
      updateKey_(keySheet, checkoutPayload);
      // (2) Catat log & rekap harian.
      appendLog_(logSheet, checkoutPayload);
      updateDailyRecap_(dailySheet, checkoutPayload);

      markNeedsMonthlySync_();

      return {
        ok: true,
        message: 'Data keluar berhasil disimpan untuk kunci ' + payload.keyType + ' ' + payload.keyNumber + '.'
      };
    }

    throw new Error('Status tidak valid.');

  } finally {
    if (locked) {
      lock.releaseLock();
    }
  }
}

/* ===============================
   SYNC BULANAN: DITUNDA & DI LUAR LOCK
   =============================== */

function markNeedsMonthlySync_() {
  try {
    PropertiesService.getScriptProperties().setProperty(PROP_NEEDS_SYNC, '1');
  } catch (error) {
    Logger.log('Gagal set flag sync: ' + error.message);
  }
}

// Dipanggil setelah lock dilepas. Aman kalau gagal — transaksi utama sudah sukses.
function flushPendingMonthlySync_() {
  try {
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty(PROP_NEEDS_SYNC) !== '1') return;

    props.deleteProperty(PROP_NEEDS_SYNC);
    syncSemuaRekapHarianKeBulanan();
  } catch (error) {
    Logger.log('Sync bulanan tertunda gagal (tidak fatal): ' + error.message);
  }
}

/* ===============================
   BASIC HELPERS
   =============================== */

function ensureReady_() {
  if (!SPREADSHEET_ID || SPREADSHEET_ID === 'PASTE_GOOGLE_SHEET_ID_HERE') {
    throw new Error('SPREADSHEET_ID belum diisi di Code.gs.');
  }

  if (
    String(SPREADSHEET_ID).includes('/edit') ||
    String(SPREADSHEET_ID).includes('docs.google.com') ||
    String(SPREADSHEET_ID).includes('script.google.com')
  ) {
    throw new Error('SPREADSHEET_ID salah. Isi cuma ID Google Sheet.');
  }
}

function getSpreadsheet_() {
  return SpreadsheetApp.openById(SPREADSHEET_ID);
}

function getOrCreateSheet_(ss, name) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  return sheet;
}

function setupHeader_(sheet, headers) {
  const neededCols = headers.length;

  if (sheet.getMaxColumns() < neededCols) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), neededCols - sheet.getMaxColumns());
  }

  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), neededCols);

  if (lastRow === 0) {
    sheet.getRange(1, 1, 1, neededCols).setValues([headers]);
    styleHeader_(sheet, neededCols);
    return;
  }

  const current = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const rowEmpty = current.every(function (v) { return String(v || '').trim() === ''; });

  if (rowEmpty) {
    sheet.getRange(1, 1, 1, neededCols).setValues([headers]);
    styleHeader_(sheet, neededCols);
    return;
  }

  for (let i = 0; i < headers.length; i++) {
    const existing = String(current[i] || '').trim();
    if (!existing && headers[i]) {
      sheet.getRange(1, i + 1).setValue(headers[i]);
    }
  }
}

function setupDailyHeader_(sheet) {
  const neededCols = DAILY_HEADERS.length;

  if (sheet.getMaxColumns() < neededCols) {
    sheet.insertColumnsAfter(sheet.getMaxColumns(), neededCols - sheet.getMaxColumns());
  }

  const lastRow = sheet.getLastRow();

  if (lastRow === 0) {
    sheet.getRange(1, 1, 1, neededCols).setValues([DAILY_HEADERS]);
    styleHeader_(sheet, neededCols);
    sheet.setFrozenRows(1);
    return;
  }

  const lastCol = Math.max(sheet.getLastColumn(), neededCols);
  const current = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const rowEmpty = current.every(function (v) { return String(v || '').trim() === ''; });

  if (rowEmpty) {
    sheet.getRange(1, 1, 1, neededCols).setValues([DAILY_HEADERS]);
    styleHeader_(sheet, neededCols);
  }

  for (let i = 0; i < DAILY_HEADERS.length; i++) {
    const existing = String(current[i] || '').trim();
    if (!existing && DAILY_HEADERS[i]) {
      sheet.getRange(1, i + 1).setValue(DAILY_HEADERS[i]);
    }
  }

  sheet.setFrozenRows(1);
}

function setupDailyCheckboxes_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  sheet.getRange(2, 9, lastRow - 1, 1).insertCheckboxes();
}

function styleHeader_(sheet, length) {
  sheet.getRange(1, 1, 1, length)
    .setFontWeight('bold')
    .setBackground('#eaf1ff')
    .setFontColor('#111827')
    .setHorizontalAlignment('center');
}

/* ===============================
   SEED KEYS (200 SLOT)
   =============================== */

function seedKeys_(sheet, maxKey) {
  setupHeader_(sheet, KEY_HEADERS);

  const existingKeys = new Set();
  const lastRow = sheet.getLastRow();

  if (lastRow >= 2) {
    const values = sheet.getRange(2, 1, lastRow - 1, KEY_HEADERS.length).getValues();

    values.forEach(function (row, index) {
      const rowIndex = index + 2;
      const key = normalizeKeyNumber_(row[0]);
      let keyType = normalizeKeyType_(row[5]);

      if (key && !cleanText_(row[5])) {
        keyType = 'Cowo';
        sheet.getRange(rowIndex, 6).setValue(keyType);
        styleKeyTypeCell_(sheet, rowIndex, 6, keyType);
      }

      if (key) {
        existingKeys.add(makeKeyIdentity_(keyType, key));
      }
    });
  }

  const rows = [];

  KEY_TYPES.forEach(function (keyType) {
    for (let i = 1; i <= maxKey; i++) {
      const key = String(i).padStart(2, '0');
      const identity = makeKeyIdentity_(keyType, key);
      if (!existingKeys.has(identity)) {
        rows.push([key, 'Kosong', '', '', '', keyType]);
      }
    }
  });

  if (rows.length) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, KEY_HEADERS.length).setValues(rows);
    // Warnai hanya kalau ada baris baru (hemat waktu eksekusi).
    styleKeyTypeColumn_(sheet, 6);
  }
}

/* ===============================
   PAYLOAD NORMALIZATION
   =============================== */

function normalizePayload_(params) {
  const admin = cleanText_(params.admin);
  const customerName = cleanText_(params.customerName || params.nama || params.namaPelanggan);
  const keyNumber = normalizeKeyNumber_(params.keyNumber || params.noKunci || params.nomorKunci);
  const keyType = normalizeKeyType_(params.keyType || params.jenisKunci || params.gender || params.tipeKunci);
  const status = normalizeStatus_(params.status);

  if (!admin) throw new Error('Nama admin/pegawai wajib diisi.');
  if (!keyNumber) throw new Error('Nomor kunci wajib diisi.');
  if (!keyType) throw new Error('Jenis kunci wajib dipilih: Cowo atau Cewe.');
  if (!status) throw new Error('Status tidak valid.');
  if (status === 'Masuk' && !customerName) throw new Error('Nama pelanggan wajib diisi untuk check-in.');

  return {
    admin: admin,
    customerName: customerName,
    keyNumber: keyNumber,
    keyType: keyType,
    status: status,
    timestamp: new Date()
  };
}

function cleanText_(value) {
  return String(value || '').trim().replace(/\s+/g, ' ');
}

function normalizeKeyNumber_(value) {
  const raw = String(value || '').trim();
  if (!raw) return '';
  const number = Number(raw);
  if (!Number.isFinite(number) || number <= 0) return '';
  return String(Math.floor(number)).padStart(2, '0');
}

function normalizeStatus_(value) {
  const raw = String(value || '').trim().toLowerCase();
  if (raw === 'masuk') return 'Masuk';
  if (raw === 'keluar') return 'Keluar';
  return '';
}

function normalizeKeyType_(value) {
  const raw = String(value || '').trim().toLowerCase();
  if (!raw) return 'Cowo';
  if (['cowo', 'cowok', 'pria', 'laki', 'laki-laki', 'male'].indexOf(raw) !== -1) return 'Cowo';
  if (['cewe', 'cewek', 'wanita', 'perempuan', 'female'].indexOf(raw) !== -1) return 'Cewe';
  return '';
}

function makeKeyIdentity_(keyType, keyNumber) {
  return normalizeKeyType_(keyType) + '__' + normalizeKeyNumber_(keyNumber);
}

/* ===============================
   LOG_GYM
   =============================== */

function appendLog_(sheet, payload) {
  const timestamp = payload.timestamp || new Date();
  const no = Math.max(sheet.getLastRow(), 1);

  sheet.appendRow([
    no,
    formatDateTime_(timestamp),
    formatDate_(timestamp),
    formatTime_(timestamp),
    payload.customerName,
    payload.keyNumber,
    payload.status,
    payload.admin,
    payload.keyType || 'Cowo'
  ]);

  styleKeyTypeCell_(sheet, sheet.getLastRow(), 9, payload.keyType);
}

/* ===============================
   REKAP_HARIAN MASTER
   =============================== */

function updateDailyRecap_(sheet, payload) {
  setupDailyHeader_(sheet);
  setupDailyCheckboxes_(sheet);

  if (payload.status === 'Masuk') {
    appendDailyCheckIn_(sheet, payload);
    return;
  }

  if (payload.status === 'Keluar') {
    markDailyCheckout_(sheet, payload);
  }
}

function appendDailyCheckIn_(sheet, payload) {
  const timestamp = payload.timestamp || new Date();
  const tanggal = formatDate_(timestamp);
  const jamMasuk = formatTime_(timestamp);
  const nomorHarian = getNextDailyNumber_(sheet, tanggal);

  const rowValues = [
    tanggal, nomorHarian, payload.customerName, payload.keyNumber, '',
    jamMasuk, payload.admin, '', false, '', '', payload.keyType || 'Cowo'
  ];

  sheet.appendRow(rowValues);

  const rowIndex = sheet.getLastRow();
  sheet.getRange(rowIndex, 9).insertCheckboxes();
  sheet.getRange(rowIndex, 9).setValue(false);
  styleKeyTypeCell_(sheet, rowIndex, 12, payload.keyType);
  // CATATAN: tidak ada safeSyncMonthly_ di sini lagi. Sync ditunda.
}

function markDailyCheckout_(sheet, payload) {
  const timestamp = payload.timestamp || new Date();
  const jamKeluar = formatTime_(timestamp);
  const waktuKeluarLengkap = formatDateTime_(timestamp);

  const rowIndex = findOpenDailyRow_(sheet, payload.keyNumber, payload.keyType);

  if (rowIndex) {
    sheet.getRange(rowIndex, 8).setValue(jamKeluar);
    sheet.getRange(rowIndex, 9).insertCheckboxes();
    sheet.getRange(rowIndex, 9).setValue(true);
    sheet.getRange(rowIndex, 10).setValue(waktuKeluarLengkap);
    sheet.getRange(rowIndex, 11).setValue(payload.admin);
    return;
  }

  appendRecoveryCheckout_(sheet, payload);
}

function appendRecoveryCheckout_(sheet, payload) {
  const timestamp = payload.timestamp || new Date();
  const tanggal = formatDate_(timestamp);
  const jamKeluar = formatTime_(timestamp);
  const waktuKeluarLengkap = formatDateTime_(timestamp);
  const nomorHarian = getNextDailyNumber_(sheet, tanggal);

  const previousCheckInTime = cleanText_(payload.previousCheckInTime);
  const jamMasuk = extractTimeFromDateTimeText_(previousCheckInTime);

  const rowValues = [
    tanggal, nomorHarian, payload.customerName, payload.keyNumber, '',
    jamMasuk || '-', 'RECOVERY', jamKeluar, true, waktuKeluarLengkap,
    payload.admin, payload.keyType || 'Cowo'
  ];

  sheet.appendRow(rowValues);

  const rowIndex = sheet.getLastRow();
  sheet.getRange(rowIndex, 9).insertCheckboxes();
  sheet.getRange(rowIndex, 9).setValue(true);
  styleKeyTypeCell_(sheet, rowIndex, 12, payload.keyType);
}

function findOpenDailyRow_(sheet, keyNumber, keyType) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;

  const lastCol = Math.max(sheet.getLastColumn(), DAILY_HEADERS.length);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(cleanHeader_);

  const colKunci = findHeaderIndex_(headers, ['no kunci', 'nomor kunci']);
  const colType = findHeaderIndex_(headers, ['jenis kunci', 'gender', 'tipe kunci']);
  const colKeluar = findHeaderIndex_(headers, ['sudah keluar']);

  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  const targetIdentity = makeKeyIdentity_(keyType, keyNumber);

  for (let i = values.length - 1; i >= 0; i--) {
    const row = values[i];
    const rawKunci = getRawRowValue_(row, colKunci);
    const rowKunci = normalizeKeyNumber_(rawKunci) || cleanText_(rawKunci);
    const rowType = normalizeKeyType_(getRawRowValue_(row, colType));
    const rowIdentity = makeKeyIdentity_(rowType, rowKunci);
    const sudahKeluar = isChecked_(getRawRowValue_(row, colKeluar));

    if (rowIdentity === targetIdentity && !sudahKeluar) {
      return i + 2;
    }
  }

  return null;
}

function isChecked_(value) {
  if (value === true) return true;
  const text = String(value || '').trim().toLowerCase();
  return text === 'true' || text === 'yes' || text === 'checked' || text === 'centang';
}

function getNextDailyNumber_(sheet, tanggal) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return 1;

  const values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
  let count = 0;

  values.forEach(function (row) {
    if (stringifyDateOnly_(row[0]) === tanggal) count++;
  });

  return count + 1;
}

/* ===============================
   DATA_KUNCI
   =============================== */

function updateKey_(sheet, payload) {
  const record = getKeyRecord_(sheet, payload.keyNumber, payload.keyType);
  const rowIndex = record.rowIndex || appendKeyRow_(sheet, payload.keyNumber, payload.keyType);
  const nowText = formatDateTime_(new Date());

  if (payload.status === 'Masuk') {
    sheet.getRange(rowIndex, 1, 1, KEY_HEADERS.length).setValues([[
      payload.keyNumber, 'Dipakai', payload.customerName,
      formatDateTime_(payload.timestamp || new Date()), nowText, payload.keyType || 'Cowo'
    ]]);
  } else {
    sheet.getRange(rowIndex, 1, 1, KEY_HEADERS.length).setValues([[
      payload.keyNumber, 'Kosong', '', '', nowText, payload.keyType || 'Cowo'
    ]]);
  }

  styleKeyTypeCell_(sheet, rowIndex, 6, payload.keyType);
}

function appendKeyRow_(sheet, keyNumber, keyType) {
  const rowIndex = sheet.getLastRow() + 1;

  sheet.getRange(rowIndex, 1, 1, KEY_HEADERS.length).setValues([[
    keyNumber, 'Kosong', '', '', '', keyType || 'Cowo'
  ]]);

  styleKeyTypeCell_(sheet, rowIndex, 6, keyType);
  return rowIndex;
}

function getKeyRecord_(sheet, keyNumber, keyType) {
  const lastRow = sheet.getLastRow();
  const targetIdentity = makeKeyIdentity_(keyType, keyNumber);

  const emptyRecord = {
    rowIndex: null, keyNumber: keyNumber, keyType: keyType || 'Cowo',
    status: 'Kosong', customerName: '', checkInTime: '', updatedAt: ''
  };

  if (lastRow < 2) return emptyRecord;

  const values = sheet.getRange(2, 1, lastRow - 1, KEY_HEADERS.length).getValues();

  // Pass 1: match exact keyType + keyNumber
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const rowKey = normalizeKeyNumber_(row[0]);
    const rowType = normalizeKeyType_(cleanText_(row[5]));
    if (makeKeyIdentity_(rowType, rowKey) === targetIdentity) {
      return {
        rowIndex: i + 2, keyNumber: rowKey, keyType: rowType,
        status: cleanText_(row[1]) || 'Kosong', customerName: cleanText_(row[2]),
        checkInTime: stringifyCell_(row[3]), updatedAt: stringifyCell_(row[4])
      };
    }
  }

  // Pass 2: fallback data lama (kolom jenis kunci kosong) -> match by keyNumber
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const rowKey = normalizeKeyNumber_(row[0]);
    const rowTypeRaw = cleanText_(row[5]);

    if (rowKey === keyNumber && !rowTypeRaw) {
      const rowType = keyType || 'Cowo';
      sheet.getRange(i + 2, 6).setValue(rowType);
      styleKeyTypeCell_(sheet, i + 2, 6, rowType);

      return {
        rowIndex: i + 2, keyNumber: rowKey, keyType: rowType,
        status: cleanText_(row[1]) || 'Kosong', customerName: cleanText_(row[2]),
        checkInTime: stringifyCell_(row[3]), updatedAt: stringifyCell_(row[4])
      };
    }
  }

  return emptyRecord;
}

function getKeys_() {
  const ss = getSpreadsheet_();
  const sheet = getOrCreateSheet_(ss, SHEET_KEYS);

  // CATATAN: endpoint baca TIDAK lagi memanggil seedKeys_ (yang menulis +
  // styling + bisa kena merge). Seed cukup dijalankan sekali via setupGymSheets.
  // Kalau header belum ada, baru pasang header (operasi ringan).
  if (sheet.getLastRow() === 0) {
    setupHeader_(sheet, KEY_HEADERS);
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const values = sheet.getRange(2, 1, lastRow - 1, KEY_HEADERS.length).getValues();

  return values
    .filter(function (row) { return cleanText_(row[0]); })
    .map(function (row) {
      return {
        keyNumber: normalizeKeyNumber_(row[0]),
        status: cleanText_(row[1]) || 'Kosong',
        customerName: cleanText_(row[2]),
        checkInTime: stringifyCell_(row[3]),
        updatedAt: stringifyCell_(row[4]),
        keyType: normalizeKeyType_(row[5])
      };
    });
}

/* ===============================
   MEMBER_LIFETIME
   =============================== */

function getMembers_() {
  const ss = getSpreadsheet_();
  const sheet = getOrCreateSheet_(ss, SHEET_MEMBERS);

  setupHeader_(sheet, MEMBER_HEADERS);

  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), MEMBER_HEADERS.length);
  if (lastRow < 2) return [];

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(cleanHeader_);
  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

  const idx = {
    memberId: findHeaderIndex_(headers, ['no', 'id member', 'id', 'no member', 'nomor member', 'kode member']),
    memberName: findHeaderIndex_(headers, ['nama member', 'nama', 'name', 'member']),
    status: findHeaderIndex_(headers, ['status', 'status member', 'tipe member']),
    registeredAt: findHeaderIndex_(headers, ['tanggal daftar', 'tanggal', 'join date', 'mulai member', 'tanggal mulai']),
    createdBy: findHeaderIndex_(headers, ['diinput oleh', 'admin', 'input oleh', 'pegawai']),
    updatedAt: findHeaderIndex_(headers, ['update terakhir', 'updated at', 'last update', 'terakhir update'])
  };

  return values
    .map(function (row, index) {
      const fallbackNo = String(index + 1).padStart(3, '0');
      return {
        memberId: getRowValue_(row, idx.memberId) || fallbackNo,
        memberName: getRowValue_(row, idx.memberName),
        status: getRowValue_(row, idx.status) || 'Lifetime',
        registeredAt: stringifyCell_(getRawRowValue_(row, idx.registeredAt)),
        createdBy: getRowValue_(row, idx.createdBy),
        updatedAt: stringifyCell_(getRawRowValue_(row, idx.updatedAt))
      };
    })
    .filter(function (item) { return item.memberName || item.memberId; });
}

/* ===============================
   GET LOGS & DAILY
   =============================== */

function getLogs_() {
  const ss = getSpreadsheet_();
  const sheet = getOrCreateSheet_(ss, SHEET_LOG);

  setupHeader_(sheet, LOG_HEADERS);

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const numberOfRows = Math.min(lastRow - 1, 50);
  const startRow = Math.max(2, lastRow - numberOfRows + 1);
  const values = sheet.getRange(startRow, 1, numberOfRows, LOG_HEADERS.length).getValues();

  return values
    .map(function (row) {
      return {
        no: row[0],
        waktuLengkap: stringifyCell_(row[1]),
        tanggal: stringifyCell_(row[2]),
        jam: stringifyCell_(row[3]),
        nama: cleanText_(row[4]),
        noKunci: normalizeKeyNumber_(row[5]) || cleanText_(row[5]),
        status: cleanText_(row[6]),
        admin: cleanText_(row[7]),
        keyType: normalizeKeyType_(row[8])
      };
    })
    .filter(function (item) { return item.nama || item.noKunci || item.status; })
    .reverse();
}

function getDailyRecap_() {
  const ss = getSpreadsheet_();
  const sheet = getOrCreateSheet_(ss, SHEET_DAILY);

  // Endpoint baca: TIDAK menulis header/checkbox (operasi yang bisa kena merge).
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  const lastCol = Math.max(sheet.getLastColumn(), DAILY_HEADERS.length);
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(cleanHeader_);

  // Baca berdasarkan NAMA HEADER, bukan posisi kolom — tahan kalau struktur geser.
  const idx = {
    tanggal: findHeaderIndex_(headers, ['tanggal']),
    no: findHeaderIndex_(headers, ['no', 'nomor']),
    nama: findHeaderIndex_(headers, ['nama', 'nama member']),
    noKunci: findHeaderIndex_(headers, ['no kunci', 'nomor kunci']),
    jamMasuk: findHeaderIndex_(headers, ['jam masuk']),
    adminMasuk: findHeaderIndex_(headers, ['admin masuk']),
    jamKeluar: findHeaderIndex_(headers, ['jam keluar']),
    sudahKeluar: findHeaderIndex_(headers, ['sudah keluar']),
    waktuKeluarLengkap: findHeaderIndex_(headers, ['waktu keluar lengkap']),
    adminKeluar: findHeaderIndex_(headers, ['admin keluar']),
    keyType: findHeaderIndex_(headers, ['jenis kunci', 'gender', 'tipe kunci'])
  };

  const numberOfRows = Math.min(lastRow - 1, 200);
  const startRow = Math.max(2, lastRow - numberOfRows + 1);
  const values = sheet.getRange(startRow, 1, numberOfRows, lastCol).getValues();

  return values
    .map(function (row) {
      return {
        tanggal: stringifyDateOnly_(getRawRowValue_(row, idx.tanggal)),
        no: getRawRowValue_(row, idx.no),
        nama: cleanText_(getRawRowValue_(row, idx.nama)),
        noKunci: normalizeKeyNumber_(getRawRowValue_(row, idx.noKunci)) || cleanText_(getRawRowValue_(row, idx.noKunci)),
        jamMasuk: stringifyTimeOnly_(getRawRowValue_(row, idx.jamMasuk)),
        adminMasuk: cleanText_(getRawRowValue_(row, idx.adminMasuk)),
        jamKeluar: stringifyTimeOnly_(getRawRowValue_(row, idx.jamKeluar)),
        sudahKeluar: isChecked_(getRawRowValue_(row, idx.sudahKeluar)),
        waktuKeluarLengkap: stringifyDateTimeSafe_(getRawRowValue_(row, idx.waktuKeluarLengkap)),
        adminKeluar: cleanText_(getRawRowValue_(row, idx.adminKeluar)),
        keyType: normalizeKeyType_(getRawRowValue_(row, idx.keyType))
      };
    })
    .filter(function (item) { return item.nama || item.noKunci; })
    .reverse();
}

/* ===============================
   REKAP BULANAN LENGKAP
   =============================== */

function setupRekapBulananDariHarian() {
  TARGET_MONTH_INDEXES.forEach(function (monthIndex) {
    const sheet = getOrCreateMonthlySheet_(2026, monthIndex);
    setupMonthlySheetFull_(sheet, 2026, monthIndex, true);
  });

  hideUnusedMonthlySheets_();
  reorderRekapSheets_();

  return {
    success: true,
    message: 'Setup sheet bulanan lengkap selesai dengan kolom Jenis Kunci.'
  };
}

// resetRekapBulanan: HAPUS TOTAL sheet bulanan lalu bangun ulang dari nol.
// Pakai ini sekali kalau sheet bulanan kadung berantakan / ada merge nyangkut.
// Jalankan manual dari editor. REKAP_HARIAN (data harian) TIDAK disentuh.
function resetRekapBulanan() {
  const ss = getSpreadsheet_();

  TARGET_MONTH_INDEXES.forEach(function (monthIndex) {
    const name = getMonthlySheetName_(2026, monthIndex);
    const existing = ss.getSheetByName(name);
    if (existing) {
      try { ss.deleteSheet(existing); } catch (error) {
        Logger.log('Gagal hapus ' + name + ': ' + error.message);
      }
    }
  });

  // Bangun ulang dari kosong, lalu isi ulang dari REKAP_HARIAN.
  const setup = setupRekapBulananDariHarian();
  const sync = syncSemuaRekapHarianKeBulanan();

  return 'Reset selesai. ' + setup.message + ' | ' + sync.message;
}

function syncSemuaRekapHarianKeBulanan() {
  const ss = getSpreadsheet_();
  const master = ss.getSheetByName(SHEET_DAILY);

  if (!master) throw new Error('Sheet REKAP_HARIAN tidak ditemukan.');

  TARGET_MONTH_INDEXES.forEach(function (monthIndex) {
    const sheet = getOrCreateMonthlySheet_(2026, monthIndex);
    setupMonthlySheetFull_(sheet, 2026, monthIndex, true);
  });

  const lastRow = master.getLastRow();
  const lastCol = master.getLastColumn();

  if (lastRow <= 1) {
    reorderRekapSheets_();
    return { success: false, message: 'REKAP_HARIAN masih kosong.' };
  }

  const sourceHeaders = master.getRange(1, 1, 1, lastCol).getValues()[0].map(cleanHeader_);

  const idx = {
    tanggal: findHeaderIndex_(sourceHeaders, ['tanggal']),
    no: findHeaderIndex_(sourceHeaders, ['no', 'nomor']),
    nama: findHeaderIndex_(sourceHeaders, ['nama', 'nama member']),
    noKunci: findHeaderIndex_(sourceHeaders, ['no kunci', 'nomor kunci']),
    jamMasuk: findHeaderIndex_(sourceHeaders, ['jam masuk']),
    adminMasuk: findHeaderIndex_(sourceHeaders, ['admin masuk']),
    jamKeluar: findHeaderIndex_(sourceHeaders, ['jam keluar']),
    sudahKeluar: findHeaderIndex_(sourceHeaders, ['sudah keluar']),
    waktuKeluarLengkap: findHeaderIndex_(sourceHeaders, ['waktu keluar lengkap']),
    adminKeluar: findHeaderIndex_(sourceHeaders, ['admin keluar']),
    keyType: findHeaderIndex_(sourceHeaders, ['jenis kunci', 'gender', 'tipe kunci'])
  };

  if (idx.tanggal === -1 || idx.nama === -1) {
    throw new Error('Header Tanggal atau Nama tidak ditemukan di REKAP_HARIAN.');
  }

  const values = master.getRange(2, 1, lastRow - 1, lastCol).getValues();
  let totalRows = 0;

  values.forEach(function (row) {
    const tanggal = parseTanggalRekap_(getRawRowValue_(row, idx.tanggal));
    const nama = cleanText_(getRawRowValue_(row, idx.nama));

    if (!tanggal || !nama) return;
    if (tanggal.getFullYear() !== 2026) return;

    const monthIndex = tanggal.getMonth();
    if (TARGET_MONTH_INDEXES.indexOf(monthIndex) === -1) return;

    const sheet = getOrCreateMonthlySheet_(2026, monthIndex);
    const startCol = getStartColByDate_(tanggal);
    const nextRow = getNextRowInMonthlyBlock_(sheet, startCol);

    const normalizedRow = [
      formatDateOnlySafe_(getRawRowValue_(row, idx.tanggal)),
      getRawRowValue_(row, idx.no),
      nama,
      getRawRowValue_(row, idx.noKunci),
      '',
      stringifyTimeOnly_(getRawRowValue_(row, idx.jamMasuk)),
      getRawRowValue_(row, idx.adminMasuk),
      stringifyTimeOnly_(getRawRowValue_(row, idx.jamKeluar)),
      isChecked_(getRawRowValue_(row, idx.sudahKeluar)),
      stringifyDateTimeSafe_(getRawRowValue_(row, idx.waktuKeluarLengkap)),
      getRawRowValue_(row, idx.adminKeluar),
      normalizeKeyType_(getRawRowValue_(row, idx.keyType))
    ];

    writeOneRowToMonthlyBlock_(sheet, startCol, nextRow, normalizedRow);
    totalRows++;
  });

  hideUnusedMonthlySheets_();
  reorderRekapSheets_();

  return { success: true, message: 'Sync selesai. Total baris tersinkron: ' + totalRows };
}

function setupMonthlySheetFull_(sheet, year, monthIndex, clearFirst) {
  if (clearFirst) clearSheetFull_(sheet);

  const totalDays = new Date(year, monthIndex + 1, 0).getDate();
  const neededColumns = totalDays * (MONTHLY_BLOCK_WIDTH + MONTHLY_GAP_WIDTH);

  ensureColumns_(sheet, neededColumns);

  for (let day = 1; day <= totalDays; day++) {
    const date = new Date(year, monthIndex, day);
    const startCol = getStartColByDate_(date);
    setupDateBlock_(sheet, date, startCol);
  }

  sheet.setFrozenRows(2);
}

function clearSheetFull_(sheet) {
  // Break SEMUA merge dulu, dengan beberapa lapis pengaman, supaya
  // setupDateBlock_ tidak menabrak merge lama (penyebab error
  // "Anda harus memilih semua sel dalam rentang penggabungan").
  breakAllMerges_(sheet);

  // Hapus isi + format. clearContents lalu clear penuh.
  try { sheet.clear(); } catch (error) {
    try { sheet.clearContents(); } catch (e) {}
    try { sheet.clearFormats(); } catch (e) {}
  }

  // Pastikan tidak ada checkbox / data validation sisa.
  try {
    sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
      .clearDataValidations();
  } catch (error) {}
}

// breakAllMerges_: lepas semua merged range di sheet, satu per satu.
// Lebih andal daripada breakApart() di satu range besar.
function breakAllMerges_(sheet) {
  try {
    const merged = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
      .getMergedRanges();

    merged.forEach(function (range) {
      try { range.breakApart(); } catch (error) {}
    });
  } catch (error) {
    // Fallback terakhir.
    try {
      sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).breakApart();
    } catch (e) {}
  }
}

// safeBreakMerges_: versi ringan & aman untuk dipanggil di jalur kritis.
// Tidak pernah melempar error.
function safeBreakMerges_(sheet) {
  try {
    const merges = sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns())
      .getMergedRanges();
    merges.forEach(function (r) {
      try { r.breakApart(); } catch (e) {}
    });
  } catch (error) {
    Logger.log('safeBreakMerges_ dilewati (' + (sheet ? sheet.getName() : '?') + '): ' + error.message);
  }
}

function setupDateBlock_(sheet, date, startCol) {
  const title = formatTanggalPanjang_(date);
  const titleRange = sheet.getRange(1, startCol, 1, MONTHLY_BLOCK_WIDTH);
  const headerRange = sheet.getRange(2, startCol, 1, MONTHLY_BLOCK_WIDTH);

  // Break merge lama yang mungkin lebih lebar dari blok baru: cek baris 1
  // sepanjang satu blok + gap, lepas semua merge yang ketemu.
  try {
    const safeWidth = MONTHLY_BLOCK_WIDTH + MONTHLY_GAP_WIDTH;
    const maxCol = sheet.getMaxColumns();
    const width = Math.min(safeWidth, maxCol - startCol + 1);
    sheet.getRange(1, startCol, 1, width).getMergedRanges()
      .forEach(function (r) { try { r.breakApart(); } catch (e) {} });
  } catch (error) {}

  try { titleRange.breakApart(); } catch (error) {}

  titleRange.merge();
  titleRange.setValue(title);
  titleRange.setFontWeight('bold').setHorizontalAlignment('center').setBackground('#d9ead3');

  headerRange.setValues([DAILY_HEADERS]);
  headerRange.setFontWeight('bold').setHorizontalAlignment('center').setBackground('#dbeafe');

  sheet.getRange(1, startCol, MONTHLY_VISIBLE_ROWS, MONTHLY_BLOCK_WIDTH)
    .setBorder(true, true, true, true, true, true);

  const gapCol = startCol + MONTHLY_BLOCK_WIDTH;
  sheet.getRange(1, gapCol, MONTHLY_VISIBLE_ROWS, 1).setBackground('#f3f3f3');

  setMonthlyBlockColumnWidths_(sheet, startCol);
}

function writeOneRowToMonthlyBlock_(sheet, startCol, rowIndex, rowValues) {
  const range = sheet.getRange(rowIndex, startCol, 1, MONTHLY_BLOCK_WIDTH);
  range.setNumberFormat('@');
  range.setValues([rowValues]);

  const checkboxCol = startCol + 8;
  const checkboxCell = sheet.getRange(rowIndex, checkboxCol);
  checkboxCell.insertCheckboxes();
  checkboxCell.setValue(rowValues[8] === true);

  styleKeyTypeCell_(sheet, rowIndex, startCol + 11, rowValues[11]);
}

function getStartColByDate_(date) {
  const day = new Date(date).getDate();
  return 1 + (day - 1) * (MONTHLY_BLOCK_WIDTH + MONTHLY_GAP_WIDTH);
}

function getNextRowInMonthlyBlock_(sheet, startCol) {
  const namaCol = startCol + 2;
  const maxRows = Math.max(sheet.getLastRow(), MONTHLY_VISIBLE_ROWS);

  const values = sheet.getRange(MONTHLY_START_ROW, namaCol, maxRows - MONTHLY_START_ROW + 1, 1).getValues();

  for (let i = 0; i < values.length; i++) {
    if (!values[i][0]) return MONTHLY_START_ROW + i;
  }

  return maxRows + 1;
}

function setMonthlyBlockColumnWidths_(sheet, startCol) {
  const widths = [110, 55, 220, 90, 35, 120, 120, 120, 120, 180, 120, 110];
  widths.forEach(function (width, index) {
    sheet.setColumnWidth(startCol + index, width);
  });
}

function hideUnusedMonthlySheets_() {
  const ss = getSpreadsheet_();
  const allowedNames = TARGET_MONTH_INDEXES.map(function (monthIndex) {
    return getMonthlySheetName_(2026, monthIndex);
  });

  ss.getSheets().forEach(function (sheet) {
    const name = sheet.getName();
    if (/^REKAP_[A-Z]+_2026$/.test(name) && allowedNames.indexOf(name) === -1) {
      sheet.hideSheet();
    }
    if (allowedNames.indexOf(name) !== -1) sheet.showSheet();
  });
}

function reorderRekapSheets_() {
  const ss = getSpreadsheet_();
  const ordered = [SHEET_DAILY].concat(
    TARGET_MONTH_INDEXES.map(function (monthIndex) {
      return getMonthlySheetName_(2026, monthIndex);
    })
  );

  let position = 1;
  ordered.forEach(function (sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    if (sheet) {
      sheet.showSheet();
      ss.setActiveSheet(sheet);
      ss.moveActiveSheet(position);
      position++;
    }
  });
}

function getOrCreateMonthlySheet_(year, monthIndex) {
  const ss = getSpreadsheet_();
  const sheetName = getMonthlySheetName_(year, monthIndex);
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) sheet = ss.insertSheet(sheetName);
  return sheet;
}

function getMonthlySheetName_(year, monthIndex) {
  return 'REKAP_' + MONTH_ID[monthIndex] + '_' + year;
}

function ensureColumns_(sheet, neededCol) {
  const currentCols = sheet.getMaxColumns();
  if (currentCols < neededCol) {
    sheet.insertColumnsAfter(currentCols, neededCol - currentCols);
  }
}

function parseTanggalRekap_(value) {
  if (!value) return null;

  if (Object.prototype.toString.call(value) === '[object Date]') {
    if (isNaN(value.getTime())) return null;
    const date = new Date(value);
    date.setHours(0, 0, 0, 0);
    return date;
  }

  const text = String(value).trim();

  const slashMatch = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (slashMatch) {
    const date = new Date(Number(slashMatch[3]), Number(slashMatch[2]) - 1, Number(slashMatch[1]));
    date.setHours(0, 0, 0, 0);
    return date;
  }

  const isoMatch = text.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (isoMatch) {
    const date = new Date(Number(isoMatch[1]), Number(isoMatch[2]) - 1, Number(isoMatch[3]));
    date.setHours(0, 0, 0, 0);
    return date;
  }

  const parsed = new Date(text);
  if (isNaN(parsed.getTime())) return null;
  parsed.setHours(0, 0, 0, 0);
  return parsed;
}

function formatTanggalPanjang_(date) {
  const d = new Date(date);
  return d.getDate() + ' ' + MONTH_DISPLAY[d.getMonth()] + ' ' + d.getFullYear();
}

function formatDateOnlySafe_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, TIMEZONE, 'dd/MM/yyyy');
  }
  const parsed = parseTanggalRekap_(value);
  if (parsed) return Utilities.formatDate(parsed, TIMEZONE, 'dd/MM/yyyy');
  return cleanText_(value);
}

function stringifyTimeOnly_(value) {
  if (!value) return '';

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, TIMEZONE, 'HH:mm:ss');
  }

  if (typeof value === 'number') {
    const totalSeconds = Math.round(value * 24 * 60 * 60);
    const hh = String(Math.floor(totalSeconds / 3600) % 24).padStart(2, '0');
    const mm = String(Math.floor((totalSeconds % 3600) / 60)).padStart(2, '0');
    const ss = String(totalSeconds % 60).padStart(2, '0');
    return hh + ':' + mm + ':' + ss;
  }

  const text = cleanText_(value);
  const match = text.match(/(\d{1,2}:\d{2}(?::\d{2})?)/);
  if (match) {
    const parts = match[1].split(':');
    return String(parts[0]).padStart(2, '0') + ':' +
           String(parts[1] || '00').padStart(2, '0') + ':' +
           String(parts[2] || '00').padStart(2, '0');
  }

  return text;
}

function stringifyDateTimeSafe_(value) {
  if (!value) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, TIMEZONE, 'dd/MM/yyyy HH:mm:ss');
  }
  const text = cleanText_(value);
  const match = text.match(/(\d{1,2}\/\d{1,2}\/\d{4})\s+(\d{1,2}:\d{2}(?::\d{2})?)/);
  if (match) return match[1] + ' ' + stringifyTimeOnly_(match[2]);
  return text;
}

/* ===============================
   HELPER WARNA JENIS KUNCI
   Cowo = HITAM (#111827), Cewe = MERAH (#dc2626)
   =============================== */

function styleKeyTypeCell_(sheet, rowIndex, colIndex, keyType) {
  const normalized = normalizeKeyType_(keyType);
  const color = normalized === 'Cewe' ? COLOR_CEWE : COLOR_COWO;
  sheet.getRange(rowIndex, colIndex).setFontColor(color).setFontWeight('bold');
}

function styleKeyTypeColumn_(sheet, colIndex) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;

  const range = sheet.getRange(2, colIndex, lastRow - 1, 1);
  const values = range.getValues();

  const colors = values.map(function (row) {
    return [normalizeKeyType_(row[0]) === 'Cewe' ? COLOR_CEWE : COLOR_COWO];
  });
  const weights = values.map(function () { return ['bold']; });

  range.setFontColors(colors);
  range.setFontWeights(weights);
}

/* ===============================
   GENERAL UTILITIES
   =============================== */

function cleanHeader_(value) {
  return String(value || '').trim().toLowerCase().replace(/\s+/g, ' ');
}

function findHeaderIndex_(headers, aliases) {
  for (let i = 0; i < aliases.length; i++) {
    const index = headers.indexOf(aliases[i]);
    if (index !== -1) return index;
  }
  return -1;
}

function getRawRowValue_(row, index) {
  return index >= 0 ? row[index] : '';
}

function getRowValue_(row, index) {
  return cleanText_(getRawRowValue_(row, index));
}

function stringifyCell_(value) {
  if (value instanceof Date) return formatDateTime_(value);
  return cleanText_(value);
}

function stringifyDateOnly_(value) {
  if (value instanceof Date) return formatDate_(value);
  return cleanText_(value);
}

function extractTimeFromDateTimeText_(value) {
  const text = cleanText_(value);
  if (!text) return '';
  const match = text.match(/(\d{2}:\d{2}:\d{2})/);
  if (match) return match[1];
  return text;
}

function formatDate_(date) { return Utilities.formatDate(date, TIMEZONE, 'dd/MM/yyyy'); }
function formatTime_(date) { return Utilities.formatDate(date, TIMEZONE, 'HH:mm:ss'); }
function formatDateTime_(date) { return Utilities.formatDate(date, TIMEZONE, 'dd/MM/yyyy HH:mm:ss'); }

function autoResize_(sheet, length) {
  for (let i = 1; i <= length; i++) sheet.autoResizeColumn(i);
}

// safeRead_: jalankan fungsi baca, kalau gagal balikan data kosong dengan
// ok:true supaya auto-refresh web TIDAK memunculkan toast merah. Error dicatat
// di Logger untuk diagnosa, tapi tidak mengganggu pengguna.
function safeRead_(label, fn) {
  try {
    return { ok: true, message: label + ' berhasil diambil.', data: fn() };
  } catch (error) {
    Logger.log('safeRead_ ' + label + ' gagal: ' + (error.message || error));
    return { ok: true, message: label + ' kosong.', data: [] };
  }
}

function respondJson_(callback, payload) {
  const json = JSON.stringify(payload);

  if (callback) {
    const safeCallback = String(callback).replace(/[^a-zA-Z0-9_.$]/g, '');
    return ContentService.createTextOutput(safeCallback + '(' + json + ');')
      .setMimeType(ContentService.MimeType.JAVASCRIPT);
  }

  return ContentService.createTextOutput(json).setMimeType(ContentService.MimeType.JSON);
}

function respondPostMessage_(payload) {
  const safeJson = JSON.stringify(payload).replace(/</g, '\\u003c');
  const html = '<!doctype html><html><body><script>' +
    'window.parent.postMessage({source:"sistem-gym-backend",payload:' + safeJson + '},"*");' +
    '</script></body></html>';

  return HtmlService.createHtmlOutput(html)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
