/**
 * ASAS SMP NUGRAHA — Backend (Google Apps Script Web App)
 * ---------------------------------------------------------------
 * Deploy: Extensions > Apps Script > Deploy > New deployment > Web app
 *   Execute as : Me (pemilik spreadsheet)
 *   Who has access : Anyone
 * Salin URL "/exec" ke CONFIG.API_URL di index.html.
 *
 * Sheet yang dipakai:
 *   "Siswa"            : Nama Sekolah | Nama siswa | Nomor peserta | Kelas | Password
 *   "Guru / Admin"     : Nama Sekolah | Nama Guru | Username / NIP | Password
 *   "Pengaturan Ujian" : ID Ujian | Nama Sekolah | Nama Ujian | Link Google Form | Batas Pelanggaran | Durasi (Menit) | Status
 *   "Sesi ASAS"         : (dibuat otomatis) status tiap siswa, jumlah pelanggaran, dll.
 *   "Log ASAS"          : (dibuat otomatis) riwayat setiap pelanggaran.
 *
 * Catatan: untuk admin "super" (mengelola semua sekolah), isi kolom
 * Nama Sekolah di sheet Guru / Admin dengan: SEMUA
 */

const CONFIG = {
  SPREADSHEET_ID: '1H93hxkL7yq2ncJ1Tbratt59osnvfKGDY', // harus berupa Google Sheets (bukan file .xlsx)
  SHEET_SISWA: 'Siswa',
  SHEET_GURU: 'Guru / Admin',
  SHEET_UJIAN: 'Pengaturan Ujian',
  // Nama sengaja berbeda dari sheet lama "Sesi Ujian"/"Log Pelanggaran" (struktur kolomnya lain)
  SHEET_SESI: 'Sesi ASAS',
  SHEET_LOG: 'Log ASAS',
  TOKEN_TTL_SEC: 6 * 60 * 60,   // maksimum CacheService = 21600 detik (6 jam)
  ACTIVE_WINDOW_SEC: 60,        // siswa dianggap "terputus" jika tidak ada heartbeat selama ini
  RELOAD_PENALTY: true,         // memuat ulang / menutup halaman saat ujian = 1 pelanggaran
  SUPER_SCHOOL_KEYS: ['semua', 'all', '*'],
  MAX_LOGIN_FAIL: 5             // gagal login beruntun sebelum diblok 10 menit
};

const ST = { WORK: 'Sedang Mengerjakan', LOCK: 'Terkunci', DONE: 'Selesai' };

const SESI_HEADERS = ['ID Sesi', 'ID Ujian', 'Nama Sekolah', 'Nomor Peserta', 'Nama Siswa', 'Kelas',
  'Mulai', 'Terakhir Aktif', 'Pelanggaran', 'Status', 'Keterangan', 'Selesai'];
const LOG_HEADERS = ['Waktu', 'ID Sesi', 'Nomor Peserta', 'Nama Siswa', 'Nama Sekolah', 'ID Ujian',
  'Jenis Pelanggaran', 'Pelanggaran Ke', 'Status Setelahnya'];

// Alias nama kolom (sudah dinormalisasi: huruf kecil, tanpa spasi/simbol)
const K = {
  sekolah: ['namasekolah', 'sekolah'],
  namaSiswa: ['namasiswa', 'nama'],
  namaGuru: ['namaguru', 'nama'],
  nomor: ['nomorpeserta', 'nopeserta', 'nomor'],
  kelas: ['kelas'],
  password: ['password', 'katasandi', 'sandi'],
  username: ['usernamenip', 'username', 'nip'],
  idUjian: ['idujian', 'id'],
  namaUjian: ['namaujian', 'ujian'],
  link: ['linkgoogleform', 'linkform', 'link'],
  batas: ['bataspelanggaran', 'batasanpelanggaran', 'batas', 'batasan'],
  durasi: ['durasimenit', 'durasi', 'waktumenit'],
  status: ['status', 'statusaktifnonaktif', 'statusaktif']
};

/* ===================== ENTRY POINTS ===================== */

function doGet(e) {
  try {
    const a = (e && e.parameter && e.parameter.action) || 'ping';
    if (a === 'ping') return json_({ ok: true, time: new Date().toISOString() });
    if (a === 'schools') return json_({ ok: true, schools: listSchools_() });
    return json_({ ok: false, error: 'Aksi tidak dikenal.' });
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function doPost(e) {
  try {
    const body = JSON.parse((e.postData && e.postData.contents) || '{}');
    const fn = ACTIONS[body.action];
    if (!fn) throw new Error('Aksi tidak dikenal: ' + body.action);
    const res = fn(body) || {};
    res.ok = true;
    res.serverNow = Date.now();
    return json_(res);
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err), code: err.code || '' });
  }
}

const ACTIONS = {
  loginSiswa: loginSiswa_,
  loginAdmin: loginAdmin_,
  studentExams: studentExams_action_,
  startSession: startSession_,
  heartbeat: heartbeat_,
  logViolation: logViolation_,
  finishSession: finishSession_,
  adminData: adminData_,
  saveExam: saveExam_,
  deleteExam: deleteExam_,
  monitoring: monitoring_,
  unlock: unlock_,
  resetSesi: resetSesi_,
  unlockProktor: unlockProktor_
};

/** Jalankan SEKALI dari editor untuk membuat sheet bantu & memberi izin akses. */
function setupSheets() {
  sheet_(CONFIG.SHEET_SESI, SESI_HEADERS, [1, 4]);
  sheet_(CONFIG.SHEET_LOG, LOG_HEADERS, [2, 3]);
  Logger.log('Sheet "Sesi Ujian" dan "Log Pelanggaran" siap.');
}

/* ===================== UTIL ===================== */

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
function norm_(s) { return String(s == null ? '' : s).toLowerCase().replace(/[^a-z0-9]/g, ''); }
function str_(v) { return String(v == null ? '' : v).trim(); }
function toMs_(v) { return v instanceof Date ? v.getTime() : (v ? new Date(v).getTime() || 0 : 0); }
function get_(row, keys) {
  for (let i = 0; i < keys.length; i++) {
    const v = row[keys[i]];
    if (v !== undefined && v !== '') return v;
  }
  return '';
}
function authErr_(msg) {
  const e = new Error(msg || 'Sesi login berakhir. Silakan login ulang.');
  e.code = 'AUTH';
  return e;
}
function lock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(25000);
  try { return fn(); } finally { lock.releaseLock(); }
}

/** Ambil sheet berdasarkan nama (toleran spasi/simbol). Buat otomatis jika headers diberikan. */
function sheet_(name, headers, textCols) {
  const ss = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
  const target = norm_(name);
  const found = ss.getSheets().filter(function (s) { return norm_(s.getName()) === target; })[0];
  if (found) return found;
  if (!headers) throw new Error('Sheet "' + name + '" tidak ditemukan di spreadsheet.');
  const sh = ss.insertSheet(name);
  sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#E0E7FF');
  sh.setFrozenRows(1);
  (textCols || []).forEach(function (c) { sh.getRange(1, c, sh.getMaxRows(), 1).setNumberFormat('@'); });
  return sh;
}

/** Baca tabel dengan header baris 1 -> array objek berkunci header ternormalisasi. */
function readTable_(sh) {
  const lastRow = sh.getLastRow(), lastCol = sh.getLastColumn();
  if (lastCol < 1) return { headers: [], rows: [] };
  const values = sh.getRange(1, 1, Math.max(lastRow, 1), lastCol).getValues();
  const headers = values[0].map(norm_);
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const r = values[i];
    if (r.every(function (v) { return String(v).trim() === ''; })) continue;
    const o = { _row: i + 1 };
    headers.forEach(function (h, j) { if (h && o[h] === undefined) o[h] = r[j]; });
    rows.push(o);
  }
  return { headers: headers, rows: rows };
}

function put_(arr, headers, keys, value) {
  for (let i = 0; i < keys.length; i++) {
    const idx = headers.indexOf(keys[i]);
    if (idx >= 0) { arr[idx] = value; return; }
  }
  throw new Error('Kolom "' + keys[0] + '" tidak ada di sheet Pengaturan Ujian.');
}

/* --- token & anti brute force --- */
function issueToken_(payload) {
  const t = Utilities.getUuid();
  CacheService.getScriptCache().put('tok_' + t, JSON.stringify(payload), CONFIG.TOKEN_TTL_SEC);
  return t;
}
function auth_(token, role) {
  const raw = token ? CacheService.getScriptCache().get('tok_' + token) : null;
  if (!raw) throw authErr_();
  const p = JSON.parse(raw);
  if (role && p.role !== role) throw authErr_('Akses ditolak.');
  return p;
}
function checkBrute_(key) {
  if (Number(CacheService.getScriptCache().get('f_' + key) || 0) >= CONFIG.MAX_LOGIN_FAIL)
    throw new Error('Terlalu banyak percobaan gagal. Coba lagi dalam 10 menit.');
}
function addFail_(key) {
  const c = CacheService.getScriptCache();
  c.put('f_' + key, String(Number(c.get('f_' + key) || 0) + 1), 600);
}
function clearFail_(key) { CacheService.getScriptCache().remove('f_' + key); }

function inScope_(p, sekolah) { return p.isSuper || norm_(p.sekolah) === norm_(sekolah); }
function isSuperKey_(s) { return CONFIG.SUPER_SCHOOL_KEYS.indexOf(norm_(s)) >= 0; }

/* ===================== DATA: UJIAN ===================== */

function readExams_() {
  const t = readTable_(sheet_(CONFIG.SHEET_UJIAN));
  return t.rows.map(function (r) {
    return {
      row: r._row,
      id: str_(get_(r, K.idUjian)),
      sekolah: str_(get_(r, K.sekolah)),
      nama: str_(get_(r, K.namaUjian)),
      link: str_(get_(r, K.link)),
      batas: Number(get_(r, K.batas)) || 3,
      durasi: Number(get_(r, K.durasi)) || 60,
      status: str_(get_(r, K.status)),
      aktif: norm_(get_(r, K.status)) === 'aktif'
    };
  }).filter(function (e) { return e.id; });
}

/** Ambil 1 ujian; cache 20 detik untuk operasi sering (heartbeat/pelanggaran). */
function examById_(id, useCache) {
  const c = CacheService.getScriptCache();
  if (useCache) {
    const raw = c.get('ex_' + id);
    if (raw) return JSON.parse(raw);
  }
  const ex = readExams_().filter(function (e) { return e.id === String(id); })[0] || null;
  if (ex) c.put('ex_' + id, JSON.stringify(ex), 20);
  return ex;
}
function invalidateExam_(id) { CacheService.getScriptCache().remove('ex_' + id); }

function listSchools_() {
  const seen = {}, out = [];
  [[CONFIG.SHEET_SISWA, K.sekolah], [CONFIG.SHEET_GURU, K.sekolah], [CONFIG.SHEET_UJIAN, K.sekolah]].forEach(function (x) {
    try {
      readTable_(sheet_(x[0])).rows.forEach(function (r) {
        const s = str_(get_(r, x[1]));
        if (s && !isSuperKey_(s) && !seen[norm_(s)]) { seen[norm_(s)] = 1; out.push(s); }
      });
    } catch (e) { /* sheet belum ada */ }
  });
  return out.sort();
}

/* ===================== DATA: SESI & LOG ===================== */

function sesiSheet_() { return sheet_(CONFIG.SHEET_SESI, SESI_HEADERS, [1, 4]); }
function logSheet_() { return sheet_(CONFIG.SHEET_LOG, LOG_HEADERS, [2, 3]); }

function sesiFromValues_(r, row) {
  return {
    row: row, id: str_(r[0]), idUjian: str_(r[1]), sekolah: str_(r[2]), nomor: str_(r[3]), nama: str_(r[4]),
    kelas: str_(r[5]), mulai: toMs_(r[6]), last: toMs_(r[7]), pelanggaran: Number(r[8]) || 0,
    status: str_(r[9]), ket: str_(r[10]), selesai: toMs_(r[11])
  };
}
function findSesiRow_(sh, id) {
  const last = sh.getLastRow();
  if (last < 2) return 0;
  const f = sh.getRange(2, 1, last - 1, 1).createTextFinder(String(id)).matchEntireCell(true).findNext();
  return f ? f.getRow() : 0;
}
function getSesi_(sh, id) {
  const row = findSesiRow_(sh, id);
  if (!row) return null;
  return sesiFromValues_(sh.getRange(row, 1, 1, SESI_HEADERS.length).getValues()[0], row);
}
function setSesi_(sh, row, patch) {
  const col = { last: 8, pelanggaran: 9, status: 10, ket: 11, selesai: 12 };
  Object.keys(patch).forEach(function (k) { sh.getRange(row, col[k]).setValue(patch[k]); });
}
function appendLog_(s, jenis, ke, statusAfter) {
  logSheet_().appendRow([new Date(), s.id, s.nomor, s.nama, s.sekolah, s.idUjian, jenis, ke, statusAfter]);
}
function sisaDetik_(s, ex, now) {
  return Math.round((ex ? ex.durasi : 60) * 60 - (now - s.mulai) / 1000);
}

/* ===================== AKSI: LOGIN ===================== */

function loginSiswa_(b) {
  const sekolah = str_(b.sekolah), nomor = str_(b.nomor), pass = str_(b.password);
  if (!sekolah || !nomor || !pass) throw new Error('Lengkapi semua data login.');
  const key = 's_' + norm_(sekolah) + '_' + norm_(nomor);
  checkBrute_(key);
  const r = readTable_(sheet_(CONFIG.SHEET_SISWA)).rows.filter(function (x) {
    return norm_(get_(x, K.sekolah)) === norm_(sekolah) && norm_(get_(x, K.nomor)) === norm_(nomor);
  })[0];
  if (!r || str_(get_(r, K.password)) !== pass) {
    addFail_(key);
    throw new Error('Nomor peserta atau password salah.');
  }
  clearFail_(key);
  const siswa = {
    nama: str_(get_(r, K.namaSiswa)), sekolah: str_(get_(r, K.sekolah)),
    nomor: str_(get_(r, K.nomor)), kelas: str_(get_(r, K.kelas))
  };
  const token = issueToken_({ role: 'siswa', sekolah: siswa.sekolah, nomor: siswa.nomor, nama: siswa.nama, kelas: siswa.kelas });
  return { token: token, siswa: siswa, exams: studentExams_(siswa) };
}

function loginAdmin_(b) {
  const u = str_(b.username), pw = str_(b.password);
  if (!u || !pw) throw new Error('Lengkapi username dan password.');
  const key = 'a_' + norm_(u);
  checkBrute_(key);
  const r = readTable_(sheet_(CONFIG.SHEET_GURU)).rows.filter(function (x) {
    return norm_(get_(x, K.username)) === norm_(u);
  })[0];
  if (!r || str_(get_(r, K.password)) !== pw) {
    addFail_(key);
    throw new Error('Username/NIP atau password salah.');
  }
  clearFail_(key);
  const admin = {
    nama: str_(get_(r, K.namaGuru)), username: str_(get_(r, K.username)),
    sekolah: str_(get_(r, K.sekolah)), isSuper: isSuperKey_(get_(r, K.sekolah))
  };
  const token = issueToken_({ role: 'admin', nama: admin.nama, username: admin.username, sekolah: admin.sekolah, isSuper: admin.isSuper });
  return { token: token, admin: admin };
}

/* ===================== AKSI: SISWA ===================== */

/** Daftar ujian aktif untuk sekolah siswa (link TIDAK dikirim sampai ujian dimulai). */
function studentExams_(s) {
  const map = {};
  const sh = sesiSheet_();
  const last = sh.getLastRow();
  if (last >= 2) {
    sh.getRange(2, 1, last - 1, SESI_HEADERS.length).getValues().forEach(function (r, i) {
      const x = sesiFromValues_(r, i + 2);
      if (x.nomor === String(s.nomor)) map[x.idUjian] = x.status;
    });
  }
  return readExams_().filter(function (e) {
    return e.aktif && norm_(e.sekolah) === norm_(s.sekolah);
  }).map(function (e) {
    return { id: e.id, nama: e.nama, sekolah: e.sekolah, batas: e.batas, durasi: e.durasi, sesi: map[e.id] || '' };
  });
}
function studentExams_action_(b) {
  const p = auth_(b.token, 'siswa');
  return { exams: studentExams_(p) };
}

function startSession_(b) {
  const p = auth_(b.token, 'siswa');
  return lock_(function () {
    const ex = examById_(b.idUjian, false); // selalu baca data terbaru dari Sheet
    if (!ex || !ex.aktif || norm_(ex.sekolah) !== norm_(p.sekolah))
      throw new Error('Ujian tidak tersedia atau sudah dinonaktifkan oleh admin.');
    if (!/^https:\/\/(docs\.google\.com\/forms\/|demo)/.test(ex.link) && ex.link !== 'demo')
      throw new Error('Link Google Form belum valid. Hubungi admin.');
    const sh = sesiSheet_();
    const id = p.nomor + '_' + ex.id;
    const now = Date.now();
    const cur = getSesi_(sh, id);
    const base = function (o) {
      return Object.assign({ sesiId: id, link: ex.link, batas: ex.batas, durasi: ex.durasi, namaUjian: ex.nama }, o);
    };

    if (!cur) {
      sh.appendRow([id, ex.id, p.sekolah, p.nomor, p.nama, p.kelas, new Date(now), new Date(now), 0, ST.WORK, '', '']);
      return base({ status: ST.WORK, pelanggaran: 0, sisaDetik: ex.durasi * 60, resumed: false });
    }
    const sisa = sisaDetik_(cur, ex, now);
    if (cur.status === ST.DONE)
      return base({ status: ST.DONE, pelanggaran: cur.pelanggaran, sisaDetik: 0, keterangan: cur.ket });
    if (cur.status === ST.LOCK)
      return base({ status: ST.LOCK, pelanggaran: cur.pelanggaran, sisaDetik: Math.max(0, sisa) });
    if (sisa <= 0) {
      setSesi_(sh, cur.row, { status: ST.DONE, ket: 'Waktu habis', selesai: new Date(now), last: new Date(now) });
      return base({ status: ST.DONE, pelanggaran: cur.pelanggaran, sisaDetik: 0, keterangan: 'Waktu habis' });
    }
    // Melanjutkan sesi yang sedang berjalan = siswa sempat keluar / memuat ulang halaman
    let count = cur.pelanggaran, status = ST.WORK, penalized = false;
    if (CONFIG.RELOAD_PENALTY) {
      count++;
      penalized = true;
      if (count >= ex.batas) status = ST.LOCK;
      appendLog_(cur, 'Memuat ulang / menutup halaman ujian', count, status);
    }
    setSesi_(sh, cur.row, {
      pelanggaran: count, status: status, last: new Date(now),
      ket: status === ST.LOCK ? 'Melebihi batas pelanggaran' : cur.ket
    });
    return base({ status: status, pelanggaran: count, sisaDetik: sisa, resumed: true, penalized: penalized });
  });
}

function heartbeat_(b) {
  const p = auth_(b.token, 'siswa');
  const sh = sesiSheet_();
  const cur = getSesi_(sh, b.sesiId);
  if (!cur || cur.nomor !== p.nomor) throw new Error('Sesi ujian tidak ditemukan.');
  const ex = examById_(cur.idUjian, true);
  const now = Date.now();
  const sisa = sisaDetik_(cur, ex, now);
  if (cur.status === ST.WORK && sisa <= 0) {
    lock_(function () {
      setSesi_(sh, cur.row, { status: ST.DONE, ket: 'Waktu habis', selesai: new Date(now), last: new Date(now) });
    });
    cur.status = ST.DONE;
    cur.ket = 'Waktu habis';
  } else {
    setSesi_(sh, cur.row, { last: new Date(now) });
  }
  return { status: cur.status, pelanggaran: cur.pelanggaran, batas: ex ? ex.batas : 3, sisaDetik: Math.max(0, sisa), keterangan: cur.ket };
}

function logViolation_(b) {
  const p = auth_(b.token, 'siswa');
  return lock_(function () {
    const sh = sesiSheet_();
    const cur = getSesi_(sh, b.sesiId);
    if (!cur || cur.nomor !== p.nomor) throw new Error('Sesi ujian tidak ditemukan.');
    const ex = examById_(cur.idUjian, true);
    const batas = ex ? ex.batas : 3;
    if (cur.status !== ST.WORK) return { status: cur.status, pelanggaran: cur.pelanggaran, batas: batas };
    const n = cur.pelanggaran + 1;
    const st = n >= batas ? ST.LOCK : ST.WORK;
    setSesi_(sh, cur.row, {
      pelanggaran: n, status: st, last: new Date(),
      ket: st === ST.LOCK ? 'Melebihi batas pelanggaran' : cur.ket
    });
    appendLog_(cur, str_(b.jenis).slice(0, 200), n, st);
    return { status: st, pelanggaran: n, batas: batas };
  });
}

function finishSession_(b) {
  const p = auth_(b.token, 'siswa');
  return lock_(function () {
    const sh = sesiSheet_();
    const cur = getSesi_(sh, b.sesiId);
    if (!cur || cur.nomor !== p.nomor) throw new Error('Sesi ujian tidak ditemukan.');
    if (cur.status === ST.WORK) {
      setSesi_(sh, cur.row, { status: ST.DONE, ket: str_(b.keterangan) || 'Dikumpulkan siswa', selesai: new Date(), last: new Date() });
    }
    return { status: cur.status === ST.WORK ? ST.DONE : cur.status };
  });
}

/* ===================== AKSI: ADMIN ===================== */

function adminData_(b) {
  const p = auth_(b.token, 'admin');
  const exams = readExams_().filter(function (e) { return inScope_(p, e.sekolah); });
  return { exams: exams, schools: p.isSuper ? listSchools_() : [p.sekolah] };
}

function saveExam_(b) {
  const p = auth_(b.token, 'admin');
  const x = b.exam || {};
  const sekolah = str_(p.isSuper ? x.sekolah : p.sekolah);
  const nama = str_(x.nama), link = str_(x.link);
  const batas = Math.round(Number(x.batas)), durasi = Math.round(Number(x.durasi));
  const status = norm_(x.status) === 'aktif' ? 'Aktif' : 'Nonaktif';
  if (!sekolah) throw new Error('Nama sekolah wajib diisi.');
  if (!nama) throw new Error('Nama ujian wajib diisi.');
  if (!/^https:\/\/docs\.google\.com\/forms\//.test(link) && link !== 'demo')
    throw new Error('Link harus berupa URL Google Form lengkap (https://docs.google.com/forms/...).');
  if (!(batas >= 1 && batas <= 50)) throw new Error('Batas pelanggaran harus 1–50.');
  if (!(durasi >= 1 && durasi <= 600)) throw new Error('Durasi harus 1–600 menit.');

  return lock_(function () {
    const sh = sheet_(CONFIG.SHEET_UJIAN);
    // Pastikan semua kolom wajib ada; jika tidak ditemukan, header ditambahkan di kolom kosong paling kanan.
    const need = [[K.idUjian, 'ID Ujian'], [K.sekolah, 'Nama Sekolah'], [K.namaUjian, 'Nama Ujian'], [K.link, 'Link Google Form'],
      [K.batas, 'Batas Pelanggaran'], [K.durasi, 'Durasi (Menit)'], [K.status, 'Status']];
    let hdr = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(norm_);
    need.forEach(function (n) {
      const has = n[0].some(function (k) { return hdr.indexOf(k) >= 0; });
      if (!has) {
        sh.getRange(1, hdr.length + 1).setValue(n[1]).setFontWeight('bold');
        hdr.push(norm_(n[1]));
      }
    });
    const lastCol = hdr.length;
    const headers = hdr;
    let id = str_(x.id);
    const existing = id ? readExams_().filter(function (e) { return e.id === id; })[0] : null;
    let rowIdx, arr;
    if (existing) {
      if (!inScope_(p, existing.sekolah)) throw new Error('Anda tidak berhak mengubah ujian sekolah lain.');
      rowIdx = existing.row;
      arr = sh.getRange(rowIdx, 1, 1, lastCol).getValues()[0];
    } else {
      id = id || newExamId_();
      rowIdx = sh.getLastRow() + 1;
      arr = new Array(lastCol).fill('');
    }
    put_(arr, headers, K.idUjian, id);
    put_(arr, headers, K.sekolah, sekolah);
    put_(arr, headers, K.namaUjian, nama);
    put_(arr, headers, K.link, link);
    put_(arr, headers, K.batas, batas);
    put_(arr, headers, K.durasi, durasi);
    put_(arr, headers, K.status, status);
    sh.getRange(rowIdx, 1, 1, lastCol).setValues([arr]);
    invalidateExam_(id);
    return { id: id };
  });
}
function newExamId_() {
  return 'UJ-' + Utilities.formatDate(new Date(), 'GMT+7', 'yyMMdd') + '-' + Math.random().toString(36).slice(2, 6).toUpperCase();
}

function deleteExam_(b) {
  const p = auth_(b.token, 'admin');
  return lock_(function () {
    const ex = readExams_().filter(function (e) { return e.id === str_(b.id); })[0];
    if (!ex) throw new Error('Ujian tidak ditemukan.');
    if (!inScope_(p, ex.sekolah)) throw new Error('Anda tidak berhak menghapus ujian sekolah lain.');
    sheet_(CONFIG.SHEET_UJIAN).deleteRow(ex.row);
    invalidateExam_(ex.id);
    return {};
  });
}

function monitoring_(b) {
  const p = auth_(b.token, 'admin');
  const now = Date.now();
  const em = {};
  readExams_().forEach(function (e) { em[e.id] = e; });

  const rows = [];
  const sh = sesiSheet_();
  const last = sh.getLastRow();
  if (last >= 2) {
    sh.getRange(2, 1, last - 1, SESI_HEADERS.length).getValues().forEach(function (r, i) {
      if (!r[0]) return;
      const s = sesiFromValues_(r, i + 2);
      if (!inScope_(p, s.sekolah)) return;
      const e = em[s.idUjian];
      let status = s.status, ket = s.ket;
      if (status === ST.WORK && e && now - s.mulai > e.durasi * 60000) { status = ST.DONE; ket = 'Waktu habis'; }
      rows.push({
        sesiId: s.id, idUjian: s.idUjian, namaUjian: e ? e.nama : s.idUjian, sekolah: s.sekolah,
        nomor: s.nomor, nama: s.nama, kelas: s.kelas, mulai: s.mulai, last: s.last,
        pelanggaran: s.pelanggaran, batas: e ? e.batas : 0, status: status,
        online: status === ST.WORK && now - s.last < CONFIG.ACTIVE_WINDOW_SEC * 1000, keterangan: ket
      });
    });
  }
  const summary = {
    total: rows.length,
    mengerjakan: rows.filter(function (r) { return r.status === ST.WORK; }).length,
    terkunci: rows.filter(function (r) { return r.status === ST.LOCK; }).length,
    selesai: rows.filter(function (r) { return r.status === ST.DONE; }).length,
    terputus: rows.filter(function (r) { return r.status === ST.WORK && !r.online; }).length
  };

  const logs = [];
  const lg = logSheet_();
  const ll = lg.getLastRow();
  if (ll >= 2) {
    const from = Math.max(2, ll - 149);
    lg.getRange(from, 1, ll - from + 1, LOG_HEADERS.length).getValues().reverse().forEach(function (r) {
      if (logs.length >= 30 || !inScope_(p, r[4])) return;
      logs.push({ waktu: toMs_(r[0]), nomor: str_(r[2]), nama: str_(r[3]), sekolah: str_(r[4]), jenis: str_(r[6]), ke: r[7], status: str_(r[8]) });
    });
  }
  return { rows: rows, summary: summary, logs: logs };
}

function sesiForAdmin_(p, sesiId) {
  const sh = sesiSheet_();
  const cur = getSesi_(sh, sesiId);
  if (!cur) throw new Error('Sesi tidak ditemukan.');
  if (!inScope_(p, cur.sekolah)) throw new Error('Akses ditolak.');
  return { sh: sh, cur: cur };
}

function unlock_(b) {
  const p = auth_(b.token, 'admin');
  return lock_(function () {
    const x = sesiForAdmin_(p, b.sesiId);
    setSesi_(x.sh, x.cur.row, { status: ST.WORK, pelanggaran: 0, ket: 'Kunci dibuka oleh ' + p.nama, last: new Date() });
    appendLog_(x.cur, 'Kunci dibuka oleh proktor (' + p.nama + ')', 0, ST.WORK);
    return {};
  });
}

function resetSesi_(b) {
  const p = auth_(b.token, 'admin');
  return lock_(function () {
    const x = sesiForAdmin_(p, b.sesiId);
    x.sh.deleteRow(x.cur.row); // siswa dapat memulai ujian dari awal
    return {};
  });
}

/**
 * Buka kunci langsung dari layar siswa: proktor memasukkan Username/NIP + password guru
 * (admin sekolah yang sama, atau admin SEMUA). Sesi dipulihkan, pelanggaran direset ke 0.
 */
function unlockProktor_(b) {
  const p = auth_(b.token, 'siswa');
  const u = str_(b.username), pw = str_(b.password);
  if (!u || !pw) throw new Error('Isi username/NIP dan password proktor.');
  const key = 'u_' + norm_(p.nomor);
  checkBrute_(key);
  const r = readTable_(sheet_(CONFIG.SHEET_GURU)).rows.filter(function (x) {
    return norm_(get_(x, K.username)) === norm_(u);
  })[0];
  const okPass = r && str_(get_(r, K.password)) === pw;
  const guruSekolah = r ? str_(get_(r, K.sekolah)) : '';
  const berhak = okPass && (isSuperKey_(guruSekolah) || norm_(guruSekolah) === norm_(p.sekolah));
  if (!berhak) {
    addFail_(key);
    throw new Error('Username/NIP atau password proktor salah, atau bukan guru sekolah ini.');
  }
  clearFail_(key);
  const namaGuru = str_(get_(r, K.namaGuru));
  return lock_(function () {
    const sh = sesiSheet_();
    const cur = getSesi_(sh, b.sesiId);
    if (!cur || cur.nomor !== p.nomor) throw new Error('Sesi ujian tidak ditemukan.');
    const ex = examById_(cur.idUjian, false);
    const now = Date.now();
    const sisa = sisaDetik_(cur, ex, now);
    if (cur.status === ST.DONE) return { status: ST.DONE, pelanggaran: cur.pelanggaran, batas: ex ? ex.batas : 3, sisaDetik: 0, keterangan: cur.ket };
    if (sisa <= 0) {
      setSesi_(sh, cur.row, { status: ST.DONE, ket: 'Waktu habis', selesai: new Date(now), last: new Date(now) });
      return { status: ST.DONE, pelanggaran: cur.pelanggaran, batas: ex ? ex.batas : 3, sisaDetik: 0, keterangan: 'Waktu habis' };
    }
    if (cur.status === ST.LOCK) {
      setSesi_(sh, cur.row, { status: ST.WORK, pelanggaran: 0, ket: 'Kunci dibuka proktor ' + namaGuru, last: new Date(now) });
      appendLog_(cur, 'Kunci dibuka di perangkat siswa oleh proktor (' + namaGuru + ')', 0, ST.WORK);
    }
    return { status: ST.WORK, pelanggaran: 0, batas: ex ? ex.batas : 3, sisaDetik: sisa };
  });
}