// ====================== CONFIG ======================
const SHEET_NAMES = {
  SENSEI: 'Sensei',
  MULAI: 'Mulai kelas',
  KURIKULUM: 'Kurikulum',
  HARILIBUR: 'HariLibur'
};

// Kode cabang untuk skema ID. Kunci = nama wilayah di kolom "wilayah" sheet Kelas (huruf kecil).
const KODE_CABANG = {
  'bandung': 'BDG',
  'brebes': 'BRB',
  'cirebon': 'CRB',
  'semarang': 'SMG',
  'sukabumi': 'SKB',
  'tangerang': 'TNG',
  'tangerang (bsd)': 'TNG',
  'tegal': 'TGL',
  'yogyakarta': 'YGY',
  'manado': 'MND'
};

// ====================== HELPER KODE ======================
// Nama wilayah (atau kode yang sudah benar) -> kode cabang, '' kalau tidak dikenal
function kodeCabang_(wilayah) {
  const w = String(wilayah || '').trim();
  if (KODE_CABANG[w.toLowerCase()]) return KODE_CABANG[w.toLowerCase()];
  const up = w.toUpperCase();
  return Object.keys(KODE_CABANG).some(k => KODE_CABANG[k] === up) ? up : '';
}

function pad_(n, len) { return ('000' + n).slice(-len); }

// Cari nomor kolom (1-based) berdasarkan judul header; -1 jika tidak ada
function colIdx_(headers, name) {
  return headers.map(h => String(h).trim().toLowerCase()).indexOf(name) + 1 || -1;
}

// ====================== WEB APP ======================
function doGet(e) {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Joy Progress Monitor')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// ====================== HELPER ======================
function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function tz_() { return ss_().getSpreadsheetTimeZone(); }

function todayStr_() {
  return Utilities.formatDate(new Date(), tz_(), 'yyyy-MM-dd');
}

// Ubah nilai sel tanggal (Date atau teks "2026/7/14", "2026-07-14") menjadi 'yyyy-MM-dd'
function fmtDate_(v, tz) {
  if (v instanceof Date) return Utilities.formatDate(v, tz || tz_(), 'yyyy-MM-dd');
  const m = String(v || '').trim().match(/^(\d{4})[\/\-.](\d{1,2})[\/\-.](\d{1,2})/);
  if (!m) return '';
  return m[1] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[3]).slice(-2);
}

// Baca sheet -> array of object, kunci = nama header (huruf kecil).
// Jadi urutan kolom di spreadsheet boleh berubah tanpa merusak kode.
function readSheet_(name) {
  const sheet = ss_().getSheetByName(name);
  if (!sheet) throw new Error('Sheet "' + name + '" tidak ditemukan');
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0].map(h => { const x = String(h).trim().toLowerCase(); return /^batch\b/.test(x) ? 'batch' : x; });
  const rows = [];
  for (let i = 1; i < values.length; i++) {
    const row = values[i];
    if (row.every(c => c === '' || c === null)) continue;
    const obj = { _row: i + 1 };
    headers.forEach((h, c) => { if (h) obj[h] = row[c]; });
    rows.push(obj);
  }
  return rows;
}

function dayOfWeek_(dateStr) {
  const p = dateStr.split('-').map(Number);
  return new Date(Date.UTC(p[0], p[1] - 1, p[2])).getUTCDay(); // 0=Minggu, 6=Sabtu
}

function getLiburMap_() {
  const map = {};
  readSheet_(SHEET_NAMES.HARILIBUR).forEach(r => {
    const t = fmtDate_(r.tanggal);
    if (t) map[t] = String(r.keterangan || 'Libur');
  });
  return map;
}

// ====================== DATABASE KURIKULUM (per cabang) ======================
// Spreadsheet "Database Kurikulum": satu tab per cabang (nama tab = nama cabang, seperti Database Kelas).
// Satu tab berisi beberapa BLOK kurikulum bertumpuk ke bawah, satu blok per group (Joycare, Joybridge, Joygreen, Joyserve):
//   baris judul  : A = nama group (mis. JOYCARE; boleh kosong) | "Batch 1" di D | "Batch 2" di F | ...   (tiap batch 2 kolom)
//   baris header : A hari_ke | B 前後 | C materi | 担当者 | 指摘 | 担当者 | 指摘 | ...
//   baris data   : tiap hari 2 baris (午前 lalu 午後); hari_ke dan materi boleh di-merge 2 baris
// Blok dikenali dari baris header (kolom B = 前後), jadi posisi barisnya bebas. Nama group diambil dari kolom A baris judul;
// kalau kosong dipakai urutan blok: Joycare, Joybridge, Joygreen, Joyserve.
// Kurikulum utama (bahasa Jepang) = hari 1 sampai UTAMA_HARI dari blok pertama; hari berikutnya khusus group.
const ID_DB_KURIKULUM = '1Xw-ZZmGxk61QxvIdcHf8smL6p1FiofuFv_C_exH3JUE';
const UTAMA_HARI = 98;
// Batas baris tiap blok group di tab cabang (nomor baris sheet, 1-based). Kalau sebuah blok berada di rentang ini,
// namanya diambil dari sini dan datanya dibaca sampai baris 'akhir' (baris kosong di tengah dilewati, tidak memutus pembacaan).
// akhir: 0 = tanpa batas (sampai header blok berikutnya). Joygreen & Joyserve tinggal ditambah di bawah.
const BATAS_BLOK = [
  { grup: 'Joycare',   awal: 1,   akhir: 275 },
  { grup: 'Joybridge', awal: 276, akhir: 545 }   // data mulai baris 279; awal 276 supaya baris judul/header di atasnya ikut terbaca
  // , { grup: 'Joygreen',  awal: 549, akhir: 0 }
  // , { grup: 'Joyserve',  awal: 0,   akhir: 0 }
];
const GRUP_KURIKULUM = ['Joycare', 'Joybridge', 'Joygreen', 'Joyserve'];
let KUR_CACHE_ = {};

function teks_(x) { return String(x == null ? '' : x).replace(/\s+/g, ' ').trim(); }
function noBatch_(s) { const m = String(s || '').match(/\d+/); return m ? Number(m[0]) : 0; }

// Baca tab cabang di Database Kurikulum -> { sheet, v, tab, blok: [{grup, key, hari:[{hari_ke, items, rows}], batch:{no:{p,c}}}] }
// rows = nomor baris (1-based) tiap sesi; batch[no] = kolom (0-based) 担当者 (p) dan 指摘 (c). `segar` = abaikan cache.
function bukaKurikulum_(cabang, segar) {
  const key = kunciTab_(cabang);
  if (!segar && KUR_CACHE_[key]) return KUR_CACHE_[key];
  let db;
  try {
    db = SpreadsheetApp.openById(ID_DB_KURIKULUM);
  } catch (e) {
    throw new Error('Spreadsheet Database Kurikulum tidak bisa dibuka. Pastikan akun pemilik script punya akses ke spreadsheet tersebut.');
  }
  const sheet = db.getSheets().filter(sh => sh.getName().trim().toLowerCase() === key)[0];
  if (!sheet) throw new Error('Tab untuk cabang "' + cabang + '" tidak ditemukan di Database Kurikulum');

  const v = sheet.getDataRange().getValues();
  const adaSesi = r => r >= 0 && r < v.length && SESI.indexOf(teks_(v[r][1])) >= 0;
  const blok = [];

  for (let r = 0; r < v.length; r++) {
    if (!/^(前後|sesi)$/i.test(teks_(v[r][1]))) continue;            // baris header blok

    const judul = r >= 1 && !adaSesi(r - 1) ? r - 1 : -1;            // baris judul (nama group + label Batch) tepat di atas header
    const namaJudul = judul >= 0 ? teks_(v[judul][0]) : '';
    const dikenal = GRUP_KURIKULUM.filter(g => g.toLowerCase() === namaJudul.toLowerCase())[0];
    const ruang = BATAS_BLOK.filter(b => r + 1 >= b.awal && (!b.akhir || r + 1 <= b.akhir))[0] || null;   // rentang baris blok ini (kalau diatur)
    const grup = (ruang ? ruang.grup : '') || dikenal || (namaJudul && !/^hari/i.test(namaJudul) ? namaJudul : (GRUP_KURIKULUM[blok.length] || 'Group ' + (blok.length + 1)));

    // pasangan kolom per batch: 担当者 (pengajar) lalu 指摘 (catatan); label "Batch N" ada di baris judul
    const batch = {};
    let urut = 0;
    for (let c = 0; c < v[r].length; c++) {
      if (!/担当|pengajar/i.test(teks_(v[r][c]))) continue;
      urut++;
      const label = judul >= 0 ? teks_(v[judul][c]) : '';
      batch[noBatch_(label) || urut] = { p: c, c: c + 1 };
    }

    // baris data: berurutan selama kolom B berisi 午前 / 午後
    const hari = [];
    let cur = null;
    const batas = ruang && ruang.akhir ? Math.min(ruang.akhir, v.length) : 0;   // nomor baris terakhir blok (0 = ikuti baris 午前/午後 berurutan)
    for (let rr = r + 1; batas ? rr < batas : adaSesi(rr); rr++) {
      if (batas) {
        if (/^(前後|sesi)$/i.test(teks_(v[rr][1]))) break;      // header blok berikutnya
        if (!adaSesi(rr)) continue;                              // baris kosong/pemisah di dalam blok: dilewati
      }
      const a = v[rr][0];
      if (a !== '' && a !== null && !isNaN(Number(a))) {              // awal hari baru
        cur = { hari_ke: Number(a), items: [], rows: {} };
        hari.push(cur);
      }
      if (!cur) continue;
      const sesi = teks_(v[rr][1]);
      let materi = String(v[rr][2] == null ? '' : v[rr][2]).trim();
      if (!materi && cur.items.length) materi = cur.items[cur.items.length - 1].materi;   // materi di-merge 2 baris
      cur.items.push({ sesi: sesi, materi: materi });
      cur.rows[sesi] = rr + 1;
    }
    blok.push({ grup: grup, key: grup.toLowerCase(), hari: hari, batch: batch });
  }
  const out = { sheet: sheet, v: v, tab: sheet.getName(), blok: blok };
  KUR_CACHE_[key] = out;
  return out;
}

// Kurikulum yang dipakai satu kelas: blok group-nya (kalau ada), kalau tidak kurikulum utama.
// Bentuknya {hari, kolomTanggal} supaya cocok dengan infoHari_.
function kurikulumKelas_(k) {
  const kd = bukaKurikulum_(k.cabang);
  const g = norm_(k.group);
  const blok = kd.blok.filter(b => b.key === g)[0] || null;
  const hari = blok ? blok.hari : (kd.blok[0] ? kd.blok[0].hari.filter(h => h.hari_ke <= UTAMA_HARI) : []);
  return { hari: hari, kolomTanggal: {}, blok: blok, kd: kd };
}

// {"3|午前": {p: pengajar, c: catatan}, ...} untuk satu batch di satu blok
function catatanBatch_(kd, blok, noBatch) {
  const out = {};
  const bc = blok && noBatch ? blok.batch[noBatch] : null;
  if (!bc) return out;
  blok.hari.forEach(h => SESI.forEach(s => {
    const r = h.rows[s];
    if (!r) return;
    const p = teks_(kd.v[r - 1][bc.p]), c = String(kd.v[r - 1][bc.c] == null ? '' : kd.v[r - 1][bc.c]).trim();
    if (p || c) out[h.hari_ke + '|' + s] = { p: p, c: c };
  }));
  return out;
}

// Kotak catatan di Home: hari ini; kalau hari ini libur, hari terakhir yang sudah berjalan
function catatanHome_(kurK, batches, info) {
  const nb = noBatch_(batches[0]);
  const blok = kurK.blok;
  const hari = info.hari_ke || info.selesai || 0;
  const out = { bisa: false, hari: hari, terakhir: !info.hari_ke, pagi: '', siang: '', batch: nb };
  if (!blok || !nb || !blok.batch[nb] || hari < 1 || !blok.hari.some(x => x.hari_ke === hari)) return out;
  const c = catatanBatch_(kurK.kd, blok, nb);
  out.bisa = true;
  out.pagi = (c[hari + '|' + SESI[0]] || {}).c || '';
  out.siang = (c[hari + '|' + SESI[1]] || {}).c || '';
  return out;
}

function addDays_(dateStr, n) {
  const p = dateStr.split('-').map(Number);
  const d = new Date(Date.UTC(p[0], p[1] - 1, p[2] + n));
  return d.getUTCFullYear() + '-' + ('0' + (d.getUTCMonth() + 1)).slice(-2) + '-' + ('0' + d.getUTCDate()).slice(-2);
}

// Hitung pertemuan ke-berapa pada `tanggal`, dimulai dari `start`.
// Sabtu, Minggu, dan tanggal di sheet HariLibur TIDAK dihitung.
// hari_ke = 0 kalau `tanggal` libur. hari_kerja = jumlah hari belajar dari start s/d tanggal (termasuk tanggal bila bukan libur).
function hitungJadwal_(start, tanggal, liburMap) {
  let count = 0, hariKe = 0, libur = '';
  let d = start;
  for (let i = 0; d <= tanggal && i < 1500; i++) {
    const dow = dayOfWeek_(d);
    const ket = liburMap[d] || ((dow === 0 || dow === 6) ? 'Akhir pekan' : '');
    if (!ket) { count++; if (d === tanggal) hariKe = count; }
    else if (d === tanggal) libur = ket;
    d = addDays_(d, 1);
  }
  return { hari_ke: hariKe, hari_kerja: count, libur_ket: libur };
}

// Info satu kelas (baris sheet "Mulai kelas") pada satu tanggal: pertemuan ke-berapa, materi, libur, progres.
function infoHari_(k, kur, liburMap, tanggal) {
  const planned = kur.hari.length;
  const dow = dayOfWeek_(tanggal);
  const libur = liburMap[tanggal] || ((dow === 0 || dow === 6) ? 'Akhir pekan' : '');
  const start = fmtDate_(k.start_date);
  const out = { planned: planned, start: start, libur_ket: libur, hari_ke: 0, selesai: 0, status: 'ok', items: [] };

  if (!start) {   // cadangan: tanggal di kolom cabang sheet Kurikulum
    const w = kunciTab_(k.cabang);
    if (!(w in kur.kolomTanggal)) { out.status = 'no_start'; return out; }
    const h = kur.hari.find(x => x.dates[w] === tanggal);
    out.selesai = kur.hari.filter(x => x.dates[w] && x.dates[w] < tanggal).length;
    if (h) { out.hari_ke = h.hari_ke; out.items = h.items; } else out.status = libur ? 'libur' : 'no_date';
    return out;
  }
  if (start > tanggal) { out.status = 'belum_mulai'; return out; }

  const j = hitungJadwal_(start, tanggal, liburMap);
  out.hari_ke = j.hari_ke;
  out.selesai = Math.min(planned, j.hari_ke ? j.hari_ke - 1 : j.hari_kerja);
  if (!j.hari_ke) { out.status = 'libur'; out.libur_ket = j.libur_ket; return out; }
  if (j.hari_ke > planned) { out.status = 'selesai'; return out; }
  const h = kur.hari.find(x => x.hari_ke === j.hari_ke);
  if (!h) { out.status = 'no_materi'; return out; }
  out.items = h.items;
  return out;
}

const SESI = ['午前', '午後'];
// Sesi saat ini menurut jam spreadsheet: sebelum jam 12 = 午前, setelahnya = 午後
function sesiSekarang_() {
  return Number(Utilities.formatDate(new Date(), tz_(), 'H')) < 12 ? SESI[0] : SESI[1];
}

function isAktif_(v) { return String(v || '').trim().toLowerCase() === 'aktif'; }
function norm_(v) { return String(v || '').trim().toLowerCase(); }

// ====================== HAK AKSES (Sensei / Admin) ======================
// Admin = baris di sheet Sensei dengan kolom "role" berisi Admin. Admin hanya MEMANTAU:
//  - boleh melihat semua cabang (atau satu cabang pilihan),
//  - tidak boleh menyimpan absensi / catatan (ditolak di server).
// Sensei biasa tetap hanya melihat cabangnya sendiri (kolom wilayah_utama); pilihan cabang dari browser diabaikan.
function ambilAkun_(idSensei) {
  const s = readSheet_(SHEET_NAMES.SENSEI).find(r => norm_(r.id_sensei) === norm_(idSensei));
  if (!s || !isAktif_(s.status)) throw new Error('Akun sensei tidak ditemukan atau tidak aktif');
  return s;
}
function isAdmin_(s) { return norm_(s && s.role) === 'admin'; }

// -> { admin, gabung, cabang: [nama cabang yang boleh dilihat] }. gabung = beberapa cabang sekaligus.
function cakupan_(sensei, pilih) {
  if (!isAdmin_(sensei)) {
    const c = String(sensei.wilayah_utama || '').trim();
    if (!c) throw new Error('Cabang akun ini belum diisi (kolom wilayah_utama)');
    return { admin: false, gabung: false, cabang: [c] };
  }
  const p = kunciTab_(pilih);
  const cocok = p ? CABANG_LIST.filter(c => kunciTab_(c) === p)[0] : '';
  return cocok ? { admin: true, gabung: false, cabang: [cocok] }
               : { admin: true, gabung: true, cabang: CABANG_LIST.slice() };
}

// Mode "semua cabang": nama kelas yang sama di dua cabang tidak boleh tercampur -> kunci "Nama · Cabang"
function kelasKunci_(kelas, cabang, gabung) {
  return gabung ? (String(kelas || '').trim() || '(tanpa kelas)') + ' · ' + cabang : kelas;
}

// Pembaca tab absen per cabang (di-cache). Untuk Admin, tab yang belum ada dilewati (dianggap kosong).
function pembacaAbsen_(cak) {
  const cache = {};
  return cabang => {
    const key = kunciTab_(cabang);
    if (!(key in cache)) {
      try { cache[key] = bacaAbsenTab_(cabang); }
      catch (e) { if (!cak.admin) throw e; cache[key] = { siswa: [], blok: [], v: [], tz: tz_() }; }
    }
    return cache[key];
  };
}

// Kurikulum satu kelas; kalau tab kurikulum cabangnya belum ada, pakai kurikulum kosong (khusus Admin)
function kurikulumKelasAman_(k) {
  try { return kurikulumKelas_(k); }
  catch (e) { return { hari: [], kolomTanggal: {}, blok: null, kd: { v: [], blok: [], tab: '' } }; }
}


// ====================== AKUN SENSEI ======================
const CABANG_LIST = ['Tangerang (BSD)', 'Bandung', 'Brebes', 'Cirebon', 'Semarang', 'Sukabumi', 'Tegal', 'Yogyakarta', 'Manado'];
const GROUP_LIST = ['Joycare', 'Joybridge', 'Joygreen', 'Joyserve'];
const MIN_PASSWORD = 6;

// Buat akun sensei baru. ID otomatis: JC + kode cabang + nomor 3 digit per cabang (JCBDG001, JCCRB001, JCCRB002, ...)
// Status langsung Aktif.
function buatAkunSensei(p) {
  p = p || {};
  const nama = String(p.nama || '').trim();
  const password = String(p.password || '');
  if (!nama) throw new Error('Nama wajib diisi');
  if (CABANG_LIST.indexOf(p.cabang) < 0) throw new Error('Cabang tidak valid');
  // Group boleh dari daftar atau diketik sendiri (opsi "Lainnya..."); nama yang cocok daftar dirapikan hurufnya
  let group = String(p.group || '').trim();
  if (!group || /^lainnya/i.test(group) || group.indexOf('--') === 0) throw new Error('Group wajib dipilih atau diisi');
  if (group.length > 40) throw new Error('Nama group terlalu panjang (maksimal 40 karakter)');
  const dikenal = GROUP_LIST.filter(g => g.toLowerCase() === group.toLowerCase())[0];
  if (dikenal) group = dikenal;
  if (password.length < MIN_PASSWORD) throw new Error('Password minimal ' + MIN_PASSWORD + ' karakter');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = ss_().getSheetByName(SHEET_NAMES.SENSEI);
    if (!sheet) throw new Error('Sheet "Sensei" tidak ditemukan');
    let lastCol = sheet.getLastColumn();
    let headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

    // Kolom "group" ditambahkan otomatis di ujung kanan kalau belum ada
    if (colIdx_(headers, 'group') < 0) {
      sheet.getRange(1, lastCol).copyTo(sheet.getRange(1, lastCol + 1), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
      sheet.getRange(1, lastCol + 1).setValue('group');
      lastCol++;
      headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    }

    const cId = colIdx_(headers, 'id_sensei'), cNama = colIdx_(headers, 'nama'), cGroup = colIdx_(headers, 'group'),
          cWil = colIdx_(headers, 'wilayah_utama'), cStatus = colIdx_(headers, 'status'), cPass = colIdx_(headers, 'password');
    if ([cId, cNama, cWil, cStatus, cPass].some(c => c < 0)) {
      throw new Error('Sheet Sensei harus punya kolom id_sensei, nama, wilayah_utama, status, password');
    }

    const kode = kodeCabang_(p.cabang);
    if (!kode) throw new Error('Kode cabang untuk "' + p.cabang + '" tidak ditemukan');
    const lastRow = sheet.getLastRow();
    let maks = 0;
    if (lastRow >= 2) {
      sheet.getRange(2, cId, lastRow - 1, 1).getValues().forEach(r => {
        const m = String(r[0]).trim().toUpperCase().match(/^JC([A-Z]{3})(\d{3})$/);
        if (m && m[1] === kode) maks = Math.max(maks, Number(m[2]));
      });
    }
    const idBaru = 'JC' + kode + pad_(maks + 1, 3);

    const row = new Array(lastCol).fill('');
    row[cId - 1] = idBaru;
    row[cNama - 1] = nama;
    row[cGroup - 1] = group;
    row[cWil - 1] = p.cabang;
    row[cStatus - 1] = 'Aktif';
    row[cPass - 1] = password;

    const r = lastRow + 1;
    sheet.getRange(r, cPass).setNumberFormat('@'); // teks, supaya "0123" tidak jadi 123
    sheet.getRange(r, 1, 1, lastCol).setValues([row]);
    return { success: true, id_sensei: idBaru, message: 'Akun berhasil dibuat' };
  } finally {
    lock.releaseLock();
  }
}

// Ganti password: wajib menyebut ID dan password lama
function gantiPassword(idSensei, passwordLama, passwordBaru) {
  idSensei = String(idSensei || '').trim();
  passwordBaru = String(passwordBaru || '');
  if (!idSensei || !passwordLama) return { success: false, message: 'ID Sensei dan password lama wajib diisi' };
  if (passwordBaru.length < MIN_PASSWORD) return { success: false, message: 'Password baru minimal ' + MIN_PASSWORD + ' karakter' };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const sheet = ss_().getSheetByName(SHEET_NAMES.SENSEI);
    if (!sheet) throw new Error('Sheet "Sensei" tidak ditemukan');
    const v = sheet.getDataRange().getValues();
    const cId = colIdx_(v[0], 'id_sensei'), cPass = colIdx_(v[0], 'password'), cStatus = colIdx_(v[0], 'status');
    if (cId < 0 || cPass < 0) throw new Error('Kolom id_sensei / password tidak ditemukan di sheet Sensei');

    for (let i = 1; i < v.length; i++) {
      if (norm_(v[i][cId - 1]) !== norm_(idSensei)) continue;
      if (String(v[i][cPass - 1]) !== String(passwordLama)) return { success: false, message: 'ID Sensei atau password lama salah' };
      if (cStatus > 0 && !isAktif_(v[i][cStatus - 1])) return { success: false, message: 'Akun sensei tidak aktif' };
      const cell = sheet.getRange(i + 1, cPass);
      cell.setNumberFormat('@');
      cell.setValue(passwordBaru);
      return { success: true, message: 'Password berhasil diganti. Silakan masuk dengan password baru.' };
    }
    return { success: false, message: 'ID Sensei atau password lama salah' };
  } finally {
    lock.releaseLock();
  }
}

// ====================== DATABASE KELAS (per cabang) ======================
// Spreadsheet "Database Kelas": satu tab per cabang (Tangerang, Bandung, Tegal, ...).
// Kolom dibaca lewat judul baris 1 untuk tiap blok group (Joycare, Joybridge, Joygreen, Joyserve):
// Nama, Batch, Group, Kelas; kolom setelah Kelas (JFT, SSW 1, SSW 2, SSW 3, NAITEI) ikut ditampilkan.
const ID_DB_KELAS = '1AeLRUjpmFsIyuroeRvkS465qNjjLB7RNeHxtafRgRRw';

// "Tangerang (BSD)" -> "tangerang" (sama dengan nama tab)
function kunciTab_(cabang) {
  return String(cabang || '').replace(/\(.*?\)/g, '').trim().toLowerCase();
}

// Cabang ditentukan SERVER dari akun sensei (kolom wilayah_utama), bukan dari kiriman browser,
// jadi sensei hanya mendapat siswa dari tab cabangnya sendiri.
// Baca satu tab cabang di Database Kelas -> { tab, siswa } (null kalau tab cabang itu tidak ada)
function bacaKelasTab_(db, cabang) {
  const key = kunciTab_(cabang);
  const sheet = db.getSheets().filter(sh => sh.getName().trim().toLowerCase() === key)[0];
  if (!sheet) return null;

  const v = sheet.getDataRange().getDisplayValues();
  if (v.length < 2) return { tab: sheet.getName(), siswa: [] };

  const H = v[0].map(h => String(h).trim());
  const GRUP_BLOK = ['Joycare', 'Joybridge', 'Joygreen', 'Joyserve']; // urutan blok dari kiri ke kanan

  // Satu tab berisi beberapa blok tabel berdampingan (A-J Joycare, L-U Joybridge, W-AF Joygreen, AH-AQ Joyserve),
  // dipisah satu kolom dengan judul kosong. Blok dimulai di kolom berjudul "Nama" dan berakhir di judul kosong berikutnya.
  const blok = [];
  for (let c = 0; c < H.length; c++) {
    if (H[c].toLowerCase() !== 'nama') continue;
    let end = c + 1;
    while (end < H.length && H[end]) end++;
    const cari = nama => { for (let k = c; k < end; k++) if (H[k].toLowerCase() === nama) return k; return -1; };
    const cBatch = cari('batch'), cGroup = cari('group'), cKelas = cari('kelas');
    if (cBatch < 0) { c = end; continue; }
    const extras = []; // JFT, SSW 1, SSW 2, SSW 3, NAITEI
    for (let k = Math.max(c, cBatch, cGroup, cKelas) + 1; k < end; k++) extras.push({ k: H[k], c: k });
    blok.push({ cNama: c, cBatch: cBatch, cGroup: cGroup, cKelas: cKelas, extras: extras, grup: GRUP_BLOK[blok.length] || '' });
    c = end;
  }
  if (!blok.length) throw new Error('Tab "' + sheet.getName() + '" harus punya judul kolom Nama dan Batch di baris 1');

  const siswa = [];
  blok.forEach(b => {
    for (let i = 1; i < v.length; i++) {
      const r = v[i];
      const nama = String(r[b.cNama]).trim();
      if (!nama) continue;
      siswa.push({
        nama: nama,
        batch: String(r[b.cBatch]).trim() || '(tanpa batch)',
        group: (b.cGroup >= 0 && String(r[b.cGroup]).trim()) || b.grup,
        kelas: b.cKelas >= 0 ? String(r[b.cKelas]).trim() : '',
        extra: b.extras.map(x => ({ k: x.k, v: String(r[x.c]).trim() })).filter(x => x.v)
      });
    }
  });
  return { tab: sheet.getName(), siswa: siswa };
}

// Cabang ditentukan SERVER dari akun (wilayah_utama untuk sensei; pilihan sidebar untuk Admin), bukan dari kiriman browser.
function getDataKelas(idSensei, pilih) {
  const sensei = ambilAkun_(idSensei);
  const cak = cakupan_(sensei, pilih);

  let db;
  try {
    db = SpreadsheetApp.openById(ID_DB_KELAS);
  } catch (e) {
    throw new Error('Spreadsheet Database Kelas tidak bisa dibuka. Pastikan akun pemilik script punya akses ke spreadsheet tersebut.');
  }

  const siswa = [], tabs = [], peringatan = [];
  cak.cabang.forEach(c => {
    let t;
    try { t = bacaKelasTab_(db, c); }
    catch (e) { if (!cak.admin) throw e; peringatan.push(c + ': ' + e.message); return; }
    if (!t) {
      if (!cak.admin) throw new Error('Tab untuk cabang "' + c + '" tidak ditemukan di Database Kelas');
      return;                                          // Admin: cabang tanpa tab dilewati
    }
    tabs.push(t.tab);
    t.siswa.forEach(s => {
      s.cabang = c;
      s.kelas = kelasKunci_(s.kelas, c, cak.gabung);
      siswa.push(s);
    });
  });
  return {
    success: true, cabang: cak.gabung ? 'Semua cabang' : cak.cabang[0], tab: tabs.join(', '), siswa: siswa,
    gabung: cak.gabung, admin: cak.admin, peringatan: peringatan
  };
}

// ====================== API ======================
function loginSensei(idSensei, password) {
  const rows = readSheet_(SHEET_NAMES.SENSEI);
  for (const r of rows) {
    if (String(r.id_sensei).toLowerCase() === String(idSensei).trim().toLowerCase() &&
        String(r.password) === String(password)) {
      if (!isAktif_(r.status)) return { success: false, message: 'Akun sensei tidak aktif' };
      return {
        success: true,
        sensei: {
          id: r.id_sensei, nama: r.nama, group: r.group, wilayah: r.wilayah_utama, status: r.status,
          role: String(r.role || '').trim(), admin: isAdmin_(r), cabang_list: isAdmin_(r) ? CABANG_LIST : []
        }
      };
    }
  }
  return { success: false, message: 'ID Sensei atau password salah' };
}

// ====================== KELAS & ABSENSI ======================
// Kelas aktif  : sheet "Mulai kelas" (Database Akun): NO | Cabang | Kelas | Group | Start_date | End_date  (+ Batch, opsional)
// Siswa+absensi: Database Kelas, tab "<Cabang> (absen)": A No | B Nama | C Batch | D Group | E Kelas | F (pemisah) | G dst = absensi
// Tiap BATCH adalah satu blok baris siswa. Dua baris tepat di atas blok = header:
//     baris tanggal : G = tanggal (di-merge G:H)
//     baris sesi    : G = 午前, H = 午後
// Tanggal berikutnya memakai pasangan kolom berikutnya (I:J, K:L, ...). Isi sel: O = Hadir, S = Sakit, I = Izin, TK = Tanpa keterangan.
// Kalau di atas blok belum ada 2 baris kosong untuk header, script menyisipkan 2 baris sendiri.
// Khusus blok pertama (siswa mulai baris 2) tanggal ditulis di baris 1 (G1) tanpa baris sesi.
const KOLOM_TANGGAL_AWAL = 7; // kolom G (午前); H = 午後
const KODE_ABSEN = { 'Hadir': 'O', 'Sakit': 'S', 'Izin': 'I', 'Tidak Hadir': 'TK' };
const STATUS_DARI_KODE = { 'O': 'Hadir', '0': 'Hadir', '○': 'Hadir', 'S': 'Sakit', 'I': 'Izin', 'TK': 'Tidak Hadir' };

function openDbKelas_() {
  try {
    return SpreadsheetApp.openById(ID_DB_KELAS);
  } catch (e) {
    throw new Error('Spreadsheet Database Kelas tidak bisa dibuka. Pastikan akun pemilik script punya akses ke spreadsheet tersebut.');
  }
}

// Dua baris header (tanggal + sesi) untuk blok yang siswa pertamanya di baris f (1-based).
// null = belum ada baris kosong di atas blok.
function headerBaris_(v, cNama, f) {
  const kosong = r => r >= 1 && r <= v.length && !String(v[r - 1][cNama]).trim();
  if (f - 2 >= 1 && kosong(f - 2) && kosong(f - 1)) return { d: f - 2, s: f - 1 };
  if (f === 2) return { d: 1, s: 0 };   // blok pertama: tanggal di baris 1, tanpa baris sesi
  return null;
}

// Baca tab "<cabang> (absen)": siswa dikelompokkan per blok (baris berurutan dengan kelas+group+batch sama)
function bacaAbsenTab_(cabang) {
  const db = openDbKelas_();
  const nama = kunciTab_(cabang) + ' (absen)';
  const sheet = db.getSheets().filter(sh => sh.getName().trim().toLowerCase().replace(/\s+/g, ' ') === nama)[0];
  if (!sheet) throw new Error('Tab "' + kunciTab_(cabang) + ' (absen)" tidak ditemukan di Database Kelas');

  const v = sheet.getDataRange().getValues();
  const H = (v[0] || []).map(h => String(h).trim().toLowerCase());
  const cNama = H.indexOf('nama'), cBatch = H.indexOf('batch'), cGroup = H.indexOf('group'), cKelas = H.indexOf('kelas');
  if (cNama < 0 || cKelas < 0) throw new Error('Tab "' + sheet.getName() + '" harus punya judul kolom Nama dan Kelas di baris 1');
  const tz = db.getSpreadsheetTimeZone();

  const siswa = [], blok = [];
  let cur = null;
  for (let i = 1; i < v.length; i++) {
    const nm = String(v[i][cNama]).trim();
    if (!nm || nm.toLowerCase() === 'nama') { cur = null; continue; }
    const s = {
      row: i + 1,
      nama: nm,
      batch: cBatch >= 0 ? String(v[i][cBatch]).trim() : '',
      group: cGroup >= 0 ? String(v[i][cGroup]).trim() : '',
      kelas: String(v[i][cKelas]).trim()
    };
    const kunci = [norm_(s.kelas), norm_(s.group), norm_(s.batch)].join('|');
    if (!cur || cur.kunci !== kunci) {
      cur = { kunci: kunci, f: i + 1, last: i + 1, hdr: null, pairs: {} };
      blok.push(cur);
    }
    cur.last = i + 1;
    s.blok = blok.length - 1;
    siswa.push(s);
  }

  // tanggal tiap blok: pasangan kolom G:H, I:J, ... pada baris tanggal blok itu
  blok.forEach(b => {
    b.hdr = headerBaris_(v, cNama, b.f);
    if (!b.hdr) return;
    const baris = v[b.hdr.d - 1];
    for (let c = KOLOM_TANGGAL_AWAL - 1; c < baris.length; c += 2) {
      const d = fmtDate_(baris[c], tz);
      if (d) b.pairs[d] = c;
    }
  });
  return { sheet: sheet, v: v, siswa: siswa, blok: blok, tz: tz, cNama: cNama };
}

function siswaKelas_(absen, k) {
  const nk = norm_(k.kelas), ng = norm_(k.group), nb = norm_(k.batch);
  return absen.siswa.filter(s =>
    norm_(s.kelas) === nk &&
    (!ng || !s.group || norm_(s.group) === ng) &&
    (!nb || norm_(s.batch) === nb));
}

// '' kalau tanggal itu belum ada kolomnya / sel kosong
function statusSiswa_(absen, s, tanggal, sesi) {
  const c = absen.blok[s.blok].pairs[tanggal];
  if (c === undefined) return '';
  const x = absen.v[s.row - 1][c + (sesi === SESI[1] ? 1 : 0)];
  return STATUS_DARI_KODE[String(x === undefined ? '' : x).trim().toUpperCase()] || '';
}

// Kelas milik sensei: baris "Mulai kelas" dengan Cabang = wilayah_utama sensei.
// Semua kelas di cabang itu tampil (semua group: Joycare, Joybridge, dst), tidak difilter menurut group akun.
// Kelas hilang dari Home setelah End_date lewat. Yang tampil hanya kelas yang SEDANG BERJALAN (Start_date <= hari ini);
// kalau belum ada yang berjalan, kelas yang belum mulai tetap ditampilkan supaya muncul keterangan "Kelas mulai pada ...".
function kelasSensei_(sensei, cak) {
  cak = cak || cakupan_(sensei, '');
  const today = todayStr_();
  const mulai = readSheet_(SHEET_NAMES.MULAI).filter(k => String(k.kelas || '').trim());
  let hasil = [];
  cak.cabang.forEach(c => {                       // aturan "berjalan / belum mulai" berlaku per cabang
    const kunci = kunciTab_(c);
    const semua = mulai.filter(k => kunciTab_(k.cabang) === kunci);
    const belumLewat = semua.filter(k => {
      const end = fmtDate_(k.end_date);
      return !end || end >= today;
    });
    const berjalan = belumLewat.filter(k => {
      const st = fmtDate_(k.start_date);
      return !st || st <= today;
    });
    hasil = hasil.concat(berjalan.length ? berjalan : belumLewat);
  });
  return hasil;
}

// Batch yang sedang berjalan untuk satu baris "Mulai kelas":
//  - kalau sheet Mulai kelas punya kolom Batch dan terisi -> pakai itu (satu baris = satu batch, dengan Start/End_date sendiri)
//  - kalau tidak -> sementara ambil batch paling baru (nomor terbesar) dari siswa kelas itu
function batchBerjalan_(k, siswa) {
  const b = String(k.batch || '').trim();
  if (b) return [b];
  const list = uniqBatch_(siswa);
  return list.length ? [list[list.length - 1]] : [];
}

function uniqBatch_(list) {
  return [...new Set(list.map(s => s.batch).filter(Boolean))]
    .sort((a, b) => a.localeCompare(b, undefined, { numeric: true }));
}

// Data untuk halaman Home
function getHomeData(idSensei, pilih) {
  const today = todayStr_();
  const sensei = ambilAkun_(idSensei);
  const cak = cakupan_(sensei, pilih);

  const kelasRows = kelasSensei_(sensei, cak);
  const sesi = sesiSekarang_();
  let result = [];

  if (kelasRows.length) {
    const liburMap = getLiburMap_();
    const absenDari = pembacaAbsen_(cak);

    result = kelasRows.map(k => {
      const absen = absenDari(k.cabang);
      const batches = batchBerjalan_(k, siswaKelas_(absen, k));
      const siswa = siswaKelas_(absen, k).filter(s => !batches.length || batches.indexOf(s.batch) >= 0);
      const kurK = cak.admin ? kurikulumKelasAman_(k) : kurikulumKelas_(k);
      const info = infoHari_(k, kurK, liburMap, today);
      // Admin hanya memantau: kotak catatan tidak dikirim
      const catatan = cak.admin ? { bisa: false, hari: 0, terakhir: false, pagi: '', siang: '', batch: 0 } : catatanHome_(kurK, batches, info);
      const stat = siswa.map(s => ({ s: s, st: statusSiswa_(absen, s, today, sesi) }));
      const sudahAbsen = stat.some(x => x.st !== '');
      const hitung = st => stat.filter(x => x.st === st).length;
      const kode = { 'Sakit': 'S', 'Izin': 'I', 'Tidak Hadir': 'TK' };
      const tidakHadirList = stat.filter(x => kode[x.st]).map(x => ({ nama: x.s.nama, kode: kode[x.st] }));

      const planned = info.planned || 1;
      const completed = Math.min(planned, info.selesai + (sudahAbsen && info.hari_ke ? 1 : 0));

      return {
        id_kelas: String(k._row),
        group: k.group || '',
        wilayah: k.cabang,
        grup_key: kunciTab_(k.cabang),
        nama_kelas: k.kelas,
        batch: batches.join(', '),
        start_date: info.start,
        planned_days: info.planned,
        completed_days: completed,
        progress: Math.round((completed / planned) * 100),
        total_siswa: siswa.length,
        sudah_absen: sudahAbsen,
        hadir: hitung('Hadir'),
        sakit: hitung('Sakit'),
        izin: hitung('Izin'),
        tidak_hadir: hitung('Tidak Hadir'),
        tidak_hadir_list: tidakHadirList,
        hari_ke: info.hari_ke,
        libur_ket: info.libur_ket,
        materi_status: info.status,
        materi: info.items,
        catatan: catatan
      };
    });
  }

  const diagnosa = result.length ? [] : readSheet_(SHEET_NAMES.MULAI).map(k =>
    'cabang: ' + k.cabang + ' | kelas: ' + k.kelas + ' | group: ' + k.group +
    ' | start: ' + (fmtDate_(k.start_date) || '-') + ' | end: ' + (fmtDate_(k.end_date) || '-'));

  return {
    success: true, today: today, sesi: sesi, kelas: result, diagnosa: diagnosa,
    admin: cak.admin, cabang_list: cak.admin ? CABANG_LIST : [],
    sensei_login: sensei.id_sensei + ' | group: ' + sensei.group + ' | wilayah: ' + sensei.wilayah_utama
  };
}

// Daftar siswa + status sesi itu pada tanggal itu + materi sesuai pertemuan ke-berapa
function getSiswaAbsen(idKelas, tanggal, sesi) {
  if (SESI.indexOf(sesi) < 0) sesi = sesiSekarang_();
  const k = readSheet_(SHEET_NAMES.MULAI).find(x => String(x._row) === String(idKelas));
  if (!k) return { siswa: [], materi: '', sudah_ada: false, terisi: {}, alasan: 'Kelas tidak ditemukan di sheet Mulai kelas.' };

  const absen = bacaAbsenTab_(k.cabang);
  // hanya batch yang sedang berjalan (batch lama sudah lewat, tidak diabsen lagi)
  const semuaSiswa = siswaKelas_(absen, k);
  const batchAktif = batchBerjalan_(k, semuaSiswa);
  const siswa = semuaSiswa.filter(s => !batchAktif.length || batchAktif.indexOf(s.batch) >= 0);
  const info = infoHari_(k, kurikulumKelas_(k), getLiburMap_(), tanggal);

  const out = siswa.map(s => ({
    id_siswa: String(s.row), nama_siswa: s.nama, batch: s.batch, status: statusSiswa_(absen, s, tanggal, sesi)
  }));
  // terisi[batch][sesi] = true kalau sesi itu sudah ada isinya (untuk tanda ✓)
  const terisi = {};
  siswa.forEach(s => {
    const key = s.batch || '(tanpa batch)';
    terisi[key] = terisi[key] || {};
    SESI.forEach(x => { if (statusSiswa_(absen, s, tanggal, x)) terisi[key][x] = true; });
  });
  const ket = { 'belum_mulai': 'Kelas belum dimulai pada tanggal ini.', 'selesai': 'Kurikulum sudah selesai.' };
  return {
    libur: info.status === 'libur' ? (info.libur_ket || 'Libur') : '',
    alasan: siswa.length ? '' : 'Tidak ada siswa untuk kelas "' + k.kelas + '" di tab "' + kunciTab_(k.cabang) + ' (absen)".',
    siswa: out,
    hari_ke: info.hari_ke,
    materi: info.items.length ? 'Hari ke-' + info.hari_ke + '\n' + info.items.map(x => x.sesi + ': ' + x.materi).join('\n') : (ket[info.status] || ''),
    sudah_ada: out.some(x => x.status !== ''),
    terisi: terisi
  };
}

// Simpan absensi satu sesi. Nilai ditulis di blok batch masing-masing: 午前 di kolom kiri pasangan tanggal (G), 午後 di kolom kanan (H).
// Pasangan kolom untuk tanggal baru dibuat otomatis (tanggal + label sesi di 2 baris header di atas blok).
// Menyimpan ulang sesi yang sama menimpa nilai siswa yang dikirim saja.
function simpanAbsensi(payload) {
  if (!payload || !payload.id_kelas || !payload.tanggal || !payload.dataSiswa || !payload.dataSiswa.length) {
    throw new Error('Data absensi tidak lengkap');
  }
  if (SESI.indexOf(payload.sesi) < 0) throw new Error('Sesi tidak valid');
  if (payload.dataSiswa.some(x => !KODE_ABSEN[x.status])) throw new Error('Ada siswa yang belum dipilih statusnya');
  const pengguna = ambilAkun_(payload.id_sensei);
  if (isAdmin_(pengguna)) throw new Error('Akun Admin hanya bisa memantau progress, tidak bisa menyimpan absensi.');

  const k = readSheet_(SHEET_NAMES.MULAI).find(x => String(x._row) === String(payload.id_kelas));
  if (!k) throw new Error('Kelas tidak ditemukan di sheet Mulai kelas');

  const dow = dayOfWeek_(payload.tanggal);
  const libur = getLiburMap_()[payload.tanggal] || ((dow === 0 || dow === 6) ? 'Akhir pekan' : '');
  if (libur) throw new Error('Tanggal ' + payload.tanggal + ' libur (' + libur + '). Absensi tidak disimpan.');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const absen = bacaAbsenTab_(k.cabang);
    const sheet = absen.sheet;
    const barisKelas = {};
    siswaKelas_(absen, k).forEach(s => { barisKelas[s.row] = s; });

    // pastikan baris yang dikirim memang siswa kelas ini dan namanya belum bergeser
    const perBlok = {};
    payload.dataSiswa.forEach(sw => {
      const s = barisKelas[sw.id_siswa];
      if (!s || (sw.nama && s.nama !== sw.nama)) {
        throw new Error('Daftar siswa di spreadsheet berubah. Muat ulang halaman absen lalu coba lagi.');
      }
      (perBlok[s.blok] = perBlok[s.blok] || []).push({ s: s, kode: KODE_ABSEN[sw.status] });
    });

    // Dari blok paling bawah dulu: menyisipkan baris di bawah tidak menggeser blok di atasnya
    Object.keys(perBlok).map(Number).sort((a, b) => b - a).forEach(bi => {
      const b = absen.blok[bi];
      let f = b.f, hdr = b.hdr;
      if (!hdr) {                        // belum ada baris header -> sisipkan 2 baris tepat di atas blok
        sheet.insertRowsBefore(f, 2);
        hdr = { d: f, s: f + 1 };
        f += 2;
      }

      // cari pasangan kolom untuk tanggal ini di baris tanggal blok
      const lebar = Math.max(sheet.getLastColumn(), KOLOM_TANGGAL_AWAL + 1);
      const baris = sheet.getRange(hdr.d, 1, 1, lebar).getValues()[0];
      let left = -1, kosong = -1, c = KOLOM_TANGGAL_AWAL - 1;
      for (; c < lebar; c += 2) {
        if (fmtDate_(baris[c], absen.tz) === payload.tanggal) { left = c; break; }
        if (kosong < 0 && String(baris[c] === undefined ? '' : baris[c]).trim() === '') kosong = c;
      }
      const baru = left < 0;
      if (baru) left = kosong >= 0 ? kosong : c;

      const perlu = left + 2;
      if (perlu > sheet.getMaxColumns()) sheet.insertColumnsAfter(sheet.getMaxColumns(), perlu - sheet.getMaxColumns());
      const colL = left + 1;

      if (baru) {
        sheet.getRange(hdr.d, colL).setValue(payload.tanggal).setNumberFormat('yyyy/M/d');
        sheet.getRange(hdr.d, colL, 1, 2).merge().setHorizontalAlignment('center');
        if (hdr.s > 0) sheet.getRange(hdr.s, colL, 1, 2).setValues([[SESI[0], SESI[1]]]).setHorizontalAlignment('center');
      }

      const col = colL + (payload.sesi === SESI[1] ? 1 : 0);
      const rng = sheet.getRange(f, col, b.last - b.f + 1, 1);
      const vals = rng.getValues();
      perBlok[bi].forEach(x => { vals[x.s.row - b.f][0] = x.kode; });
      rng.setValues(vals);
    });

    const batches = [...new Set(payload.dataSiswa.map(sw => barisKelas[sw.id_siswa].batch).filter(Boolean))].join(', ');
    return { success: true, message: 'Absensi ' + payload.sesi + ' · ' + payload.tanggal + (batches ? ' · ' + batches : '') + ' (' + payload.dataSiswa.length + ' siswa) berhasil disimpan' };
  } finally {
    lock.releaseLock();
  }
}

// ====================== REKAP KEHADIRAN ======================
// Dihitung per hari dari kedua sesi (hanya hari yang 午前 dan 午後-nya sudah terisi):
//   H  = hadir di 午前 dan 午後                         dihitung 1 kehadiran
//   I / S / TK = tidak hadir di kedua sesi (alasan sama)  tidak dihitung hadir
//   前半 = tidak hadir di 午前 tetapi hadir di 午後        dihitung 0,5 kehadiran
//   後半 = hadir di 午前 tetapi tidak hadir di 午後        dihitung 0,5 kehadiran
//   Total (kehadiran) = H + 0,5 x (前半 + 後半)   -> I, S, TK tidak ikut dihitung
// Kalau kedua sesi tidak hadir dengan alasan berbeda (mis. 午前 Sakit, 午後 Izin), dihitung sesuai alasan 午前.
function kunciSiswa_(nama, kelas, batch) {
  return [nama, kelas, batch].map(x => String(x || '').replace(/\s+/g, ' ').trim().toLowerCase()).join('|');
}

function rekapSiswa_(absen, s) {
  const r = { h: 0, i: 0, s: 0, tk: 0, depan: 0, belakang: 0, hari: 0 };
  const kode = { 'Izin': 'i', 'Sakit': 's', 'Tidak Hadir': 'tk' };
  Object.keys(absen.blok[s.blok].pairs).forEach(tgl => {
    const a = statusSiswa_(absen, s, tgl, SESI[0]);
    const p = statusSiswa_(absen, s, tgl, SESI[1]);
    if (!a || !p) return;                       // hari belum lengkap
    r.hari++;
    if (a === 'Hadir' && p === 'Hadir') r.h++;
    else if (a !== 'Hadir' && p === 'Hadir') r.depan++;
    else if (a === 'Hadir' && p !== 'Hadir') r.belakang++;
    else r[kode[a]]++;
  });
  r.total = r.h + 0.5 * (r.depan + r.belakang);
  return r;
}

// Rekap semua siswa di tab absen cabang sensei (cabang diambil dari akun, bukan dari browser)
function getRekapAbsen(idSensei, pilih) {
  const sensei = ambilAkun_(idSensei);
  const cak = cakupan_(sensei, pilih);
  const rekap = [];
  cak.cabang.forEach(c => {
    let absen;
    try { absen = bacaAbsenTab_(c); } catch (e) { if (!cak.admin) throw e; return; }
    absen.siswa.forEach(s => {
      const r = rekapSiswa_(absen, s);
      r.key = kunciSiswa_(s.nama, kelasKunci_(s.kelas, c, cak.gabung), s.batch);
      rekap.push(r);
    });
  });
  return { success: true, rekap: rekap };
}


// ====================== KURIKULUM (menu Kurikulum) ======================
// Tanggal pelaksanaan hari ke-1..planned dari `start`; Sabtu, Minggu, dan HariLibur dilewati
// (aturan yang sama dengan hitungJadwal_, jadi sinkron dengan halaman Home).
function tanggalHari_(start, planned, liburMap) {
  const out = {};
  let d = start, n = 0;
  for (let i = 0; i < 1500 && n < planned; i++) {
    const dow = dayOfWeek_(d);
    if (!liburMap[d] && dow !== 0 && dow !== 6) { n++; out[n] = d; }
    d = addDays_(d, 1);
  }
  return out;
}

// Isi Database Kurikulum (tab cabang sensei) + jadwal tiap kelas sensei (tanggal tiap hari, hari ini hari ke berapa).
// Cabang diambil dari akun sensei di server. Pengajar & catatan (指摘) dibaca dari kolom batch kelas masing-masing.
function getKurikulumData(idSensei, pilih) {
  const sensei = ambilAkun_(idSensei);
  const cak = cakupan_(sensei, pilih);

  const today = todayStr_();
  const liburMap = getLiburMap_();

  // Kurikulum acuan (isi tabel): tab cabang sensei. Admin: tab pertama yang ada di cabang terpilih.
  let kd = null, cabangAcuan = '';
  for (let i = 0; i < cak.cabang.length && !kd; i++) {
    try { kd = bukaKurikulum_(cak.cabang[i]); cabangAcuan = cak.cabang[i]; }
    catch (e) { if (!cak.admin) throw e; }
  }
  if (!kd) throw new Error('Tab kurikulum untuk cabang yang dipilih tidak ditemukan di Database Kurikulum');
  const absenDari = pembacaAbsen_(cak);

  const salin = hari => hari.map(h => ({ hari_ke: h.hari_ke, items: h.items, pr: '' }));
  const utama = kd.blok[0] ? kd.blok[0].hari.filter(h => h.hari_ke <= UTAMA_HARI) : [];
  const grup = {};
  GRUP_KURIKULUM.forEach(g => {
    const b = kd.blok.filter(x => x.key === g.toLowerCase())[0];
    grup[g.toLowerCase()] = { nama: g, sheet: kd.tab, ada: !!b, hari: b ? salin(b.hari) : [] };
  });
  const maksHari = Math.max.apply(null, [utama.length].concat(kd.blok.map(b => b.hari.length)));

  const kelas = kelasSensei_(sensei, cak).map(k => {
    const kurK = cak.admin ? kurikulumKelasAman_(k) : kurikulumKelas_(k);
    const info = infoHari_(k, kurK, liburMap, today);
    const batches = batchBerjalan_(k, siswaKelas_(absenDari(k.cabang), k));
    const b = batches.join(', ');
    return {
      id_kelas: String(k._row),
      group: String(k.group || '').trim(),
      nama_kelas: String(k.kelas || '').trim(),
      batch: b,
      label: [k.kelas, b, k.group, cak.gabung ? k.cabang : ''].map(x => String(x || '').trim()).filter(Boolean).join(' · '),
      start: info.start,
      hari_ke: info.hari_ke,
      status: info.status,
      libur_ket: info.libur_ket,
      selesai: info.selesai,
      tanggal: info.start ? tanggalHari_(info.start, maksHari + KUR_GRUP_OFFSET, liburMap) : {},
      catatan: catatanBatch_(kurK.kd, kurK.blok, noBatch_(batches[0]))
    };
  });

  return {
    success: true,
    today: today,
    cabang: cak.gabung ? 'Semua cabang' : cak.cabang[0],
    tab_cabang: cabangAcuan,
    planned: utama.length,
    hari: salin(utama),
    kelas: kelas,
    grup: grup,
    offset: KUR_GRUP_OFFSET
  };
}

// Kurikulum group dimulai pada hari belajar ke-(OFFSET + 1) kelas.
// 0 = hari ke-N di Database Kurikulum sama dengan hari belajar ke-N kelas (blok group memuat seluruh hari, dari hari 1).
const KUR_GRUP_OFFSET = 0;

// Simpan catatan (指摘) hari tertentu untuk kelas sensei: 午前 dan 午後 ditulis di kolom 指摘 batch kelas itu,
// pada baris hari tersebut di blok group-nya (mis. Batch 4 hari ke-64 -> K129:K130).
// Hanya boleh untuk kelas milik akun sensei, dan hanya sampai hari yang sudah berjalan.
function simpanCatatan(p) {
  p = p || {};
  const sensei = ambilAkun_(p.id_sensei);
  if (isAdmin_(sensei)) throw new Error('Akun Admin hanya bisa memantau progress, tidak bisa menyimpan catatan.');
  const k = kelasSensei_(sensei, cakupan_(sensei, '')).filter(x => String(x._row) === String(p.id_kelas))[0];
  if (!k) throw new Error('Kelas tidak ditemukan untuk akun ini');
  const hari = Number(p.hari_ke);
  if (!hari || hari < 1 || hari !== Math.floor(hari)) throw new Error('Hari ke- tidak valid');

  const absen = bacaAbsenTab_(k.cabang);
  const nb = noBatch_(batchBerjalan_(k, siswaKelas_(absen, k))[0]);
  if (!nb) throw new Error('Nomor batch kelas ini tidak terbaca');

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const kd = bukaKurikulum_(k.cabang, true);                       // baca ulang supaya posisi baris terbaru
    const blok = kd.blok.filter(b => b.key === norm_(k.group))[0];
    if (!blok) throw new Error('Kurikulum group "' + k.group + '" belum ada di tab ' + kd.tab + ' Database Kurikulum');
    const bc = blok.batch[nb];
    if (!bc) throw new Error('Kolom Batch ' + nb + ' tidak ditemukan di blok ' + blok.grup);
    const h = blok.hari.filter(x => x.hari_ke === hari)[0];
    if (!h) throw new Error('Hari ke-' + hari + ' tidak ada di kurikulum ' + blok.grup);

    const info = infoHari_(k, { hari: blok.hari, kolomTanggal: {} }, getLiburMap_(), todayStr_());
    const maks = info.hari_ke || info.selesai;
    if (hari > maks) throw new Error('Hari ke-' + hari + ' belum berjalan. Catatan hanya bisa diisi sampai hari ke-' + maks + '.');

    const tulis = (sesi, teks) => {
      if (teks === undefined || teks === null || !h.rows[sesi]) return;
      let t = String(teks).replace(/\r/g, '').trim().slice(0, 2000);
      if (/^[=+@]/.test(t)) t = "'" + t;                              // cegah teks dibaca sebagai rumus
      kd.sheet.getRange(h.rows[sesi], bc.c + 1).setValue(t);
    };
    tulis(SESI[0], p.pagi);
    tulis(SESI[1], p.siang);
    SpreadsheetApp.flush();
    KUR_CACHE_ = {};
    return { success: true, message: 'Catatan hari ke-' + hari + ' · Batch ' + nb + ' berhasil disimpan' };
  } finally {
    lock.releaseLock();
  }
}

// ====================== DIAGNOSA KURIKULUM ======================
// Jalankan manual dari editor Apps Script (pilih cekKurikulum -> Run -> lihat Execution log).
// Menunjukkan blok group, jumlah hari, dan kolom batch yang terbaca di tiap tab cabang Database Kurikulum.
function cekKurikulum() {
  const db = SpreadsheetApp.openById(ID_DB_KURIKULUM);
  db.getSheets().forEach(sh => {
    let kd;
    try { kd = bukaKurikulum_(sh.getName(), true); } catch (e) { Logger.log(sh.getName() + ': ' + e.message); return; }
    Logger.log(sh.getName() + ': ' + kd.blok.length + ' blok');
    kd.blok.forEach(b => Logger.log('  ' + b.grup + ': ' + b.hari.length + ' hari (hari ke-' + (b.hari[0] ? b.hari[0].hari_ke : '-') + ' s/d ' +
      (b.hari.length ? b.hari[b.hari.length - 1].hari_ke : '-') + '); batch = ' +
      Object.keys(b.batch).map(n => 'Batch ' + n + ' (kolom ' + (b.batch[n].p + 1) + '-' + (b.batch[n].c + 1) + ')').join(', ')));
  });
}

// ====================== SERTIFIKASI (menu Kelas) ======================
// Spreadsheet "Database Sertifikasi": satu sheet per group (Joybridge, Joycare).
// Kolom dibaca lewat JUDUL (baris 1-3), bukan posisi tetap, jadi kalau ada kolom ujian baru
// (mis. SSW 4x, ujian baru) cukup tambah di spreadsheet dengan pola judul yang sama:
//   baris judul : nama ujian (JFT/JLPT, 土木, 設備, 建築, SSW JP, SSW ID, ...) | 受験日・合格日 (Tanggal) | 面接採用状況
//   baris 1x/2x : tiap percobaan ujian. Isi sel: ✓ = lulus, X/✕ = belum lulus, kosong = belum ikut
//   kolom tanggal berada tepat setelah kolom percobaan ujian tersebut.
const ID_DB_SERTIFIKASI = '1KPmMfE14ucCFNTbJ6sDGHuv3g1CW3O-Wf1ezvO7FsAg';
const SHEET_SERTIFIKAT = { 'joybridge': 'Joybridge', 'joycare': 'Joycare' };   // kunci = nama group (huruf kecil)
const KET_UJIAN = { '土木': 'Teknik sipil', '設備': 'Instalasi / fasilitas', '建築': 'Bangunan' };
const LULUS_RE = /^[✓✔☑√]|[★☆⭐]/;      // ✓ atau ✓★ (bintang saja juga dianggap lulus)
const BINTANG_RE = /[★☆⭐]/;               // bintang = lulus dengan nilai 満点 (sempurna)
const GAGAL_RE = /^[xX✕✗×]$/;

// "M. Reihansyah  Rifqi" -> "m reihansyah rifqi" (untuk mencocokkan nama siswa antar spreadsheet)
function normNama_(s) {
  return String(s || '').toLowerCase().replace(/[.,'’`"]/g, '').replace(/\s+/g, ' ').trim();
}

function labelUjian_(l) {
  l = String(l).replace(/\s+/g, ' ').trim();
  if (/jft|jlpt/i.test(l)) return 'JFT';
  if (KET_UJIAN[l]) return 'SSW ' + l;
  return l;
}

function bacaSertifikatSheet_(db, namaSheet) {
  const sheet = db.getSheets().filter(sh => norm_(sh.getName()) === norm_(namaSheet))[0];
  if (!sheet) throw new Error('Sheet "' + namaSheet + '" tidak ditemukan di Database Sertifikasi');
  const v = sheet.getDataRange().getDisplayValues();      // tampilan apa adanya (tanggal tetap dd/mm/yyyy)
  const cell = (r, c) => String((v[r] || [])[c] === undefined ? '' : v[r][c]).trim();

  // baris "1x 2x 3x" = baris percobaan; baris di atasnya = nama ujian / tanggal / status
  let rAtt = -1;
  for (let r = 0; r < Math.min(8, v.length); r++) {
    if ((v[r] || []).filter(x => /^\d+x$/i.test(String(x).trim())).length >= 2) { rAtt = r; break; }
  }
  if (rAtt < 1) throw new Error('Sheet "' + namaSheet + '": baris judul percobaan (1x, 2x, ...) tidak ditemukan');
  const rLabel = rAtt - 1;
  const lebar = Math.max.apply(null, v.map(r => r.length));

  const cari = nama => { for (let r = 0; r <= rAtt; r++) for (let c = 0; c < lebar; c++) if (cell(r, c).toLowerCase() === nama) return c; return -1; };
  const cNama = cari('nama'), cKelas = cari('kelas'), cGender = cari('gender');
  if (cNama < 0) throw new Error('Sheet "' + namaSheet + '" harus punya judul kolom Nama');

  const ujian = [];
  let cStatus = -1, cur = null;
  for (let c = Math.max(cNama, cKelas, cGender) + 1; c < lebar; c++) {
    const l2 = cell(rLabel, c), l3 = cell(rAtt, c);
    if (l2) {
      if (/受験日|tanggal/i.test(l2)) { if (cur) cur.cTgl = c; }
      else if (/面接|採用|penerimaan/i.test(l2)) { cStatus = c; cur = null; }
      else { cur = { label: labelUjian_(l2), ket: KET_UJIAN[l2.replace(/\s+/g, ' ').trim()] || '', cAtt: [], cTgl: -1 }; ujian.push(cur); }
    }
    if (cur && /^\d+x$/i.test(l3)) cur.cAtt.push(c);
  }

  const kolom = ujian.map((u, i) => ({ key: 'u' + i, label: u.label, ket: u.ket }));
  const baris = [];
  for (let r = rAtt + 1; r < v.length; r++) {
    const nama = cell(r, cNama);
    if (!nama || nama.toLowerCase() === 'nama') continue;
    const hasil = {};
    let jumlah = 0;
    ujian.forEach((u, i) => {
      const att = u.cAtt.map(c => cell(r, c));
      const idx = att.findIndex(a => LULUS_RE.test(a));
      const gagal = att.filter(a => GAGAL_RE.test(a)).length;
      const lulus = idx >= 0;
      const sempurna = lulus && BINTANG_RE.test(att[idx]);
      if (lulus) jumlah++;
      hasil['u' + i] = {
        status: lulus ? 'lulus' : (gagal ? 'gagal' : 'belum'),
        percobaan: lulus ? idx + 1 : 0,
        sempurna: sempurna,
        gagal: gagal,
        tanggal: u.cTgl >= 0 ? cell(r, u.cTgl) : ''
      };
    });
    baris.push({
      nama: nama,
      kelas: cKelas >= 0 ? cell(r, cKelas) : '',
      gender: cGender >= 0 ? cell(r, cGender) : '',
      ujian: hasil,
      jumlah: jumlah,
      status: cStatus >= 0 ? cell(r, cStatus) : ''
    });
  }
  return { kolom: kolom, baris: baris };
}

// Sertifikasi siswa di cabang sensei. Hanya siswa yang ada di Database Kelas cabang itu yang dikirim ke browser.
// Siswa dicocokkan lewat NAMA, ke sheet sesuai group-nya (Joybridge -> sheet Joybridge, Joycare -> sheet Joycare).
function getSertifikasi(idSensei, pilih) {
  const kelas = getDataKelas(idSensei, pilih);          // sekaligus memvalidasi akun & menentukan cabang dari server
  const butuh = {};
  kelas.siswa.forEach(s => {
    const g = norm_(s.group);
    if (SHEET_SERTIFIKAT[g]) (butuh[g] = butuh[g] || {})[normNama_(s.nama)] = 1;
  });

  let db;
  try {
    db = SpreadsheetApp.openById(ID_DB_SERTIFIKASI);
  } catch (e) {
    throw new Error('Spreadsheet Database Sertifikasi tidak bisa dibuka. Pastikan akun pemilik script punya akses ke spreadsheet tersebut.');
  }
  const grup = {};
  Object.keys(butuh).forEach(g => {
    const t = bacaSertifikatSheet_(db, SHEET_SERTIFIKAT[g]);
    const siswa = {};
    t.baris.forEach(b => { const k = normNama_(b.nama); if (butuh[g][k] && !siswa[k]) siswa[k] = b; });
    grup[g] = { sheet: SHEET_SERTIFIKAT[g], kolom: t.kolom, siswa: siswa };
  });
  return { success: true, grup: grup };
}

// Diagnosa manual: pilih cekSertifikasi -> Run -> lihat Execution log. Menunjukkan kolom yang terbaca & siswa pertama.
function cekSertifikasi() {
  const db = SpreadsheetApp.openById(ID_DB_SERTIFIKASI);
  Object.keys(SHEET_SERTIFIKAT).forEach(g => {
    const t = bacaSertifikatSheet_(db, SHEET_SERTIFIKAT[g]);
    Logger.log(SHEET_SERTIFIKAT[g] + ': ' + t.baris.length + ' siswa; ujian = ' + t.kolom.map(k => k.label).join(', '));
    t.baris.slice(0, 3).forEach(b => Logger.log('  ' + JSON.stringify(b)));
  });
}

// ====================== TIMELINE (menu Timeline) ======================
// Semua kelas di sheet "Mulai kelas" (semua cabang). Tiap kelas = 3 fase berurutan, dihitung dari Start_date
// dengan aturan hari kerja yang sama dengan Home/Kurikulum (Sabtu, Minggu, dan HariLibur dilewati):
//   Reguler 98 hari  ->  SSW 30 hari  ->  Mensetsu 10 hari
// Kalau durasi berubah, cukup ubah tiga angka ini.
const TL_HARI = { REGULER: 98, SSW: 30, MENSETSU: 10 };

function getTimelineData(idSensei) {
  const sensei = readSheet_(SHEET_NAMES.SENSEI).find(r => norm_(r.id_sensei) === norm_(idSensei));
  if (!sensei || !isAktif_(sensei.status)) throw new Error('Akun sensei tidak ditemukan atau tidak aktif');

  const today = todayStr_();
  const libur = getLiburMap_();
  const t1 = TL_HARI.REGULER, t2 = t1 + TL_HARI.SSW, t3 = t2 + TL_HARI.MENSETSU;
  const kelas = [];

  readSheet_(SHEET_NAMES.MULAI).forEach(k => {
    const nama = String(k.kelas || '').trim();
    const start = fmtDate_(k.start_date);
    if (!nama || !start) return;                       // tanpa nama kelas / Start_date tidak bisa digambar

    const tgl = tanggalHari_(start, t3, libur);        // {1: tanggal hari kerja ke-1, ..., t3: ...}
    const d = n => tgl[n] || '';
    if (!d(t3)) return;

    // posisi hari ini: sudah berapa hari kerja terlewati
    const n = Object.keys(tgl).filter(x => tgl[x] <= today).length;
    let fase = 'selesai', hari = 0, dari = 0;
    if (today < d(1)) fase = 'belum';
    else if (today > d(t3)) fase = 'selesai';
    else if (n <= t1) { fase = 'reguler'; hari = n; dari = TL_HARI.REGULER; }
    else if (n <= t2) { fase = 'ssw'; hari = n - t1; dari = TL_HARI.SSW; }
    else { fase = 'mensetsu'; hari = n - t2; dari = TL_HARI.MENSETSU; }

    kelas.push({
      id: String(k._row),
      cabang: String(k.cabang || '').trim(),
      kelas: nama,
      group: String(k.group || '').trim(),
      batch: String(k.batch || '').trim(),
      start: d(1),
      end_sheet: fmtDate_(k.end_date),
      fase: {
        reguler:  { mulai: d(1),      selesai: d(t1) },
        ssw:      { mulai: d(t1 + 1), selesai: d(t2) },
        mensetsu: { mulai: d(t2 + 1), selesai: d(t3) }
      },
      status: { fase: fase, hari: hari, dari: dari }
    });
  });

  return { success: true, today: today, durasi: TL_HARI, cabang_saya: String(sensei.wilayah_utama || ''), kelas: kelas };
}

// ====================== REKAP NILAI (menu Rekap Nilai) ======================
// Spreadsheet "Database Nilai": satu tab per cabang (Tangerang, Bandung, ...), nama tab = nama cabang.
// Satu tab boleh berisi beberapa blok tabel BERTUMPUK (mis. blok Joycare, lalu baris pemisah, lalu blok Joybridge),
// dan tiap blok boleh punya kolom nilai yang berbeda. Karena itu kolom TIDAK ditulis tetap di kode, tapi dibaca dari judul:
//   - Baris judul = baris yang punya sel "Nama" (kolom identitas: No, Group, Nama, Batch, Kelas).
//   - Judul bisa 2 tingkat (mis. "Test N4 (1)" yang di-merge, lalu Moji Goi / Dokkai / Choukai). Judul tingkat atas boleh
//     di atas atau di bawah baris "Nama"; sel merge dibaca dari struktur merge sheet.
//   - Angka di baris sub-judul (mis. 44 di bawah 施設の言葉) dianggap SKOR MAKSIMAL kolom itu.
//   - Group dibaca dari kolom Group tiap siswa, jadi group baru (Joycare, Joybridge, Joygreen, Joyserve, atau lainnya) muncul otomatis.
const ID_DB_NILAI = '1-GGdvwNGd1lW5Gp8RrbSHakjrsJLuWHIOIovH2tUVDM';
const GRUP_NILAI = ['Joycare', 'Joybridge', 'Joygreen', 'Joyserve'];   // hanya untuk urutan tampil; group lain tetap terdeteksi
const KOLOM_IDENTITAS_NILAI = ['no', 'group', 'nama', 'batch', 'kelas'];

function bacaNilaiTab_(sheet) {
  const out = { blok: [], grup: [] };
  const rng = sheet.getDataRange();
  const v = rng.getDisplayValues();
  if (v.length < 2) return out;

  const T = x => String(x == null ? '' : x).replace(/\s+/g, ' ').trim();
  const LOW = x => T(x).toLowerCase();
  const ANGKA = /^-?\d+([.,]\d+)?$/;
  const merged = rng.getMergedRanges();

  // 1) baris judul: setiap baris yang punya sel "Nama"
  const hdr = [];
  for (let r = 0; r < v.length; r++) {
    const cNama = v[r].findIndex(c => LOW(c) === 'nama');
    if (cNama < 0) continue;
    const gi = v[r].findIndex(c => LOW(c) === 'group');
    const cGroup = gi >= 0 ? gi : 1;                      // kolom B kalau tidak ada judul "Group"
    // baris judul tambahan = Nama & Group kosong tapi ada isi lain
    const barisJudul = rr => rr >= 0 && rr < v.length && !T(v[rr][cNama]) && !T(v[rr][cGroup]) && v[rr].some(c => T(c));
    let top = r, bottom = r;
    if (barisJudul(r - 1)) top = r - 1;                   // judul tingkat atas di atas baris "Nama"
    else if (barisJudul(r + 1)) bottom = r + 1;           // sub-judul di bawah baris "Nama"
    hdr.push({ r: r, top: top, bottom: bottom, cNama: cNama, cGroup: cGroup });
  }

  // 2) kolom & siswa tiap blok
  const hitung = {}, namaGrup = {};
  hdr.forEach((h, hi) => {
    const ident = {};
    v[h.r].forEach((c, i) => { const l = LOW(c); if (KOLOM_IDENTITAS_NILAI.indexOf(l) >= 0 && ident[l] === undefined) ident[l] = i; });
    const kolomIdent = Object.keys(ident).map(k => ident[k]);

    // judul tingkat atas yang di-merge: kolom -> teks judul
    const parentMerge = {};
    let adaMerge = false;
    merged.forEach(m => {
      const r0 = m.getRow() - 1, c0 = m.getColumn() - 1;
      if (m.getNumColumns() < 2 || r0 > h.top || r0 + m.getNumRows() - 1 < h.top) return;
      adaMerge = true;
      const lab = T((v[r0] || [])[c0]);
      for (let k = 0; k < m.getNumColumns(); k++) parentMerge[c0 + k] = lab;
    });

    // siswa: dari baris sesudah judul sampai judul blok berikutnya
    const akhir = hdr[hi + 1] ? hdr[hi + 1].top : v.length;
    const adaData = c => { for (let rr = h.bottom + 1; rr < akhir; rr++) if (T(v[rr][h.cNama]) && T(v[rr][c])) return true; return false; };
    const hurufKol = i => { let s = '', n = i + 1; while (n > 0) { const m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); } return s; };
    // judul cadangan: teks di baris tepat di atas baris judul (kalau itu bukan baris siswa); kalau tidak ada -> "Kolom O"
    const labelAtas = c => { const rr = h.top - 1; if (rr < 0 || T(v[rr][h.cNama])) return ''; const t = T(v[rr][c]); return (t && !ANGKA.test(t)) ? t : ''; };
    const nomor = t => parseFloat(String(t).replace(',', '.'));

    const kolom = [];
    let carry = '';                                       // cadangan kalau judul atas tidak di-merge
    for (let c = 0; c < v[h.r].length; c++) {
      if (kolomIdent.indexOf(c) >= 0 || c === h.cGroup || c === h.cNama) { carry = ''; continue; }
      const a = T(v[h.top][c]);
      const b = h.bottom !== h.top ? T(v[h.bottom][c]) : '';
      let auto = false;                                   // true = judul tidak ada di sheet (dipakai nama huruf kolom)
      const cadangan = () => { const t = labelAtas(c); if (t) return t; auto = true; return 'Kolom ' + hurufKol(c); };
      let parent = '', leaf = '', maks = null;
      if (h.bottom === h.top) {                           // satu baris judul
        if (ANGKA.test(a)) { maks = nomor(a); leaf = cadangan(); }       // angka di baris judul = jumlah soal (skor maksimal)
        else if (a) leaf = a;
        else if (adaData(c)) leaf = cadangan();                          // kolom berisi nilai tapi judulnya kosong
        else { carry = ''; continue; }
      } else if (a && b) {
        if (ANGKA.test(b)) { leaf = a; maks = nomor(b); carry = ''; }
        else if (ANGKA.test(a)) { maks = nomor(a); leaf = b; carry = ''; }
        else { parent = a; leaf = b; carry = a; }
      } else if (a) {
        carry = '';
        if (ANGKA.test(a)) { maks = nomor(a); leaf = cadangan(); } else leaf = a;
      } else if (b) {
        if (ANGKA.test(b)) { carry = ''; continue; }
        leaf = b;
        parent = parentMerge[c] !== undefined ? parentMerge[c] : (!adaMerge ? carry : '');
      } else if (adaData(c)) { leaf = cadangan(); carry = ''; }
      else { carry = ''; continue; }
      if (!leaf) continue;
      kolom.push({ i: c, label: parent ? parent + ' · ' + leaf : leaf, parent: parent, leaf: leaf, maks: maks, auto: auto });
    }

    const baris = [];
    for (let rr = h.bottom + 1; rr < akhir; rr++) {
      const nama = T(v[rr][h.cNama]);
      if (!nama) continue;
      const group = T(v[rr][h.cGroup]);
      if (group) { const g = group.toLowerCase(); hitung[g] = (hitung[g] || 0) + 1; if (!namaGrup[g]) namaGrup[g] = group; }
      baris.push({
        nama: nama,
        group: group,
        batch: ident.batch !== undefined ? T(v[rr][ident.batch]) : '',
        nilai: kolom.map(k => T(v[rr][k.i]))
      });
    }
    out.blok.push({ kolom: kolom, baris: baris });
  });

  // Batch lain mengikuti judul batch 1: untuk tiap group, blok dengan judul paling lengkap (biasanya Batch 1) jadi acuan.
  // Judul, judul atas (parent), dan jumlah soal (maks) kolom yang SAMA POSISINYA di blok lain disalin dari acuan itu.
  const grupBlok = b => { const x = b.baris.filter(r => r.group)[0]; return x ? x.group.toLowerCase() : ''; };
  const skorBlok = b => b.kolom.filter(k => !k.auto).length;
  const acuan = {};
  out.blok.forEach(b => { const gk = grupBlok(b); if (gk && (!acuan[gk] || skorBlok(b) > skorBlok(acuan[gk]))) acuan[gk] = b; });
  out.blok.forEach(b => {
    const t = acuan[grupBlok(b)];
    if (!t || t === b) return;
    const peta = {};
    t.kolom.forEach(k => { if (!k.auto) peta[k.i] = k; });
    b.kolom.forEach(k => {
      const s = peta[k.i];
      if (!s) return;
      k.label = s.label; k.parent = s.parent; k.leaf = s.leaf; k.auto = false;
      if (k.maks == null) k.maks = s.maks;
    });
  });

  const urut = g => { const i = GRUP_NILAI.map(x => x.toLowerCase()).indexOf(g); return i < 0 ? 99 : i; };
  out.grup = Object.keys(hitung).sort((a, b) => urut(a) - urut(b) || a.localeCompare(b)).map(g => ({ nama: namaGrup[g], n: hitung[g] }));
  return out;
}

// Nilai siswa di cabang sensei. Cabang ditentukan server dari akun; siswa dicocokkan lewat NAMA ke Database Kelas
// (lebih disukai yang group & batch-nya sama). Yang dikirim ke browser hanya siswa yang ada di Database Kelas cabang itu.
function getRekapNilai(idSensei, pilih) {
  const cak = cakupan_(ambilAkun_(idSensei), pilih);
  const kelas = getDataKelas(idSensei, pilih);    // sekaligus memvalidasi akun & menentukan cabang dari server
  let db;
  try {
    db = SpreadsheetApp.openById(ID_DB_NILAI);
  } catch (e) {
    throw new Error('Spreadsheet Database Nilai tidak bisa dibuka. Pastikan akun pemilik script punya akses ke spreadsheet tersebut.');
  }

  // satu tab per cabang; blok semua tab digabung (indeks blok global), siswa dicari lewat "cabang|nama"
  const blok = [], idx = {}, tabs = [], hitungGrup = {};
  cak.cabang.forEach(c => {
    const key = kunciTab_(c);
    const sheet = db.getSheets().filter(sh => sh.getName().trim().toLowerCase() === key)[0];
    if (!sheet) return;
    tabs.push(sheet.getName());
    const t = bacaNilaiTab_(sheet);
    t.blok.forEach(b => {
      const bi = blok.length;
      blok.push({ kolom: b.kolom.map(k => ({ label: k.label, parent: k.parent, leaf: k.leaf, maks: k.maks, col: k.i })) });
      b.baris.forEach(r => { const kk = key + '|' + normNama_(r.nama); (idx[kk] = idx[kk] || []).push({ bi: bi, r: r }); });
    });
    t.grup.forEach(g => {
      const gk = g.nama.toLowerCase();
      if (!hitungGrup[gk]) hitungGrup[gk] = { nama: g.nama, n: 0 };
      hitungGrup[gk].n += g.n;
    });
  });
  if (!tabs.length) return { success: true, cabang: kelas.cabang, ada: false, tab: '', grup: [], blok: [], nilai: {} };

  const nilai = {};
  kelas.siswa.forEach(s => {
    const cand = idx[kunciTab_(s.cabang) + '|' + normNama_(s.nama)];
    if (!cand) return;
    const g = norm_(s.group), bt = norm_(s.batch);
    const pilihan = cand.filter(x => norm_(x.r.group) === g && norm_(x.r.batch) === bt)[0] ||
                    cand.filter(x => norm_(x.r.group) === g)[0] || cand[0];
    nilai[kunciSiswa_(s.nama, s.kelas, s.batch === '(tanpa batch)' ? '' : s.batch)] = { b: pilihan.bi, g: pilihan.r.group, v: pilihan.r.nilai };
  });

  const urutGrup = g => { const i = GRUP_NILAI.map(x => x.toLowerCase()).indexOf(g); return i < 0 ? 99 : i; };
  const grup = Object.keys(hitungGrup).sort((a, b) => urutGrup(a) - urutGrup(b) || a.localeCompare(b)).map(g => hitungGrup[g]);

  return {
    success: true,
    cabang: kelas.cabang,
    ada: true,
    tab: tabs.join(', '),
    grup: grup,
    blok: blok,
    nilai: nilai
  };
}

// Diagnosa manual: pilih cekNilai -> Run -> lihat Execution log. Menunjukkan blok, group, dan kolom yang terbaca di tiap tab cabang.
function cekNilai() {
  const db = SpreadsheetApp.openById(ID_DB_NILAI);
  db.getSheets().forEach(sh => {
    const t = bacaNilaiTab_(sh);
    Logger.log(sh.getName() + ': ' + t.blok.length + ' blok; group = ' + (t.grup.map(g => g.nama + ' (' + g.n + ')').join(', ') || '-'));
    t.blok.forEach((b, i) => Logger.log('  blok ' + (i + 1) + ': ' + b.baris.length + ' siswa; kolom = ' +
      b.kolom.map(k => k.label + (k.maks ? ' [maks ' + k.maks + ']' : '')).join(' | ')));
  });
}