/**
 * ============================================================
 *  Code.gs — BACKEND APLIKASI TO-DO LIST
 *  Google Apps Script + Google Sheets sebagai database
 * ============================================================
 *  Isi file ini:
 *   1. Konfigurasi (nama sheet, header, zona waktu)
 *   2. Web App (doGet, include) & menu di Spreadsheet
 *   3. setup()  -> membuat semua sheet & header otomatis
 *   4. Fungsi bantu (helper) tanggal & sheet
 *   5. CRUD Tugas (tambah, edit, hapus, ubah status)
 *   6. CRUD Kategori
 *   7. Pengaturan
 *   8. Trigger: buatTrigger() & hapusTrigger()
 *   9. Pengingat email: kirimPengingatHarian()
 * ============================================================
 */


/* ============================================================
 * 1. KONFIGURASI
 * ============================================================ */

// Zona waktu yang dipakai untuk semua perhitungan tanggal (WIB)
const ZONA_WAKTU = 'Asia/Jakarta';

// Nama-nama sheet
const SHEET_TUGAS = 'Tugas';
const SHEET_KATEGORI = 'Kategori';
const SHEET_PENGATURAN = 'Pengaturan';

// Header (baris pertama) setiap sheet
const HEADER_TUGAS = ['ID', 'Judul', 'Deskripsi', 'Kategori', 'Prioritas', 'Deadline', 'Status', 'Tanggal Dibuat', 'Tanggal Selesai'];
const HEADER_KATEGORI = ['Nama Kategori', 'Warna'];
const HEADER_PENGATURAN = ['Email Penerima', 'Jam Pengingat', 'Pengingat Aktif'];

// Kategori bawaan yang dibuat saat setup() pertama kali
const KATEGORI_AWAL = [
  ['Mengajar', '#2563eb'],
  ['Administrasi Sekolah', '#7c3aed'],
  ['Operator', '#0d9488'],
  ['Pribadi', '#db2777']
];

// Nilai status tugas
const STATUS_SELESAI = 'Selesai';
const STATUS_BELUM = 'Belum Selesai';

// Urutan prioritas (angka kecil = lebih penting)
const URUTAN_PRIORITAS = { 'Tinggi': 1, 'Sedang': 2, 'Rendah': 3 };

// Nama fungsi yang dijalankan oleh trigger harian
const FUNGSI_PENGINGAT = 'kirimPengingatHarian';

// Jam pengingat bawaan (6 = antara pukul 06.00–07.00 WIB)
const JAM_BAWAAN = 6;

// Nama hari & bulan dalam Bahasa Indonesia
const NAMA_HARI = ['Minggu', 'Senin', 'Selasa', 'Rabu', 'Kamis', 'Jumat', 'Sabtu'];
const NAMA_BULAN = ['Januari', 'Februari', 'Maret', 'April', 'Mei', 'Juni', 'Juli',
  'Agustus', 'September', 'Oktober', 'November', 'Desember'];


/* ============================================================
 * 2. WEB APP & MENU SPREADSHEET
 * ============================================================ */

/**
 * Dijalankan saat URL Web App dibuka.
 * Menampilkan Index.html beserta meta viewport agar pas di layar HP.
 */
function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Tugasku — To-Do List')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .addMetaTag('mobile-web-app-capable', 'yes')
    // Izinkan aplikasi ditampilkan di dalam "bungkus" PWA (GitHub Pages)
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/**
 * Menyisipkan isi file HTML lain (Style.html / Script.html) ke Index.html.
 * Dipanggil di Index.html dengan: <?!= include('Style'); ?>
 */
function include(namaFile) {
  return HtmlService.createHtmlOutputFromFile(namaFile).getContent();
}

/**
 * Menambahkan menu "📋 To-Do List" di Google Sheets
 * supaya fungsi penting bisa dijalankan tanpa membuka editor.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📋 To-Do List')
    .addItem('1. Setup (buat sheet & header)', 'setup')
    .addSeparator()
    .addItem('Aktifkan pengingat email', 'buatTrigger')
    .addItem('Matikan pengingat email', 'hapusTrigger')
    .addItem('Uji kirim pengingat sekarang', 'kirimPengingatHarian')
    .addToUi();
}


/* ============================================================
 * 3. SETUP — membuat semua sheet & header jika belum ada
 * ============================================================ */

/**
 * Jalankan SEKALI dari editor Apps Script (pilih "setup" lalu klik Jalankan).
 * Aman dijalankan berulang kali: data yang sudah ada tidak akan terhapus.
 */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  // Samakan zona waktu spreadsheet dengan WIB
  ss.setSpreadsheetTimeZone(ZONA_WAKTU);

  // --- Sheet Tugas ---
  const shTugas = siapkanSheet_(ss, SHEET_TUGAS, HEADER_TUGAS);
  const barisMaks = shTugas.getMaxRows() - 1;
  shTugas.getRange(2, 1, barisMaks, 1).setNumberFormat('@');                 // ID sebagai teks
  shTugas.getRange(2, 6, barisMaks, 1).setNumberFormat('dd/mm/yyyy');        // Deadline
  shTugas.getRange(2, 8, barisMaks, 2).setNumberFormat('dd/mm/yyyy hh:mm');  // Tgl dibuat & selesai
  shTugas.setColumnWidth(2, 260);
  shTugas.setColumnWidth(3, 320);

  // --- Sheet Kategori (isi kategori bawaan jika masih kosong) ---
  const shKategori = siapkanSheet_(ss, SHEET_KATEGORI, HEADER_KATEGORI);
  if (shKategori.getLastRow() < 2) {
    shKategori.getRange(2, 1, KATEGORI_AWAL.length, 2).setValues(KATEGORI_AWAL);
  }
  shKategori.setColumnWidth(1, 200);

  // --- Sheet Pengaturan (isi nilai bawaan jika masih kosong) ---
  const shPengaturan = siapkanSheet_(ss, SHEET_PENGATURAN, HEADER_PENGATURAN);
  if (shPengaturan.getLastRow() < 2) {
    const emailSaya = Session.getEffectiveUser().getEmail();
    shPengaturan.getRange(2, 1, 1, 3).setValues([[emailSaya, JAM_BAWAAN, 'Tidak']]);
  }
  shPengaturan.setColumnWidth(1, 260);

  // Hapus sheet kosong bawaan ("Sheet1" / "Lembar1") jika ada
  ['Sheet1', 'Lembar1', 'Lembar 1'].forEach(function (nama) {
    const sh = ss.getSheetByName(nama);
    if (sh && sh.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });

  const pesan = 'Setup selesai ✔ Sheet "Tugas", "Kategori", dan "Pengaturan" sudah siap.';
  Logger.log(pesan);
  return pesan;
}

/**
 * Membuat sheet bila belum ada, lalu memastikan header di baris 1 benar.
 */
function siapkanSheet_(ss, nama, header) {
  let sh = ss.getSheetByName(nama);
  if (!sh) sh = ss.insertSheet(nama);

  const rangeHeader = sh.getRange(1, 1, 1, header.length);
  const headerSekarang = rangeHeader.getValues()[0].map(String);
  if (headerSekarang.join('|') !== header.join('|')) {
    rangeHeader.setValues([header]);
  }
  rangeHeader.setFontWeight('bold').setBackground('#1e293b').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  return sh;
}


/* ============================================================
 * 4. FUNGSI BANTU (HELPER)
 *    Nama fungsi diakhiri "_" = privat, tidak bisa dipanggil dari browser.
 * ============================================================ */

/** Mengambil sheet berdasarkan nama, beri pesan jelas jika belum di-setup. */
function sheet_(nama) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nama);
  if (!sh) throw new Error('Sheet "' + nama + '" belum ada. Jalankan fungsi setup() terlebih dahulu.');
  return sh;
}

/** Cek apakah sebuah nilai adalah objek Date. */
function isDate_(v) {
  return Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime());
}

/** Menambah angka 0 di depan (6 -> "06"). */
function pad2_(n) {
  return ('0' + Number(n)).slice(-2);
}

/** Tanggal hari ini (WIB) dalam format "yyyy-MM-dd". */
function hariIni_() {
  return Utilities.formatDate(new Date(), ZONA_WAKTU, 'yyyy-MM-dd');
}

/**
 * Mengubah nilai sel (Date atau teks) menjadi string "yyyy-MM-dd".
 * Format string dipakai agar perbandingan tanggal selalu akurat & aman dikirim ke browser.
 */
function tglStr_(v) {
  if (v === '' || v === null || v === undefined) return '';
  if (isDate_(v)) return Utilities.formatDate(v, ZONA_WAKTU, 'yyyy-MM-dd');
  const s = String(v).trim();
  let m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);          // 2026-11-20
  if (m) return m[1] + '-' + pad2_(m[2]) + '-' + pad2_(m[3]);
  m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);            // 20/11/2026
  if (m) return m[3] + '-' + pad2_(m[2]) + '-' + pad2_(m[1]);
  return '';
}

/** Mengubah nilai sel waktu menjadi string "yyyy-MM-dd HH:mm". */
function waktuStr_(v) {
  if (!v) return '';
  if (isDate_(v)) return Utilities.formatDate(v, ZONA_WAKTU, 'yyyy-MM-dd HH:mm');
  return String(v);
}

/** Mengubah "yyyy-MM-dd" menjadi objek Date (pukul 00.00 WIB) untuk disimpan di sheet. */
function strKeTanggal_(s) {
  const p = s.split('-').map(Number);
  return new Date(p[0], p[1] - 1, p[2]);
}

/** Selisih hari antara dua tanggal "yyyy-MM-dd" (b - a). */
function selisihHari_(a, b) {
  const pa = a.split('-').map(Number), pb = b.split('-').map(Number);
  return Math.round((Date.UTC(pb[0], pb[1] - 1, pb[2]) - Date.UTC(pa[0], pa[1] - 1, pa[2])) / 86400000);
}

/** Format tanggal panjang Bahasa Indonesia: "Kamis, 20 November 2026". */
function tanggalPanjang_(s) {
  if (!s) return '-';
  const p = s.split('-').map(Number);
  const d = new Date(p[0], p[1] - 1, p[2]);
  return NAMA_HARI[d.getDay()] + ', ' + p[2] + ' ' + NAMA_BULAN[p[1] - 1] + ' ' + p[0];
}

/**
 * Mencegah teks dianggap rumus oleh Google Sheets
 * (misalnya judul yang diawali "=", "+", "-", atau "@").
 */
function amanSel_(s) {
  s = String(s || '').trim();
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

/** Escape karakter HTML agar aman ditampilkan di email. */
function esc_(s) {
  return String(s || '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
}

/** Pembanding untuk mengurutkan berdasarkan prioritas lalu deadline. */
function urutPrioritas_(a, b) {
  return (URUTAN_PRIORITAS[a.prioritas] || 9) - (URUTAN_PRIORITAS[b.prioritas] || 9) ||
    String(a.deadline).localeCompare(String(b.deadline));
}


/* ============================================================
 * 5. DATA TUGAS (CRUD)
 * ============================================================ */

/** Membaca semua baris di sheet Tugas menjadi array objek. */
function bacaTugas_() {
  const sh = sheet_(SHEET_TUGAS);
  const akhir = sh.getLastRow();
  if (akhir < 2) return [];
  const data = sh.getRange(2, 1, akhir - 1, HEADER_TUGAS.length).getValues();
  const hasil = [];
  data.forEach(function (r, i) {
    if (!r[0]) return; // lewati baris tanpa ID
    hasil.push({
      id: String(r[0]),
      judul: String(r[1]),
      deskripsi: String(r[2]),
      kategori: String(r[3]),
      prioritas: URUTAN_PRIORITAS[r[4]] ? String(r[4]) : 'Sedang',
      deadline: tglStr_(r[5]),
      status: String(r[6]) === STATUS_SELESAI ? STATUS_SELESAI : STATUS_BELUM,
      dibuat: waktuStr_(r[7]),
      selesai: waktuStr_(r[8]),
      baris: i + 2 // nomor baris di sheet (dipakai saat edit/hapus)
    });
  });
  return hasil;
}

/** Mencari tugas berdasarkan ID. */
function cariTugas_(id) {
  return bacaTugas_().filter(function (t) { return t.id === String(id); })[0] || null;
}

/**
 * Mengambil SEMUA data yang dibutuhkan halaman web dalam satu kali panggilan.
 * Dipanggil dari browser: google.script.run.ambilSemuaData()
 */
function ambilSemuaData() {
  return {
    tugas: bacaTugas_().map(function (t) { delete t.baris; return t; }),
    kategori: bacaKategori_(),
    pengaturan: ambilPengaturan(),
    hariIni: hariIni_()
  };
}

/**
 * Menambah tugas baru (jika data.id kosong) atau mengedit tugas (jika data.id ada).
 * @param {Object} data {id, judul, deskripsi, kategori, prioritas, deadline}
 */
function simpanTugas(data) {
  const judul = String(data.judul || '').trim();
  if (!judul) throw new Error('Judul tugas wajib diisi.');

  const deskripsi = amanSel_(data.deskripsi);
  const kategori = amanSel_(data.kategori);
  const prioritas = URUTAN_PRIORITAS[data.prioritas] ? data.prioritas : 'Sedang';
  const deadline = tglStr_(data.deadline);
  const nilaiDeadline = deadline ? strKeTanggal_(deadline) : '';

  // Kunci agar tidak bentrok jika dua perangkat menyimpan bersamaan
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const sh = sheet_(SHEET_TUGAS);
    if (data.id) {
      // --- EDIT ---
      const t = cariTugas_(data.id);
      if (!t) throw new Error('Tugas tidak ditemukan. Mungkin sudah dihapus.');
      sh.getRange(t.baris, 2, 1, 5).setValues([[amanSel_(judul), deskripsi, kategori, prioritas, nilaiDeadline]]);
    } else {
      // --- TAMBAH BARU ---
      const id = 'T' + Utilities.formatDate(new Date(), ZONA_WAKTU, 'yyMMddHHmmss') +
        Math.floor(Math.random() * 900 + 100);
      sh.appendRow([id, amanSel_(judul), deskripsi, kategori, prioritas, nilaiDeadline, STATUS_BELUM, new Date(), '']);
    }
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return ambilSemuaData();
}

/** Menghapus tugas berdasarkan ID. */
function hapusTugas(id) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const t = cariTugas_(id);
    if (!t) throw new Error('Tugas tidak ditemukan. Mungkin sudah dihapus.');
    sheet_(SHEET_TUGAS).deleteRow(t.baris);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return ambilSemuaData();
}

/**
 * Menandai tugas selesai / belum selesai (dari checkbox).
 * @param {string} id
 * @param {boolean} selesai
 */
function ubahStatus(id, selesai) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const t = cariTugas_(id);
    if (!t) throw new Error('Tugas tidak ditemukan. Mungkin sudah dihapus.');
    const sh = sheet_(SHEET_TUGAS);
    sh.getRange(t.baris, 7).setValue(selesai ? STATUS_SELESAI : STATUS_BELUM);
    sh.getRange(t.baris, 9).setValue(selesai ? new Date() : '');
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return ambilSemuaData();
}


/* ============================================================
 * 6. KATEGORI
 * ============================================================ */

/** Membaca daftar kategori beserta warnanya. */
function bacaKategori_() {
  const sh = sheet_(SHEET_KATEGORI);
  const akhir = sh.getLastRow();
  if (akhir < 2) return [];
  return sh.getRange(2, 1, akhir - 1, 2).getValues()
    .filter(function (r) { return String(r[0]).trim(); })
    .map(function (r) {
      const warna = /^#[0-9a-f]{6}$/i.test(String(r[1]).trim()) ? String(r[1]).trim() : '#64748b';
      return { nama: String(r[0]).trim(), warna: warna };
    });
}

/**
 * Menambah kategori baru, atau mengubah warna jika nama kategori sudah ada.
 * @param {Object} data {nama, warna}
 */
function simpanKategori(data) {
  const nama = String(data.nama || '').trim();
  if (!nama) throw new Error('Nama kategori wajib diisi.');
  const warna = /^#[0-9a-f]{6}$/i.test(String(data.warna)) ? data.warna : '#64748b';

  const sh = sheet_(SHEET_KATEGORI);
  const akhir = sh.getLastRow();
  const nama2 = akhir >= 2 ? sh.getRange(2, 1, akhir - 1, 1).getValues().map(function (r) { return String(r[0]).trim().toLowerCase(); }) : [];
  const idx = nama2.indexOf(nama.toLowerCase());

  if (idx >= 0) {
    sh.getRange(idx + 2, 2).setValue(warna);           // ubah warna
  } else {
    sh.appendRow([amanSel_(nama), warna]);             // kategori baru
  }
  return ambilSemuaData();
}

/**
 * Menghapus kategori. Tugas yang memakai kategori ini TIDAK ikut terhapus
 * (kategorinya tetap tertulis, hanya tampil berwarna abu-abu).
 */
function hapusKategori(nama) {
  const sh = sheet_(SHEET_KATEGORI);
  const akhir = sh.getLastRow();
  if (akhir < 2) return ambilSemuaData();
  const daftar = sh.getRange(2, 1, akhir - 1, 1).getValues();
  for (let i = daftar.length - 1; i >= 0; i--) {
    if (String(daftar[i][0]).trim() === String(nama).trim()) sh.deleteRow(i + 2);
  }
  return ambilSemuaData();
}


/* ============================================================
 * 7. PENGATURAN
 * ============================================================ */

/** Membaca pengaturan dari sheet "Pengaturan" (baris 2). */
function ambilPengaturan() {
  const sh = sheet_(SHEET_PENGATURAN);
  const v = sh.getRange(2, 1, 1, 3).getValues()[0];

  // Jam pengingat: bisa berupa angka (6), teks ("06.00"), atau format waktu
  let jam = isDate_(v[1]) ? v[1].getHours() : parseInt(String(v[1]), 10);
  if (isNaN(jam) || jam < 0 || jam > 23) jam = JAM_BAWAAN;

  let urlApp = '';
  try { urlApp = ScriptApp.getService().getUrl() || ''; } catch (e) { /* belum di-deploy */ }

  return {
    email: String(v[0] || '').trim(),
    jam: jam,
    aktif: String(v[2]).trim() === 'Ya' ? 'Ya' : 'Tidak',
    triggerTerpasang: daftarTriggerPengingat_().length > 0,
    urlApp: urlApp
  };
}

/**
 * Menyimpan pengaturan dari halaman web, lalu otomatis
 * memasang atau mematikan trigger sesuai pilihan "Pengingat Aktif".
 * @param {Object} data {email, jam, aktif}
 */
function simpanPengaturan(data) {
  // Validasi email (boleh lebih dari satu, pisahkan dengan koma)
  const daftarEmail = String(data.email || '').split(',')
    .map(function (e) { return e.trim(); })
    .filter(String);
  if (!daftarEmail.length) throw new Error('Email penerima wajib diisi.');
  daftarEmail.forEach(function (e) {
    if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(e)) throw new Error('Format email tidak valid: ' + e);
  });

  const jam = parseInt(data.jam, 10);
  if (isNaN(jam) || jam < 0 || jam > 23) throw new Error('Jam pengingat harus antara 0 dan 23.');
  const aktif = data.aktif === 'Ya' ? 'Ya' : 'Tidak';

  sheet_(SHEET_PENGATURAN).getRange(2, 1, 1, 3).setValues([[daftarEmail.join(', '), jam, aktif]]);
  SpreadsheetApp.flush();

  // Pasang / matikan trigger sesuai pengaturan
  if (aktif === 'Ya') buatTrigger(); else hapusTrigger();

  return ambilSemuaData();
}


/* ============================================================
 * 8. TRIGGER WAKTU HARIAN
 * ============================================================ */

/** Daftar trigger milik proyek ini yang menjalankan kirimPengingatHarian. */
function daftarTriggerPengingat_() {
  return ScriptApp.getProjectTriggers().filter(function (t) {
    return t.getHandlerFunction() === FUNGSI_PENGINGAT;
  });
}

/**
 * Memasang trigger harian untuk kirimPengingatHarian().
 * Trigger lama dihapus lebih dulu, jadi TIDAK akan ada trigger ganda
 * walaupun fungsi ini dijalankan berkali-kali.
 */
function buatTrigger() {
  const p = ambilPengaturan();

  // Hapus trigger lama (mencegah dobel)
  daftarTriggerPengingat_().forEach(function (t) { ScriptApp.deleteTrigger(t); });

  // Pasang trigger baru: setiap hari pada rentang jam yang dipilih (WIB)
  ScriptApp.newTrigger(FUNGSI_PENGINGAT)
    .timeBased()
    .everyDays(1)
    .atHour(p.jam)
    .inTimezone(ZONA_WAKTU)
    .create();

  // Catat status di sheet Pengaturan
  const sh = sheet_(SHEET_PENGATURAN);
  if (!p.email) sh.getRange(2, 1).setValue(Session.getEffectiveUser().getEmail());
  sh.getRange(2, 3).setValue('Ya');

  const pesan = 'Pengingat AKTIF ✔ Email dikirim setiap hari antara pukul ' +
    pad2_(p.jam) + '.00–' + pad2_((p.jam + 1) % 24) + '.00 WIB.';
  Logger.log(pesan);
  return pesan;
}

/** Mematikan pengingat: menghapus semua trigger kirimPengingatHarian. */
function hapusTrigger() {
  const daftar = daftarTriggerPengingat_();
  daftar.forEach(function (t) { ScriptApp.deleteTrigger(t); });
  sheet_(SHEET_PENGATURAN).getRange(2, 3).setValue('Tidak');

  const pesan = 'Pengingat DIMATIKAN. ' + daftar.length + ' trigger dihapus.';
  Logger.log(pesan);
  return pesan;
}


/* ============================================================
 * 9. PENGINGAT EMAIL HARIAN
 * ============================================================ */

/**
 * Mengirim SATU email ringkasan berisi:
 *  - tugas "Belum Selesai" dengan deadline HARI INI (prioritas Tinggi paling atas, warna merah)
 *  - tugas yang sudah TERLAMBAT (bagian terpisah)
 * Tidak mengirim email jika keduanya kosong.
 *
 * @param {Object} e  Objek event dari trigger (kosong jika dijalankan manual).
 */
function kirimPengingatHarian(e) {
  const p = ambilPengaturan();

  // Jika dijalankan oleh trigger tapi pengingat dimatikan di sheet -> lewati
  if (e && p.aktif !== 'Ya') {
    Logger.log('Pengingat sedang nonaktif. Email tidak dikirim.');
    return 'Pengingat sedang nonaktif. Email tidak dikirim.';
  }

  const hari = hariIni_();
  const belum = bacaTugas_().filter(function (t) { return t.status !== STATUS_SELESAI && t.deadline; });

  // Tugas deadline hari ini, prioritas Tinggi paling atas
  const tugasHariIni = belum.filter(function (t) { return t.deadline === hari; }).sort(urutPrioritas_);

  // Tugas terlambat, yang paling lama terlambat di atas
  const tugasTerlambat = belum.filter(function (t) { return t.deadline < hari; })
    .sort(function (a, b) { return a.deadline.localeCompare(b.deadline) || urutPrioritas_(a, b); });

  if (!tugasHariIni.length && !tugasTerlambat.length) {
    const kosong = 'Tidak ada tugas untuk hari ini maupun yang terlambat. Email tidak dikirim.';
    Logger.log(kosong);
    return kosong;
  }

  // Peta warna kategori
  const warnaKategori = {};
  bacaKategori_().forEach(function (k) { warnaKategori[k.nama] = k.warna; });

  const penerima = p.email || Session.getEffectiveUser().getEmail();
  const subjek = '📋 Pengingat Tugas ' + tanggalPanjang_(hari) + ' — ' +
    tugasHariIni.length + ' hari ini' + (tugasTerlambat.length ? ', ' + tugasTerlambat.length + ' terlambat' : '');

  const htmlBody = buatEmailHtml_(hari, tugasHariIni, tugasTerlambat, warnaKategori, p.urlApp);
  const teksBiasa = buatEmailTeks_(hari, tugasHariIni, tugasTerlambat, p.urlApp);

  MailApp.sendEmail({
    to: penerima,
    subject: subjek,
    body: teksBiasa,      // cadangan untuk aplikasi email yang tidak mendukung HTML
    htmlBody: htmlBody,
    name: 'To-Do List Saya'
  });

  const pesan = 'Email pengingat terkirim ke ' + penerima + ' (' +
    tugasHariIni.length + ' tugas hari ini, ' + tugasTerlambat.length + ' terlambat).';
  Logger.log(pesan);
  return pesan;
}

/** Menyusun isi email dalam format HTML (gaya inline agar rapi di Gmail). */
function buatEmailHtml_(hari, tugasHariIni, tugasTerlambat, warnaKategori, urlApp) {
  const WARNA_PRIO = { 'Tinggi': '#dc2626', 'Sedang': '#d97706', 'Rendah': '#16a34a' };

  // Satu kartu tugas
  function kartu(t, terlambat) {
    const tinggi = t.prioritas === 'Tinggi';
    const merah = tinggi || terlambat;
    const wKat = warnaKategori[t.kategori] || '#64748b';
    let info = '';
    if (terlambat) {
      const n = selisihHari_(t.deadline, hari);
      info = '<div style="margin-top:6px;font-size:12px;color:#dc2626;font-weight:600;">⏰ Deadline ' +
        esc_(tanggalPanjang_(t.deadline)) + ' — terlambat ' + n + ' hari</div>';
    }
    return '' +
      '<tr><td style="padding:0 0 10px 0;">' +
      '<div style="border:1px solid ' + (merah ? '#fecaca' : '#e2e8f0') + ';border-left:5px solid ' + WARNA_PRIO[t.prioritas] +
      ';background:' + (merah ? '#fef2f2' : '#ffffff') + ';border-radius:10px;padding:12px 14px;">' +
      '<div style="font-size:15px;font-weight:700;color:' + (merah ? '#b91c1c' : '#0f172a') + ';">' +
      (tinggi ? '🔴 ' : '') + esc_(t.judul) + '</div>' +
      '<div style="margin-top:6px;">' +
      '<span style="display:inline-block;font-size:11px;font-weight:700;color:#fff;background:' + WARNA_PRIO[t.prioritas] +
      ';padding:2px 8px;border-radius:20px;margin-right:4px;">Prioritas ' + esc_(t.prioritas) + '</span>' +
      (t.kategori ? '<span style="display:inline-block;font-size:11px;font-weight:700;color:#fff;background:' + wKat +
        ';padding:2px 8px;border-radius:20px;">' + esc_(t.kategori) + '</span>' : '') +
      '</div>' +
      (t.deskripsi ? '<div style="margin-top:8px;font-size:13px;color:#475569;line-height:1.5;">' +
        esc_(t.deskripsi).replace(/\n/g, '<br>') + '</div>' : '') +
      info +
      '</div></td></tr>';
  }

  // Satu bagian (judul + daftar kartu)
  function bagian(judul, warna, daftar, terlambat) {
    if (!daftar.length) return '';
    return '<h2 style="font-size:16px;margin:22px 0 10px;color:' + warna + ';">' + judul + ' (' + daftar.length + ')</h2>' +
      '<table role="presentation" width="100%" cellspacing="0" cellpadding="0">' +
      daftar.map(function (t) { return kartu(t, terlambat); }).join('') + '</table>';
  }

  const tombol = urlApp
    ? '<div style="text-align:center;margin:26px 0 8px;"><a href="' + esc_(urlApp) +
      '" style="background:#4f46e5;color:#ffffff;text-decoration:none;font-weight:700;padding:12px 26px;border-radius:10px;display:inline-block;">Buka Aplikasi To-Do List</a></div>'
    : '';

  return '' +
    '<div style="background:#f1f5f9;padding:20px 10px;font-family:Arial,Helvetica,sans-serif;">' +
    '<div style="max-width:600px;margin:0 auto;background:#ffffff;border-radius:14px;overflow:hidden;border:1px solid #e2e8f0;">' +
    '<div style="background:#4f46e5;color:#ffffff;padding:20px 22px;">' +
    '<div style="font-size:13px;opacity:.85;">Pengingat Tugas Harian</div>' +
    '<div style="font-size:20px;font-weight:700;margin-top:4px;">' + esc_(tanggalPanjang_(hari)) + '</div>' +
    '</div>' +
    '<div style="padding:6px 22px 20px;">' +
    '<p style="font-size:14px;color:#334155;margin:16px 0 0;">Selamat pagi! Berikut ringkasan tugas Anda: <b>' +
    tugasHariIni.length + '</b> tugas jatuh tempo hari ini dan <b style="color:#dc2626;">' +
    tugasTerlambat.length + '</b> tugas sudah terlambat.</p>' +
    bagian('📌 Deadline Hari Ini', '#0f172a', tugasHariIni, false) +
    bagian('⚠️ Sudah Terlambat', '#dc2626', tugasTerlambat, true) +
    tombol +
    '</div>' +
    '<div style="background:#f8fafc;color:#94a3b8;font-size:11px;text-align:center;padding:12px;">' +
    'Email ini dikirim otomatis oleh aplikasi To-Do List (Google Apps Script).<br>' +
    'Untuk mematikan pengingat, buka menu Pengaturan di aplikasi.</div>' +
    '</div></div>';
}

/** Versi teks biasa dari email (cadangan). */
function buatEmailTeks_(hari, tugasHariIni, tugasTerlambat, urlApp) {
  let s = 'PENGINGAT TUGAS — ' + tanggalPanjang_(hari) + '\n\n';
  if (tugasHariIni.length) {
    s += 'DEADLINE HARI INI:\n';
    tugasHariIni.forEach(function (t, i) {
      s += (i + 1) + '. ' + t.judul + ' [' + t.prioritas + ' | ' + (t.kategori || '-') + ']' +
        (t.deskripsi ? '\n   ' + t.deskripsi : '') + '\n';
    });
    s += '\n';
  }
  if (tugasTerlambat.length) {
    s += 'SUDAH TERLAMBAT:\n';
    tugasTerlambat.forEach(function (t, i) {
      s += (i + 1) + '. ' + t.judul + ' [' + t.prioritas + ' | ' + (t.kategori || '-') + '] — terlambat ' +
        selisihHari_(t.deadline, hari) + ' hari\n';
    });
  }
  if (urlApp) s += '\nBuka aplikasi: ' + urlApp + '\n';
  return s;
}
