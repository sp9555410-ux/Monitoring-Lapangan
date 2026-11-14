/** Monitoring Lapangan - Backend (Google Apps Script) sebagai API untuk PWA.
 *  Database: Google Sheets (sheet "Alat", "Log", "Langsir", "Galian"). Jalankan setup() sekali.
 *  Token API disimpan di Project Settings > Script Properties dengan nama API_TOKEN. */

/** Isi dengan ID Google Sheet (bagian antara /d/ dan /edit pada URL).
 *  Kosongkan jika script dibuat dari menu Extensions > Apps Script di dalam sheet. */
const SPREADSHEET_ID = '';

function ss_() {
  return SPREADSHEET_ID ? SpreadsheetApp.openById(SPREADSHEET_ID) : SpreadsheetApp.getActiveSpreadsheet();
}

const SHEET_ALAT = 'Alat';
const SHEET_LOG = 'Log';
/** Target jam kerja efektif per alat per hari (dasar hitung utilisasi). */
const TARGET_JAM_HARI = 8;

/** Target volume galian default per excavator per hari (m3). */
const TARGET_VOL_DEFAULT = 200;

/** Target TOTAL volume proyek (m3). */
const TARGET_TOTAL_GALIAN = 22000;
const TARGET_TOTAL_LANGSIR = 22000;

/** Target jam rental per unit alat. */
const TARGET_RENTAL_JAM = 200;

/** Nama proyek/lokasi yang tampil di judul laporan harian (boleh dikosongkan). */
const NAMA_PROYEK = '';

const SHEET_LANGSIR = 'Langsir';
const HEAD_LANGSIR = ['Waktu Input', 'Tanggal', 'Kode DT', 'Sopir', 'Excavator Pemuat', 'Asal (Lokasi Galian)',
  'Tujuan (Lokasi Buang)', 'Material', 'Jumlah Rit', 'Kapasitas/Rit (m3)', 'Volume (m3)', 'Jarak (km)', 'Catatan',
  'Excavator Bongkar'];
const SHEET_GALIAN = 'Galian';
const HEAD_GALIAN = ['Waktu Input', 'Tanggal', 'Excavator', 'Lokasi Galian', 'Material', 'Volume (m3)', 'Rit DT Dimuat', 'Catatan'];
const HEAD_ALAT = ['Kode', 'Nama', 'Jenis', 'Status', 'Target Vol/Hari (m3)'];
const HEAD_LOG = ['Waktu Input', 'Tanggal', 'Kode Alat', 'Operator', 'Lokasi',
  'HM Awal', 'HM Akhir', 'Jam Kerja', 'BBM (L)', 'Status', 'Catatan'];

/* ===================== API (dipanggil PWA) ===================== */

function apiToken_() {
  return PropertiesService.getScriptProperties().getProperty('API_TOKEN');
}

function doGet() {
  return ContentService.createTextOutput('API Monitoring Lapangan aktif');
}

function doPost(e) {
  const API = {
    getData, saveLog, saveLangsir, saveGalian, getProduktivitas, getLaporanHarian,
    getHasilGalian, getRekapGalian, setTargetGalian, getDashboard
  };
  let out;
  try {
    const r = JSON.parse(e.postData.contents);
    const tk = apiToken_();
    if (!tk || r.token !== tk) throw new Error('Kode akses salah.');
    if (!Object.prototype.hasOwnProperty.call(API, r.fn)) throw new Error('Fungsi tidak dikenal.');
    out = { data: API[r.fn].apply(null, r.args || []) };
  } catch (err) {
    out = { error: String(err && err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/* ===================== Setup ===================== */

function setup() {
  const ss = ss_();
  let a = ss.getSheetByName(SHEET_ALAT) || ss.insertSheet(SHEET_ALAT);
  if (a.getLastRow() === 0) {
    a.appendRow(HEAD_ALAT);
    a.appendRow(['EXC-01', 'Excavator PC200', 'Excavator', 'Aktif', 200]);
    a.appendRow(['DT-01', 'Dump Truck HD785', 'Dump Truck', 'Aktif', '']);
    a.appendRow(['DZ-01', 'Bulldozer D6R', 'Bulldozer', 'Aktif', '']);
  } else if (!a.getRange(1, 5).getValue()) {
    a.getRange(1, 5).setValue(HEAD_ALAT[4]); // migrasi sheet lama
  }
  let l = ss.getSheetByName(SHEET_LOG) || ss.insertSheet(SHEET_LOG);
  if (l.getLastRow() === 0) l.appendRow(HEAD_LOG);
  let g = ss.getSheetByName(SHEET_LANGSIR) || ss.insertSheet(SHEET_LANGSIR);
  if (g.getLastRow() === 0) g.appendRow(HEAD_LANGSIR);
  else if (!g.getRange(1, 14).getValue()) g.getRange(1, 14).setValue(HEAD_LANGSIR[13]); // migrasi sheet lama
  let gl = ss.getSheetByName(SHEET_GALIAN) || ss.insertSheet(SHEET_GALIAN);
  if (gl.getLastRow() === 0) gl.appendRow(HEAD_GALIAN);
  [a, l, g, gl].forEach(s => s.getRange(1, 1, 1, s.getLastColumn()).setFontWeight('bold').setBackground('#f5c400'));
  const def = ss.getSheetByName('Sheet1');
  if (def && ss.getSheets().length > 2) ss.deleteSheet(def);
}

/* ===================== Pembaca data ===================== */

function getAlat_() {
  const s = ss_().getSheetByName(SHEET_ALAT);
  return s.getDataRange().getValues().slice(1)
    .filter(r => r[0] && r[3] !== 'Nonaktif')
    .map(r => ({ kode: r[0], nama: r[1], jenis: r[2], target: Number(r[4]) || 0 }));
}

function targetExc_(a) { return a.target > 0 ? a.target : TARGET_VOL_DEFAULT; }

function getLogs_() {
  const s = ss_().getSheetByName(SHEET_LOG);
  const tz = Session.getScriptTimeZone();
  return s.getDataRange().getValues().slice(1).map(r => ({
    tanggal: r[1] instanceof Date ? Utilities.formatDate(r[1], tz, 'yyyy-MM-dd') : String(r[1]),
    kode: r[2], operator: r[3], lokasi: r[4],
    hmAwal: Number(r[5]), hmAkhir: Number(r[6]), jam: Number(r[7]),
    bbm: Number(r[8]), status: r[9], catatan: r[10]
  }));
}

function getLangsir_() {
  const s = ss_().getSheetByName(SHEET_LANGSIR);
  if (!s) return [];
  const tz = Session.getScriptTimeZone();
  return s.getDataRange().getValues().slice(1).filter(r => r[2]).map(r => ({
    tanggal: r[1] instanceof Date ? Utilities.formatDate(r[1], tz, 'yyyy-MM-dd') : String(r[1]),
    kode: r[2], sopir: r[3], excavator: r[4], asal: r[5], tujuan: r[6], material: r[7],
    rit: Number(r[8]), kap: Number(r[9]), vol: Number(r[10]), jarak: Number(r[11]), catatan: r[12],
    excBongkar: r[13] || ''
  }));
}

/** Catatan hasil galian yang diinput langsung (sheet "Galian"). */
function getGalian_() {
  const s = ss_().getSheetByName(SHEET_GALIAN);
  if (!s) return [];
  const tz = Session.getScriptTimeZone();
  return s.getDataRange().getValues().slice(1).filter(r => r[2] && Number(r[5]) > 0).map(r => ({
    tanggal: r[1] instanceof Date ? Utilities.formatDate(r[1], tz, 'yyyy-MM-dd') : String(r[1]),
    excavator: r[2], lokasi: r[3] || '', material: r[4] || '', vol: Number(r[5]),
    rit: Number(r[6]) || 0, catatan: r[7] || '', sumber: 'Input'
  }));
}

/** Sumber data hasil galian untuk semua laporan.
 *  Input langsung (sheet Galian) dipakai lebih dulu. Untuk excavator + tanggal yang belum punya input langsung,
 *  dipakai hitungan otomatis dari Langsir (kolom Excavator Pemuat). Tidak ada hitungan ganda. */
function galianRows_() {
  const direct = getGalian_();
  const key = r => r.tanggal + '|' + r.excavator;
  const have = new Set(direct.map(key));
  const derived = getLangsir_().filter(l => l.excavator && !have.has(key(l))).map(l => ({
    tanggal: l.tanggal, excavator: l.excavator, lokasi: l.asal, material: l.material,
    vol: l.vol, rit: l.rit, catatan: l.catatan, sumber: 'Langsir'
  }));
  return direct.concat(derived);
}

/** Progres kumulatif terhadap target total. upto = yyyy-MM-dd (kosong = semua data). */
function progresTotal_(rows, upto) {
  rows = (rows || getLangsir_()).filter(r => !upto || r.tanggal <= upto);
  const gl = galianRows_().filter(r => !upto || r.tanggal <= upto);
  const sum = (arr, k) => arr.reduce((t, x) => t + x[k], 0);
  const r1 = n => +n.toFixed(1);
  const mk = (target, real, hari) => {
    const rata = hari ? real / hari : 0, sisa = Math.max(0, target - real);
    return { target, real: r1(real), sisa: r1(sisa), capai: target > 0 ? r1(real / target * 100) : 0,
      hari, rata: r1(rata), estHari: rata > 0 && sisa > 0 ? Math.ceil(sisa / rata) : 0 };
  };
  return {
    galian: mk(TARGET_TOTAL_GALIAN, sum(gl, 'vol'), new Set(gl.map(r => r.tanggal)).size),
    langsir: mk(TARGET_TOTAL_LANGSIR, sum(rows, 'vol'), new Set(rows.map(r => r.tanggal)).size)
  };
}

function jenisAlat_(a) {
  const t = a.jenis + ' ' + a.nama;
  if (/dump|truck|truk/i.test(t) || /^DT/i.test(a.kode)) return 'DT';
  if (/excav|ekskav|backhoe/i.test(t) || /^EXC/i.test(a.kode)) return 'EXC';
  return 'LAIN';
}

function langsirData_(alat) {
  const isDT = a => /dump|truck|truk/i.test(a.jenis + ' ' + a.nama) || /^DT/i.test(a.kode);
  const isExc = a => /excav|ekskav|backhoe/i.test(a.jenis + ' ' + a.nama) || /^EXC/i.test(a.kode);
  let dt = alat.filter(isDT); if (!dt.length) dt = alat;
  let exc = alat.filter(isExc); if (!exc.length) exc = alat;
  const rows = getLangsir_();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const sum = (arr, k) => arr.reduce((t, x) => t + x[k], 0);
  const hari = rows.filter(r => r.tanggal === today);
  const perDT = dt.map(a => {
    const m = rows.filter(r => r.kode === a.kode);
    return { kode: a.kode, nama: a.nama, rit: sum(m, 'rit'), vol: +sum(m, 'vol').toFixed(1) };
  });
  const rute = {};
  rows.forEach(r => {
    const k = (r.asal || '?') + ' \u2192 ' + (r.tujuan || '?');
    const x = rute[k] || (rute[k] = { rute: k, rit: 0, vol: 0 });
    x.rit += r.rit; x.vol = +(x.vol + r.vol).toFixed(1);
  });
  const uniq = k => [...new Set(rows.map(r => r[k]).filter(Boolean))];
  return {
    dt, exc,
    ringkas: {
      ritHariIni: sum(hari, 'rit'), volHariIni: +sum(hari, 'vol').toFixed(1),
      dtAktif: new Set(hari.map(r => r.kode)).size, volTotal: +sum(rows, 'vol').toFixed(1)
    },
    perDT,
    perRute: Object.values(rute).sort((a, b) => b.vol - a.vol).slice(0, 6),
    terbaru: rows.slice(-12).reverse(),
    lokasi: { asal: uniq('asal'), tujuan: uniq('tujuan') },
    progres: progresTotal_(rows)
  };
}

/* ===================== Endpoint data ===================== */

/** Data awal untuk UI: daftar alat + log + ringkasan. */
function getData() {
  const alat = getAlat_();
  const logs = getLogs_();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const rekap = {};
  alat.forEach(a => rekap[a.kode] = { kode: a.kode, nama: a.nama, jam: 0, bbm: 0, hmTerakhir: 0, status: '-' });
  logs.forEach(l => {
    const r = rekap[l.kode]; if (!r) return;
    r.jam += l.jam; r.bbm += l.bbm;
    if (l.hmAkhir >= r.hmTerakhir) { r.hmTerakhir = l.hmAkhir; r.status = l.status; }
  });
  const hariIni = logs.filter(l => l.tanggal === today);
  const tz = Session.getScriptTimeZone();
  const perHari = {};
  logs.forEach(l => {
    const h = perHari[l.tanggal] || (perHari[l.tanggal] = { bbm: 0, jam: 0 });
    h.bbm += l.bbm; h.jam += l.jam;
  });
  const harian = [];
  for (let i = 29; i >= 0; i--) {
    const k = Utilities.formatDate(new Date(Date.now() - i * 86400000), tz, 'yyyy-MM-dd');
    const h = perHari[k] || { bbm: 0, jam: 0 };
    harian.push({ tanggal: k, bbm: +h.bbm.toFixed(1), jam: +h.jam.toFixed(1) });
  }
  return {
    langsir: langsirData_(alat),
    harian,
    alat,
    rekap: Object.values(rekap).map(r => ({ ...r, lPerJam: r.jam ? +(r.bbm / r.jam).toFixed(1) : 0 })),
    terbaru: logs.slice(-15).reverse(),
    ringkas: {
      jamHariIni: +hariIni.reduce((t, l) => t + l.jam, 0).toFixed(1),
      bbmHariIni: +hariIni.reduce((t, l) => t + l.bbm, 0).toFixed(1),
      alatBeroperasi: new Set(hariIni.filter(l => l.status === 'Beroperasi').map(l => l.kode)).size,
      alatRusak: Object.values(rekap).filter(r => r.status === 'Rusak').length
    },
    hmTerakhir: Object.fromEntries(Object.values(rekap).map(r => [r.kode, r.hmTerakhir]))
  };
}

function saveLog(d) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const hmAwal = Number(d.hmAwal), hmAkhir = Number(d.hmAkhir), bbm = Number(d.bbm || 0);
    if (!d.kode || !d.operator || !d.tanggal) throw new Error('Alat, operator, dan tanggal wajib diisi.');
    if (isNaN(hmAwal) || isNaN(hmAkhir) || hmAkhir < hmAwal) throw new Error('HM akhir tidak boleh lebih kecil dari HM awal.');
    const jam = +(hmAkhir - hmAwal).toFixed(1);
    ss_().getSheetByName(SHEET_LOG).appendRow([
      new Date(), d.tanggal, d.kode, d.operator, d.lokasi || '', hmAwal, hmAkhir, jam, bbm,
      d.status || 'Beroperasi', d.catatan || ''
    ]);
    return getData();
  } finally { lock.releaseLock(); }
}

/** Simpan catatan langsir (hauling) material bekas galian oleh Dump Truck. */
function saveLangsir(d) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const rit = Number(d.rit), kap = Number(d.kapasitas), jarak = Number(d.jarak || 0);
    if (!d.kode || !d.sopir || !d.tanggal) throw new Error('DT, sopir, dan tanggal wajib diisi.');
    if (!d.asal || !d.tujuan) throw new Error('Lokasi asal galian dan tujuan buang wajib diisi.');
    if (!d.excavator) throw new Error('Excavator pemuat wajib dipilih (dipakai untuk laporan hasil galian).');
    if (!(rit > 0) || !(kap > 0)) throw new Error('Jumlah rit dan kapasitas per rit harus lebih dari 0.');
    const ss = ss_();
    let sh = ss.getSheetByName(SHEET_LANGSIR);
    if (!sh) { sh = ss.insertSheet(SHEET_LANGSIR); sh.appendRow(HEAD_LANGSIR); }
    sh.appendRow([new Date(), d.tanggal, d.kode, d.sopir, d.excavator || '', d.asal, d.tujuan,
      d.material || 'Tanah', rit, kap, +(rit * kap).toFixed(1), jarak, d.catatan || '', d.excavatorBongkar || '']);
    return getData();
  } finally { lock.releaseLock(); }
}

/** Simpan hasil galian yang dicatat langsung per excavator. days = rentang tampilan (7/14/30). */
function saveGalian(d, days) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const vol = Number(d.vol), rit = Number(d.rit || 0);
    if (!d.tanggal || !d.excavator) throw new Error('Tanggal dan excavator wajib diisi.');
    if (!(vol > 0)) throw new Error('Volume hasil galian harus lebih dari 0.');
    if (!(rit >= 0)) throw new Error('Jumlah rit tidak boleh negatif.');
    const ss = ss_();
    let sh = ss.getSheetByName(SHEET_GALIAN);
    if (!sh) { sh = ss.insertSheet(SHEET_GALIAN); sh.appendRow(HEAD_GALIAN); }
    sh.appendRow([new Date(), d.tanggal, d.excavator, d.lokasi || '', d.material || 'Tanah', +vol.toFixed(1), rit, d.catatan || '']);
    return getHasilGalian(days);
  } finally { lock.releaseLock(); }
}

/** Produktivitas alat berat. days = 7 / 30 / 0 (semua data). */
function getProduktivitas(days) {
  days = Number(days) || 0;
  const tz = Session.getScriptTimeZone();
  const cutoff = days ? Utilities.formatDate(new Date(Date.now() - (days - 1) * 86400000), tz, 'yyyy-MM-dd') : '0000-00-00';
  const logs = getLogs_().filter(l => l.tanggal >= cutoff);
  const ls = getLangsir_().filter(l => l.tanggal >= cutoff);
  const sum = (arr, k) => arr.reduce((t, x) => t + x[k], 0);
  const r1 = n => +n.toFixed(1);
  const rows = getAlat_().map(a => {
    const k = jenisAlat_(a);
    const lg = logs.filter(l => l.kode === a.kode);
    const jam = sum(lg, 'jam'), bbm = sum(lg, 'bbm');
    const hari = new Set(lg.map(l => l.tanggal)).size;
    const rusak = new Set(lg.filter(l => l.status === 'Rusak').map(l => l.tanggal)).size;
    const mine = k === 'DT' ? ls.filter(x => x.kode === a.kode) : k === 'EXC' ? ls.filter(x => x.excavator === a.kode) : [];
    const vol = sum(mine, 'vol');
    return {
      kode: a.kode, nama: a.nama, jenis: k, hari,
      jam: r1(jam), bbm: r1(bbm), vol: r1(vol), rit: sum(mine, 'rit'),
      prod: jam > 0 && vol > 0 ? r1(vol / jam) : 0,
      lJam: jam > 0 ? r1(bbm / jam) : 0,
      lM3: vol > 0 ? +(bbm / vol).toFixed(2) : 0,
      util: hari ? r1(Math.min(100, jam / (hari * TARGET_JAM_HARI) * 100)) : 0,
      avail: hari ? r1((hari - rusak) / hari * 100) : 0
    };
  });
  const dt = rows.filter(r => r.jenis === 'DT');
  const jamDT = sum(dt, 'jam'), volDT = sum(dt, 'vol');
  const aktif = rows.filter(r => r.hari > 0);
  const volTotal = sum(ls, 'vol');
  const bbmAngkut = sum(rows.filter(r => r.jenis !== 'LAIN'), 'bbm');
  return {
    days, target: TARGET_JAM_HARI, rows,
    ringkas: {
      prodDT: jamDT > 0 ? r1(volDT / jamDT) : 0,
      util: aktif.length ? r1(sum(aktif, 'util') / aktif.length) : 0,
      vol: r1(volTotal), rit: sum(ls, 'rit'),
      lM3: volTotal > 0 ? +(bbmAngkut / volTotal).toFixed(2) : 0
    }
  };
}

/** Laporan harian lengkap untuk satu tanggal (format yyyy-MM-dd; kosong = hari ini). */
function getLaporanHarian(tgl) {
  const tz = Session.getScriptTimeZone();
  tgl = tgl || Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const alat = getAlat_();
  const nama = {}; alat.forEach(a => nama[a.kode] = a.nama);
  const jenis = {}; alat.forEach(a => jenis[a.kode] = jenisAlat_(a));
  const logs = getLogs_().filter(l => l.tanggal === tgl);
  const ls = getLangsir_().filter(l => l.tanggal === tgl);
  const gl = galianRows_().filter(l => l.tanggal === tgl);
  const sum = (arr, k) => arr.reduce((t, x) => t + x[k], 0);
  const r1 = n => +n.toFixed(1);

  const status = {};
  logs.forEach(l => status[l.kode] = l.status);
  const hitung = st => Object.values(status).filter(v => v === st).length;
  const jamDT = sum(logs.filter(l => jenis[l.kode] === 'DT'), 'jam');
  const volume = sum(ls, 'vol');
  const volGalian = sum(gl, 'vol');

  const galian = alat.filter(a => jenis[a.kode] === 'EXC').map(a => {
    const m = gl.filter(x => x.excavator === a.kode);
    const vol = sum(m, 'vol'), target = targetExc_(a);
    return { kode: a.kode, nama: a.nama, target, vol: r1(vol), rit: sum(m, 'rit'), capai: r1(vol / target * 100),
      catatan: [...new Set(m.map(x => String(x.catatan || '').trim()).filter(Boolean))].join(' | ') };
  });
  const kodeExc = new Set(galian.map(g => g.kode));
  const volTanpaExc = r1(sum(gl.filter(x => !kodeExc.has(x.excavator)), 'vol'));
  const targetGalian = sum(galian, 'target');

  const kendala = [];
  logs.forEach(l => {
    if (l.status === 'Rusak' || l.catatan)
      kendala.push({ sumber: 'Alat', kode: l.kode, teks: (l.status === 'Rusak' ? 'RUSAK. ' : '') + (l.catatan || '') });
  });
  ls.forEach(l => { if (l.catatan) kendala.push({ sumber: 'Langsir', kode: l.kode, teks: l.catatan }); });

  return {
    tanggal: tgl, proyek: NAMA_PROYEK,
    ringkas: {
      beroperasi: hitung('Beroperasi'), standby: hitung('Standby'), rusak: hitung('Rusak'),
      jam: r1(sum(logs, 'jam')), bbm: r1(sum(logs, 'bbm')),
      rit: sum(ls, 'rit'), volume: r1(volume),
      prodDT: jamDT > 0 && volume > 0 ? r1(volume / jamDT) : 0,
      targetGalian, volGalian: r1(volGalian),
      capaiGalian: targetGalian > 0 ? r1(volGalian / targetGalian * 100) : 0
    },
    alat: logs.map(l => ({ ...l, nama: nama[l.kode] || '' })),
    langsir: ls.map(l => ({ ...l, nama: nama[l.kode] || '' })),
    galian, volTanpaExc, progres: progresTotal_(null, tgl),
    kendala,
    belumLapor: alat.filter(a => !(a.kode in status)).map(a => a.kode + ' - ' + a.nama)
  };
}

/** Hasil galian per excavator per hari vs target. days = 7 / 14 / 30. */
function getHasilGalian(days) {
  days = [7, 14, 30].indexOf(Number(days)) >= 0 ? Number(days) : 14;
  const tz = Session.getScriptTimeZone();
  const tgls = [];
  for (let i = days - 1; i >= 0; i--)
    tgls.push(Utilities.formatDate(new Date(Date.now() - i * 86400000), tz, 'yyyy-MM-dd'));
  const sum = (arr, k) => arr.reduce((t, x) => t + x[k], 0);
  const r1 = n => +n.toFixed(1);

  const excs = getAlat_().filter(a => jenisAlat_(a) === 'EXC')
    .map(a => ({ kode: a.kode, nama: a.nama, target: targetExc_(a), custom: a.target > 0 }));
  const kodeExc = new Set(excs.map(e => e.kode));
  const semua = galianRows_();
  const ls = semua.filter(l => l.tanggal >= tgls[0] && l.tanggal <= tgls[days - 1]);
  const targetHari = sum(excs, 'target');

  const harian = tgls.map(t => {
    const dl = ls.filter(l => l.tanggal === t);
    const perExc = {}, catatanExc = {};
    excs.forEach(e => {
      const m = dl.filter(l => l.excavator === e.kode);
      perExc[e.kode] = r1(sum(m, 'vol'));
      catatanExc[e.kode] = [...new Set(m.map(l => String(l.catatan || '').trim()).filter(Boolean))].join(' | ');
    });
    const total = r1(sum(dl, 'vol'));
    const catatan = [...new Set(dl.map(l => String(l.catatan || '').trim()).filter(Boolean))].join(' | ');
    return { tanggal: t, total, rit: sum(dl, 'rit'), target: targetHari,
      capai: targetHari > 0 ? r1(total / targetHari * 100) : 0, perExc, catatanExc, catatan, aktif: dl.length > 0 };
  });
  const hariKerja = harian.filter(h => h.aktif).length;

  const rows = excs.map(e => {
    const m = ls.filter(l => l.excavator === e.kode);
    const vol = sum(m, 'vol');
    const hari = new Set(m.map(l => l.tanggal)).size;
    const targetPeriode = e.target * hariKerja;
    return { kode: e.kode, nama: e.nama, target: e.target, custom: e.custom, hari,
      vol: r1(vol), rit: sum(m, 'rit'), rata: hari ? r1(vol / hari) : 0,
      catatan: [...new Set(m.map(l => String(l.catatan || '').trim()).filter(Boolean))].join(' | '), targetPeriode,
      capai: targetPeriode > 0 ? r1(vol / targetPeriode * 100) : 0,
      hariCapai: harian.filter(h => h.perExc[e.kode] >= e.target).length };
  });

  const volTotal = sum(ls, 'vol'), targetTotal = targetHari * hariKerja;
  const kumulatif = excs.map(e => ({ kode: e.kode, nama: e.nama, vol: r1(sum(semua.filter(l => l.excavator === e.kode), 'vol')) }));
  const lokasiGalian = [...new Set(semua.map(l => l.lokasi).filter(Boolean))];
  return {
    days, harian, excavators: rows,
    ringkas: {
      vol: r1(volTotal), rit: sum(ls, 'rit'), target: targetTotal, hariKerja,
      capai: targetTotal > 0 ? r1(volTotal / targetTotal * 100) : 0,
      hariCapai: harian.filter(h => h.aktif && h.total >= h.target).length
    },
    tanpaExc: r1(sum(ls.filter(l => !kodeExc.has(l.excavator)), 'vol')),
    progres: progresTotal_(),
    kumulatif, lokasiGalian, terbaruInput: getGalian_().slice(-10).reverse()
  };
}

/** Rekap hasil galian per minggu atau bulan berdasarkan seluruh histori (input langsung + langsir). */
function getRekapGalian(mode) {
  mode = mode === 'bulan' ? 'bulan' : 'minggu';
  const excs = getAlat_().filter(a => jenisAlat_(a) === 'EXC')
    .map(a => ({ kode: a.kode, nama: a.nama, target: targetExc_(a) }));
  const kodeExc = new Set(excs.map(e => e.kode));
  const ls = galianRows_();
  const sum = (arr, k) => arr.reduce((t, x) => t + (Number(x[k]) || 0), 0);
  const r1 = n => +n.toFixed(1);
  const groups = {};
  ls.forEach(l => {
    if (!l.tanggal) return;
    const d = new Date(l.tanggal + 'T00:00:00');
    let key, label;
    if (mode === 'bulan') {
      key = l.tanggal.slice(0, 7);
      const [y, m] = key.split('-').map(Number);
      label = new Date(y, m - 1, 1).toLocaleDateString('id-ID', { month: 'long', year: 'numeric' });
    } else {
      const monday = new Date(d); monday.setDate(d.getDate() - ((d.getDay() + 6) % 7));
      const sunday = new Date(monday); sunday.setDate(monday.getDate() + 6);
      const fmt = x => Utilities.formatDate(x, Session.getScriptTimeZone(), 'yyyy-MM-dd');
      key = fmt(monday);
      label = fmt(monday) + ' s.d. ' + fmt(sunday);
    }
    if (!groups[key]) groups[key] = { key, label, dates: new Set(), data: [] };
    groups[key].dates.add(l.tanggal); groups[key].data.push(l);
  });
  const rows = Object.values(groups).sort((a, b) => b.key.localeCompare(a.key)).map(g => {
    const perExc = {};
    excs.forEach(e => perExc[e.kode] = r1(sum(g.data.filter(x => x.excavator === e.kode), 'vol')));
    const vol = r1(sum(g.data, 'vol'));
    const target = sum(excs, 'target') * g.dates.size;
    const catatan = [...new Set(g.data.map(x => String(x.catatan || '').trim()).filter(Boolean))].join(' | ');
    return { key: g.key, label: g.label, hari: g.dates.size, vol, rit: sum(g.data, 'rit'), target,
      capai: target > 0 ? r1(vol / target * 100) : 0, perExc, catatan,
      tanpaExc: r1(sum(g.data.filter(x => !kodeExc.has(x.excavator)), 'vol')) };
  });
  return { mode, excavators: excs, rows };
}

/** Ubah target galian per hari (m3) untuk satu excavator. */
function setTargetGalian(kode, target, days) {
  target = Number(target);
  if (!kode) throw new Error('Pilih excavator.');
  if (!(target > 0)) throw new Error('Target harus lebih dari 0.');
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const s = ss_().getSheetByName(SHEET_ALAT);
    if (!s.getRange(1, 5).getValue())
      s.getRange(1, 5).setValue(HEAD_ALAT[4]).setFontWeight('bold').setBackground('#f5c400');
    const v = s.getDataRange().getValues();
    for (let i = 1; i < v.length; i++) {
      if (v[i][0] === kode) { s.getRange(i + 1, 5).setValue(target); return getHasilGalian(days); }
    }
    throw new Error('Alat tidak ditemukan.');
  } finally { lock.releaseLock(); }
}

/** Data dashboard Ringkasan. days = 7 / 14 / 30. */
function getDashboard(days) {
  days = [7, 14, 30].indexOf(Number(days)) >= 0 ? Number(days) : 14;
  const tz = Session.getScriptTimeZone();
  const fmt = t => Utilities.formatDate(new Date(t), tz, 'yyyy-MM-dd');
  const r1 = n => +n.toFixed(1);
  const sum = (arr, k) => arr.reduce((t, x) => t + (Number(x[k]) || 0), 0);
  const today = fmt(Date.now());
  const alat = getAlat_(), logs = getLogs_(), ls = getLangsir_(), gal = galianRows_();

  const st = {};
  logs.filter(l => l.tanggal === today).forEach(l => st[l.kode] = l.status);
  const cnt = s => Object.values(st).filter(v => v === s).length;
  const ber = cnt('Beroperasi'), sby = cnt('Standby'), rsk = cnt('Rusak');

  const tg = [];
  for (let i = days - 1; i >= 0; i--) tg.push(fmt(Date.now() - i * 86400000));
  const harian = tg.map(t => {
    const L = logs.filter(x => x.tanggal === t);
    return { tanggal: t, bbm: r1(sum(L, 'bbm')), jam: r1(sum(L, 'jam')),
      gal: r1(sum(gal.filter(x => x.tanggal === t), 'vol')),
      lang: r1(sum(ls.filter(x => x.tanggal === t), 'vol')) };
  });

  const excs = alat.filter(a => jenisAlat_(a) === 'EXC');
  const exc = excs.map(a => {
    const v = sum(gal.filter(x => x.excavator === a.kode && x.tanggal === today), 'vol'), t = targetExc_(a);
    return { kode: a.kode, nama: a.nama, vol: r1(v), targetPeriode: t, capai: r1(v / t * 100) };
  });
  const target = sum(exc, 'targetPeriode'), volHari = r1(sum(exc, 'vol'));
  harian.forEach(h => h.target = target);

  const lsP = ls.filter(x => x.tanggal >= tg[0]);
  const grp = f => {
    const o = {};
    lsP.forEach(x => { const k = f(x), g = o[k] || (o[k] = { n: k, vol: 0, rit: 0 }); g.vol += x.vol; g.rit += x.rit; });
    return Object.values(o).map(g => ({ n: g.n, vol: r1(g.vol), rit: g.rit })).sort((a, b) => b.vol - a.vol).slice(0, 5);
  };
  const lsHari = ls.filter(x => x.tanggal === today);

  const prod = getProduktivitas(days);
  const prodAll = getProduktivitas(0);
  const top = prod.rows.filter(r => r.hari > 0).sort((a, b) => b.util - a.util).slice(0, 5);

  const unit = alat.map(a => {
    const L = logs.filter(l => l.kode === a.kode);
    return { kode: a.kode, nama: a.nama, jam: r1(sum(L, 'jam')), bbm: r1(sum(L, 'bbm')), status: st[a.kode] || '' };
  });

  return {
    days, harian, unit, targetRental: TARGET_RENTAL_JAM,
    alat: { total: alat.length, ber, sby, rsk, belum: Math.max(0, alat.length - Object.keys(st).length) },
    galian: { volHari, target, capai: target > 0 ? r1(volHari / target * 100) : 0, exc },
    langsir: { ritHari: sum(lsHari, 'rit'), volHari: r1(sum(lsHari, 'vol')),
      perDT: grp(x => x.kode), rute: grp(x => (x.asal || '?') + ' \u2192 ' + (x.tujuan || '?')) },
    progres: progresTotal_(ls),
    prod: { ringkas: prod.ringkas, top },
    total: { jam: r1(sum(logs, 'jam')), bbm: r1(sum(logs, 'bbm')), hari: new Set(logs.map(l => l.tanggal)).size,
      rit: sum(ls, 'rit'), prod: prodAll.ringkas },
    aktivitas: logs.slice(-6).reverse()
  };
}
