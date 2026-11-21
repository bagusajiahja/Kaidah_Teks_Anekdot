/**
 * Penerima data "Belajar Kaidah Anekdot" (Google Apps Script).
 *
 * Cara pasang:
 * 1. Buat Google Sheet baru. Menu Ekstensi > Apps Script. Tempel seluruh kode ini.
 * 2. Ganti TOKEN di bawah (samakan dengan isian Token di Pengaturan aplikasi).
 * 3. Jalankan fungsi setup() satu kali (beri izin) untuk membuat lembar Butir, Jawaban, Rekap_TP, Dasbor.
 * 4. Deploy > Deployment baru > Jenis: Aplikasi web. Jalankan sebagai: Saya. Yang memiliki akses: Siapa saja.
 * 5. Salin URL yang berakhiran /exec ke Pengaturan aplikasi.
 * Setiap mengubah kode, buat versi deployment baru.
 */
var TOKEN = 'ganti-token-rahasia';
var KKTP = 75; // ambang ketercapaian (%), samakan dengan KKTP satuan pendidikan

var BUTIR = [
  ['F1', 'Formatif', 1, 'C3', 'Isian', 2, 'O+K'], ['F2', 'Formatif', 1, 'C3', 'Isian', 2, 'O+K'], ['F3', 'Formatif', 1, 'C3', 'Isian', 2, 'O+K'],
  ['F4', 'Formatif', 1, 'C4', 'Pilihan + alasan', 2, 'O+K'], ['F5', 'Formatif', 1, 'C4', 'Pilihan + alasan', 2, 'O+K'], ['F6', 'Formatif', 1, 'C4', 'Pilihan + alasan', 2, 'O+K'],
  ['F7', 'Formatif', 1, 'C4', 'Isian', 2, 'O+K'], ['F8', 'Formatif', 1, 'C4', 'Isian 3 bagian', 2, 'O'],
  ['F9', 'Formatif', 1, 'C4', 'Uraian', 4, 'M'], ['F10', 'Formatif', 1, 'C4', 'Uraian', 4, 'M'],
  ['L2-5', 'Formatif (LK-2)', 3, 'C5', 'Uraian', 4, 'M'],
  ['S1', 'Sumatif', 2, 'C4', 'Uraian', 4, 'M'], ['S2', 'Sumatif', 2, 'C4', 'Uraian', 4, 'M'], ['S3', 'Sumatif', 2, 'C4', 'Uraian', 4, 'M'],
  ['S4', 'Sumatif', 2, 'C4', 'Uraian', 4, 'M'], ['S5', 'Sumatif', 2, 'C4', 'Uraian', 4, 'M'], ['S6', 'Sumatif', 3, 'C5', 'Uraian', 4, 'M']
];
var H_JAWAB = ['Kode', 'ID', 'Jenis', 'TP', 'Level', 'Skor', 'Maks', 'Komponen', 'Jawaban', 'Diperbarui'];
var H_REKAP = ['Kode', 'Nama', 'Jalur', 'TP1_Skor', 'TP1_%', 'TP1_Kategori', 'TP2_Skor', 'TP2_%', 'TP2_Kategori', 'TP3_Skor', 'TP3_%', 'TP3_Kategori', 'Diperbarui'];
var KATEGORI = ['Mahir', 'Cakap', 'Mulai berkembang', 'Belum berkembang'];

function setup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var b = sheet_(ss, 'Butir');
  b.clear(); b.getRange(1, 1, 1, 7).setValues([['ID', 'Jenis', 'TP', 'Level', 'Bentuk', 'Skor maks', 'Mode skor']]).setFontWeight('bold');
  b.getRange(2, 1, BUTIR.length, 7).setValues(BUTIR); b.setFrozenRows(1);
  var j = sheet_(ss, 'Jawaban');
  if (j.getLastRow() === 0) { j.getRange(1, 1, 1, H_JAWAB.length).setValues([H_JAWAB]).setFontWeight('bold'); j.setFrozenRows(1); }
  var r = sheet_(ss, 'Rekap_TP');
  if (r.getLastRow() === 0) { r.getRange(1, 1, 1, H_REKAP.length).setValues([H_REKAP]).setFontWeight('bold'); r.setFrozenRows(1); }
  buildDasbor_(ss);
}

function buildDasbor_(ss) {
  var d = sheet_(ss, 'Dasbor'); d.clear();
  d.getRange('A1').setValue('Dasbor ketercapaian: Belajar Kaidah Anekdot').setFontWeight('bold').setFontSize(14);
  d.getRange('A3:B3').setValues([['Jumlah peserta', '=COUNTA(Rekap_TP!A2:A)']]);
  d.getRange('A4:B4').setValues([['KKTP (%)', KKTP]]);
  d.getRange('A6:D6').setValues([['Ketercapaian per TP', 'TP1', 'TP2', 'TP3']]).setFontWeight('bold');
  d.getRange('A7:D7').setValues([['Rata-rata (%)', '=IFERROR(ROUND(AVERAGE(Rekap_TP!E2:E),1),"")', '=IFERROR(ROUND(AVERAGE(Rekap_TP!H2:H),1),"")', '=IFERROR(ROUND(AVERAGE(Rekap_TP!K2:K),1),"")']]);
  d.getRange('A8:D8').setValues([['Tercapai ≥ KKTP (%)', '=IFERROR(ROUND(COUNTIF(Rekap_TP!E2:E,">="&$B$4)/COUNT(Rekap_TP!E2:E)*100,1),"")', '=IFERROR(ROUND(COUNTIF(Rekap_TP!H2:H,">="&$B$4)/COUNT(Rekap_TP!H2:H)*100,1),"")', '=IFERROR(ROUND(COUNTIF(Rekap_TP!K2:K,">="&$B$4)/COUNT(Rekap_TP!K2:K)*100,1),"")']]);
  d.getRange('A10:D10').setValues([['Sebaran kategori', 'TP1', 'TP2', 'TP3']]).setFontWeight('bold');
  for (var i = 0; i < KATEGORI.length; i++) {
    var row = 11 + i;
    d.getRange(row, 1, 1, 4).setValues([[KATEGORI[i], '=COUNTIF(Rekap_TP!F2:F,$A' + row + ')', '=COUNTIF(Rekap_TP!I2:I,$A' + row + ')', '=COUNTIF(Rekap_TP!L2:L,$A' + row + ')']]);
  }
  d.getRange('A16:D16').setValues([['Analisis butir', 'TP', 'Skor maks', 'Rata-rata skor / maks (%)']]).setFontWeight('bold');
  for (var k = 0; k < BUTIR.length; k++) {
    var r = 17 + k, br = 2 + k;
    d.getRange(r, 1, 1, 4).setValues([['=Butir!A' + br, '=Butir!C' + br, '=Butir!F' + br, '=IFERROR(ROUND(AVERAGEIF(Jawaban!$B:$B,A' + r + ',Jawaban!$F:$F)/C' + r + '*100,1),"")']]);
  }
  d.setColumnWidth(1, 220);
}

function sheet_(ss, name) { return ss.getSheetByName(name) || ss.insertSheet(name); }
function clean_(v) { var s = String(v == null ? '' : v); return /^[=+\-@]/.test(s) ? "'" + s : s; }
function kategori_(p) { return p >= 90 ? 'Mahir' : p >= 75 ? 'Cakap' : p >= 50 ? 'Mulai berkembang' : 'Belum berkembang'; }
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function doGet() { return out_({ ok: true, app: 'kaidah-anekdot-v1' }); }

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    var p = JSON.parse(e.postData.contents);
    if (p.app !== 'kaidah-anekdot-v1' || p.token !== TOKEN) return out_({ ok: false, error: 'token' });
    var kode = clean_(String(p.kode || '').trim().slice(0, 24));
    if (!kode) return out_({ ok: false, error: 'kode' });
    lock.waitLock(20000);
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss.getSheetByName('Jawaban')) setup();
    var now = new Date();

    // Jawaban: upsert per Kode + ID
    var sh = ss.getSheetByName('Jawaban'), data = sh.getDataRange().getValues(), idx = {};
    for (var i = 1; i < data.length; i++) idx[data[i][0] + '|' + data[i][1]] = i + 1;
    var append = [];
    (p.items || []).slice(0, 40).forEach(function (it) {
      var row = [kode, clean_(it.id), clean_(it.jenis), Number(it.tp) || '', clean_(it.level), Number(it.skor) || 0, Number(it.maks) || 0, clean_(it.komponen), clean_(String(it.jawaban || '').slice(0, 5000)), now];
      var key = kode + '|' + it.id;
      if (idx[key]) sh.getRange(idx[key], 1, 1, row.length).setValues([row]); else append.push(row);
    });
    if (append.length) sh.getRange(sh.getLastRow() + 1, 1, append.length, H_JAWAB.length).setValues(append);

    // Rekap_TP: upsert per Kode
    var rk = ss.getSheetByName('Rekap_TP'), rd = rk.getDataRange().getValues(), at = 0;
    for (var r = 1; r < rd.length; r++) { if (String(rd[r][0]) === kode) { at = r + 1; break; } }
    var tp = {}; (p.tp || []).forEach(function (t) { tp[t.tp] = t; });
    function g(n, f) { return tp[n] ? tp[n][f] : ''; }
    var line = [kode, clean_(String(p.nama || '').slice(0, 40)), clean_(p.jalur), g(1, 'skor'), g(1, 'persen'), tp[1] && tp[1].persen !== '' ? kategori_(tp[1].persen) : '', g(2, 'skor'), g(2, 'persen'), tp[2] ? kategori_(tp[2].persen) : '', g(3, 'skor'), g(3, 'persen'), tp[3] ? kategori_(tp[3].persen) : '', now];
    if (at) rk.getRange(at, 1, 1, line.length).setValues([line]); else rk.appendRow(line);
    return out_({ ok: true, items: (p.items || []).length });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) { }
  }
}
