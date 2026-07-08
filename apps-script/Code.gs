/*
 * Google Sheets sync endpoint for Talebe Takip.html.
 *
 * Kurulum:
 * 1. Google E-Tabloyu acin.
 * 2. Uzantilar > Apps Script menusuyle yeni proje acin.
 * 3. Bu dosyanin tamamini Code.gs icine yapistirin.
 * 4. Deploy > New deployment > Web app:
 *    - Execute as: Me
 *    - Who has access: Anyone with the link
 * 5. Olusan /exec adresini sitedeki Ayarlar > E-Tablo baglantisi alanina yapistirin.
 */

const SHEETS = {
  attendance: 'Yoklama',
  nurlu: ['Nurlu K.(1-5)', 'Nurlu K.(6-10)', 'Nurlu K.(11-15)', 'Nurlu K.(16-20)'],
  sure: 'Ezber Takip',
  namaz: 'Namaz Takip',
  elifba: 'Elif-Ba Takip',
  kuran: 'Kuran Takip',
  hocalar: 'Hocalar',
  gecmis: 'İşlem Geçmişi'
};

const LOG_READ_LIMIT = 200;

const TABLE_FIRST_ROW = 5;
const NAME_COL = 2;
const AYLAR = ['Oca','Şub','Mar','Nis','May','Haz','Tem','Ağu','Eyl','Eki','Kas','Ara'];
const TZ = 'Europe/Istanbul';

function doGet() {
  return json_({ ok: true, message: 'Talebe Takip baglantisi hazir.' });
}

function doPost(e) {
  try {
    const payload = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const changes = Array.isArray(payload.changes) ? payload.changes : [];

    try { ensureTodayColumn_(); } catch (err) {}

    var readOnly = changes.every(function(c) {
      return c && (c.type === 'readAttendance' || c.type === 'readAllAttendance' || c.type === 'readStudent' || c.type === 'readVersion');
    });

    var lock = null;
    if (!readOnly) {
      lock = LockService.getDocumentLock();
      lock.waitLock(30000);
    }

    try {
      const results = changes.map(applyChange_);
      if (!readOnly) {
        try { appendLogs_(changes); } catch (errLog) {}
      }
      var version = readOnly ? getVersion_() : bumpVersion_();
      return json_({ ok: true, applied: results.length, results: results, version: version });
    } finally {
      if (lock) lock.releaseLock();
    }
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  }
}

function applyChange_(change) {
  if (!change || !change.type) return { ok: false, error: 'Bos degisiklik' };

  switch (change.type) {
    case 'att':
      return writeAttendance_(change.student, change.day, change.value, change.date);
    case 'nurlu':
      return writeNurlu_(change.student, change.card, change.item, change.value);
    case 'sure':
      return writeByStudent_(SHEETS.sure, change.student, 3 + Number(change.index), change.value);
    case 'elifba':
      return writeElifba_(change.student, change.value);
    case 'namaz':
      return writeByStudent_(SHEETS.namaz, change.student, 3, change.value);
    case 'namazDaily':
      return writeNamazDaily_(change.student, change.date, change.value);
    case 'addStudent':
      return addStudentEverywhere_(change.student);
    case 'removeStudent':
      return clearStudentEverywhere_(change.student);
    case 'readAttendance':
      return readAttendance_(change.date);
    case 'readAllAttendance':
      return readAllAttendance_();
    case 'readStudent':
      return readStudent_(change.student);
    case 'addHoca':
      return addHoca_(change.name);
    case 'logBulk':
      return logBulk_(change.entries);
    case 'readVersion':
      return { ok: true, type: 'readVersion', version: getVersion_() };
    default:
      return { ok: false, type: change.type, error: 'Bilinmeyen degisiklik tipi' };
  }
}

function writeAttendance_(student, day, value, date) {
  var sheet = getSheet_(SHEETS.attendance);
  var col = date
    ? findOrAppendHeader_(sheet, date, 4, 2)
    : 4 + Number(day);
  var row = findOrAppendStudent_(sheet, student, NAME_COL, TABLE_FIRST_ROW);
  sheet.getRange(row, col).setValue(value || '');
  return { ok: true, sheet: SHEETS.attendance, row: row, col: col };
}

function writeNurlu_(student, card, item, value) {
  const cardNo = Number(card);
  const itemNo = Number(item);
  const sheetName = SHEETS.nurlu[Math.floor((cardNo - 1) / 5)];
  const cardOffset = (cardNo - 1) % 5;
  const firstCol = 3 + cardOffset * 4;

  // Hem site hem Excel sirasi: Vecize, Dua/Sure, Ilmihal, Kelime.
  const uiToSheetOffset = [0, 1, 2, 3];
  return writeByStudent_(sheetName, student, firstCol + uiToSheetOffset[itemNo], value);
}

function writeElifba_(student, value) {
  const sheet = getSheet_(SHEETS.elifba);
  const row = findOrAppendStudent_(sheet, student, 1, 4);
  sheet.getRange(row, 2).setValue(value || '');
  return { ok: true, sheet: SHEETS.elifba, row, col: 2 };
}

function writeNamazDaily_(student, date, value) {
  const sheet = getSheet_(SHEETS.namaz);
  const row = findOrAppendStudent_(sheet, student, NAME_COL, TABLE_FIRST_ROW);
  const col = findOrAppendHeader_(sheet, date || '', 3, TABLE_FIRST_ROW - 1);
  const text = Array.isArray(value) ? value.join(', ') : (value || '');
  sheet.getRange(row, col).setValue(text);
  return { ok: true, sheet: SHEETS.namaz, row, col, date };
}

function writeByStudent_(sheetName, student, col, value) {
  const sheet = getSheet_(sheetName);
  const row = findOrAppendStudent_(sheet, student, NAME_COL, TABLE_FIRST_ROW);
  sheet.getRange(row, col).setValue(value || '');
  return { ok: true, sheet: sheetName, row, col };
}

function addStudentEverywhere_(student) {
  if (!student) return { ok: false, error: 'Ogrenci adi bos' };
  [SHEETS.attendance, SHEETS.sure, SHEETS.namaz, SHEETS.kuran].concat(SHEETS.nurlu).forEach(function(sheetName) {
    const sheet = getSheet_(sheetName);
    findOrAppendStudent_(sheet, student, NAME_COL, TABLE_FIRST_ROW);
  });
  findOrAppendStudent_(getSheet_(SHEETS.elifba), student, 1, 4);
  return { ok: true, type: 'addStudent', student };
}

function clearStudentEverywhere_(student) {
  if (!student) return { ok: false, error: 'Ogrenci adi bos' };
  [SHEETS.attendance, SHEETS.sure, SHEETS.namaz, SHEETS.kuran].concat(SHEETS.nurlu).forEach(function(sheetName) {
    clearStudentRow_(getSheet_(sheetName), student, NAME_COL, TABLE_FIRST_ROW);
  });
  clearStudentRow_(getSheet_(SHEETS.elifba), student, 1, 4);
  return { ok: true, type: 'removeStudent', student };
}

function clearStudentRow_(sheet, student, nameCol, firstRow) {
  const row = findStudentRow_(sheet, student, nameCol, firstRow);
  if (!row) return;
  sheet.getRange(row, 1, 1, sheet.getLastColumn()).clearContent();
}

function findOrAppendStudent_(sheet, student, nameCol, firstRow) {
  const found = findStudentRow_(sheet, student, nameCol, firstRow);
  if (found) return found;

  const lastRow = Math.max(sheet.getLastRow(), firstRow);
  const values = sheet.getRange(firstRow, nameCol, Math.max(1, lastRow - firstRow + 1), 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (!String(values[i][0] || '').trim()) {
      sheet.getRange(firstRow + i, nameCol).setValue(student);
      if (nameCol === NAME_COL) sheet.getRange(firstRow + i, 1).setValue(i + 1);
      return firstRow + i;
    }
  }

  const row = lastRow + 1;
  sheet.getRange(row, nameCol).setValue(student);
  if (nameCol === NAME_COL) sheet.getRange(row, 1).setValue(row - firstRow + 1);
  return row;
}

function findOrAppendHeader_(sheet, label, firstCol, headerRow) {
  const clean = String(label || '').trim();
  if (!clean) return firstCol;

  const lastCol = Math.max(sheet.getLastColumn(), firstCol);
  const values = sheet.getRange(headerRow, firstCol, 1, Math.max(1, lastCol - firstCol + 1)).getValues()[0];
  for (let i = 0; i < values.length; i++) {
    if (normalizeHeader_(values[i]) === clean) return firstCol + i;
  }
  for (let i = 0; i < values.length; i++) {
    if (!String(values[i] || '').trim()) {
      sheet.getRange(headerRow, firstCol + i).setValue(clean);
      return firstCol + i;
    }
  }
  const col = lastCol + 1;
  sheet.getRange(headerRow, col).setValue(clean);
  return col;
}

var ROSTER_MEMO = null;

function rosterNorms_() {
  if (ROSTER_MEMO) return ROSTER_MEMO;
  var memo = {};
  try {
    var sheet = getSheet_(SHEETS.attendance);
    var lastRow = sheet.getLastRow();
    if (lastRow >= TABLE_FIRST_ROW) {
      var vals = sheet.getRange(TABLE_FIRST_ROW, NAME_COL, lastRow - TABLE_FIRST_ROW + 1, 1).getValues();
      for (var i = 0; i < vals.length; i++) {
        var n = normalizeName_(vals[i][0]);
        if (n) memo[n] = true;
      }
    }
  } catch (e) {}
  ROSTER_MEMO = memo;
  return memo;
}

function findStudentRow_(sheet, student, nameCol, firstRow) {
  const needle = normalizeName_(student);
  if (!needle) return null;

  const lastRow = sheet.getLastRow();
  if (lastRow < firstRow) return null;

  const values = sheet.getRange(firstRow, nameCol, lastRow - firstRow + 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (normalizeName_(values[i][0]) === needle) return firstRow + i;
  }

  // Birebir eşleşme yok: kelime bazlı ön ek eşleşmesi dene.
  // "Abdullah Altun" ↔ "Abdullah", "Ömer" ↔ "Ömer İnal" gibi.
  // Satırdaki isim, yoklama listesindeki BAŞKA bir öğrencinin tam adıysa atlanır
  // (iki ayrı "Yiğit" / "Yiğit Hamza" öğrencisi karışmasın diye).
  const roster = rosterNorms_();
  let hit = null, count = 0;
  for (let i = 0; i < values.length; i++) {
    const nm = normalizeName_(values[i][0]);
    if (!nm || nm === needle) continue;
    const isPrefix = needle.indexOf(nm + ' ') === 0 || nm.indexOf(needle + ' ') === 0;
    if (!isPrefix) continue;
    if (roster[nm]) continue; // satır adı başka kayıtlı öğrencinin tam adı
    hit = firstRow + i;
    count++;
  }
  return count === 1 ? hit : null;
}

function normalizeName_(value) {
  return String(value || '')
    .trim()
    .replace(/\s+/g, ' ')
    .replace(/İ/g, 'i')
    .replace(/I/g, 'ı')
    .toLowerCase();
}

function normalizeHeader_(h) {
  if (h instanceof Date && !isNaN(h.getTime())) return dateToLabel_(h);
  var s = String(h || '').trim();
  var m = s.match(/^(\d{1,2})[./](\d{1,2})[./](\d{4})$/);
  if (m) {
    var d = new Date(Number(m[3]), Number(m[2]) - 1, Number(m[1]));
    if (!isNaN(d.getTime())) return dateToLabel_(d);
  }
  return s;
}

function dateToLabel_(d) {
  return d.getDate() + ' ' + AYLAR[d.getMonth()];
}

function readAttendance_(date) {
  var sheet = getSheet_(SHEETS.attendance);
  var headerRow = 2;
  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();

  var allStudents = [];
  if (lastRow >= TABLE_FIRST_ROW) {
    var nameVals = sheet.getRange(TABLE_FIRST_ROW, NAME_COL, lastRow - TABLE_FIRST_ROW + 1, 1).getValues();
    for (var k = 0; k < nameVals.length; k++) {
      var n = String(nameVals[k][0] || '').trim();
      if (n) allStudents.push(n);
    }
  }

  if (lastCol < 4) return { ok: true, type: 'readAttendance', date: date, data: {}, allStudents: allStudents };

  var headers = sheet.getRange(headerRow, 4, 1, lastCol - 3).getValues()[0];
  var col = -1;
  var clean = String(date || '').trim();
  for (var i = 0; i < headers.length; i++) {
    if (normalizeHeader_(headers[i]) === clean) { col = 4 + i; break; }
  }
  if (col < 0) return { ok: true, type: 'readAttendance', date: date, data: {}, allStudents: allStudents };

  if (lastRow < TABLE_FIRST_ROW) return { ok: true, type: 'readAttendance', date: date, data: {}, allStudents: allStudents };

  var rows = lastRow - TABLE_FIRST_ROW + 1;
  var values = sheet.getRange(TABLE_FIRST_ROW, col, rows, 1).getValues();

  var data = {};
  for (var i = 0; i < allStudents.length; i++) {
    var val = String(values[i][0] || '').trim();
    if (val) data[allStudents[i]] = val;
  }
  return { ok: true, type: 'readAttendance', date: date, data: data, allStudents: allStudents };
}

function todayLabel_() {
  var parts = Utilities.formatDate(new Date(), TZ, 'd:M:u').split(':');
  var isoDay = Number(parts[2]);
  if (isoDay >= 6) return null; // cumartesi/pazar: ders yok
  return Number(parts[0]) + ' ' + AYLAR[Number(parts[1]) - 1];
}

function ensureTodayColumn_() {
  var label = todayLabel_();
  if (!label) return;
  var cache = CacheService.getScriptCache();
  var key = 'day_' + label;
  if (cache.get(key)) return;
  var lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    findOrAppendHeader_(getSheet_(SHEETS.attendance), label, 4, 2);
  } finally {
    lock.releaseLock();
  }
  cache.put(key, '1', 21600);
  bumpVersion_();
}

function isDateLabel_(s) {
  var m = String(s || '').match(/^(\d{1,2}) (\S+)$/);
  return !!(m && Number(m[1]) >= 1 && Number(m[1]) <= 31 && AYLAR.indexOf(m[2]) >= 0);
}

function getOrCreateSheet_(name, headers) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    try {
      sheet = ss.insertSheet(name);
      if (headers && headers.length) sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    } catch (e) {
      sheet = ss.getSheetByName(name);
      if (!sheet) throw e;
    }
  }
  return sheet;
}

function addHoca_(name) {
  var clean = String(name || '').trim();
  if (!clean) return { ok: false, error: 'Hoca adi bos' };
  var sheet = getOrCreateSheet_(SHEETS.hocalar, ['Hoca']);
  var lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    var vals = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < vals.length; i++) {
      if (normalizeName_(vals[i][0]) === normalizeName_(clean)) return { ok: true, type: 'addHoca', name: clean, existed: true };
    }
  }
  sheet.getRange(lastRow + 1, 1).setValue(clean);
  return { ok: true, type: 'addHoca', name: clean };
}

function appendLogs_(changes) {
  var rows = [];
  for (var i = 0; i < changes.length; i++) {
    var c = changes[i];
    if (!c || !c.text || c.type === 'logBulk') continue;
    var ts = Number(c.ts) || Date.now();
    rows.push([new Date(ts), String(c.hoca || ''), String(c.text)]);
  }
  if (!rows.length) return;
  var sheet = getOrCreateSheet_(SHEETS.gecmis, ['Tarih', 'Hoca', 'İşlem']);
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 3).setValues(rows);
}

function logBulk_(entries) {
  if (!Array.isArray(entries) || !entries.length) return { ok: false, error: 'Bos liste' };
  var rows = entries.slice(-400).map(function(e) {
    var ts = Number(e && e[0]) || Date.now();
    return [new Date(ts), String((e && e[1]) || ''), String((e && e[2]) || '')];
  }).filter(function(r) { return r[2]; });
  if (!rows.length) return { ok: false, error: 'Bos liste' };
  var sheet = getOrCreateSheet_(SHEETS.gecmis, ['Tarih', 'Hoca', 'İşlem']);
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 3).setValues(rows);
  return { ok: true, type: 'logBulk', added: rows.length };
}

function readMeta_() {
  var meta = { hocalar: [], log: [] };
  var ss = SpreadsheetApp.getActive();

  var hs = ss.getSheetByName(SHEETS.hocalar);
  if (hs && hs.getLastRow() >= 2) {
    var hv = hs.getRange(2, 1, hs.getLastRow() - 1, 1).getValues();
    for (var i = 0; i < hv.length; i++) {
      var n = String(hv[i][0] || '').trim();
      if (n) meta.hocalar.push(n);
    }
  }

  var ls = ss.getSheetByName(SHEETS.gecmis);
  if (ls && ls.getLastRow() >= 2) {
    var total = ls.getLastRow() - 1;
    var count = Math.min(total, LOG_READ_LIMIT);
    var lv = ls.getRange(ls.getLastRow() - count + 1, 1, count, 3).getValues();
    for (var j = 0; j < lv.length; j++) {
      var ts = lv[j][0] instanceof Date ? lv[j][0].getTime() : Number(lv[j][0]) || 0;
      var text = String(lv[j][2] || '').trim();
      if (text) meta.log.push([ts, String(lv[j][1] || ''), text]);
    }
  }
  return meta;
}

function getVersion_() {
  var cached = CacheService.getScriptCache().get('v');
  if (cached != null) return Number(cached);
  var v = PropertiesService.getScriptProperties().getProperty('v') || '0';
  CacheService.getScriptCache().put('v', v, 21600);
  return Number(v);
}

function bumpVersion_() {
  var v = getVersion_() + 1;
  PropertiesService.getScriptProperties().setProperty('v', String(v));
  CacheService.getScriptCache().put('v', String(v), 21600);
  return v;
}

function readAllAttendance_() {
  var version = getVersion_();
  var cache = CacheService.getScriptCache();
  var cacheKey = 'att_' + version;
  var hit = cache.get(cacheKey);
  if (hit) {
    try { return JSON.parse(hit); } catch (e) {}
  }

  var sheet = getSheet_(SHEETS.attendance);
  var lastCol = sheet.getLastColumn();
  var lastRow = sheet.getLastRow();
  var rows = lastRow >= TABLE_FIRST_ROW ? lastRow - TABLE_FIRST_ROW + 1 : 0;

  var allStudents = [];
  var dates = [];
  var attendance = {};

  if (rows > 0 && lastCol >= NAME_COL) {
    var block = sheet.getRange(TABLE_FIRST_ROW, NAME_COL, rows, lastCol - NAME_COL + 1).getValues();
    for (var k = 0; k < block.length; k++) {
      var n = String(block[k][0] || '').trim();
      if (n) allStudents.push(n);
    }

    if (lastCol >= 4) {
      var headers = sheet.getRange(2, 4, 1, lastCol - 3).getValues()[0];
      var dataOffset = 4 - NAME_COL;
      for (var c = 0; c < headers.length; c++) {
        var label = normalizeHeader_(headers[c]);
        if (!label || !isDateLabel_(label)) continue; // özet kolonları (A–Z, Toplam vb.) tarih değil
        dates.push(label);
        for (var r = 0; r < allStudents.length; r++) {
          var v = String(block[r][dataOffset + c] || '').trim();
          if (v) {
            if (!attendance[allStudents[r]]) attendance[allStudents[r]] = {};
            attendance[allStudents[r]][label] = v;
          }
        }
      }
    }
  }

  var meta = readMeta_();
  var result = { ok: true, type: 'readAllAttendance', allStudents: allStudents, dates: dates, attendance: attendance, version: version, hocalar: meta.hocalar, log: meta.log };
  try { cache.put(cacheKey, JSON.stringify(result), 120); } catch (e) {}
  return result;
}

function readStudent_(student) {
  if (!student) return { ok: false, error: 'Ogrenci adi bos' };
  var cache = CacheService.getScriptCache();
  var cacheKey = 'stu_' + getVersion_() + '_' + Utilities.base64Encode(student, Utilities.Charset.UTF_8);
  var hit = cache.get(cacheKey);
  if (hit) {
    try { return JSON.parse(hit); } catch (e) {}
  }
  var result = { ok: true, type: 'readStudent', student: student, nurlu: {}, sure: {}, elifba: '', namaz: 0 };

  // Nurlu cards (4 sheets, 5 cards each, 4 items per card)
  for (var si = 0; si < SHEETS.nurlu.length; si++) {
    var sheet = getSheet_(SHEETS.nurlu[si]);
    var row = findStudentRow_(sheet, student, NAME_COL, TABLE_FIRST_ROW);
    if (!row) continue;
    var lastCol = sheet.getLastColumn();
    if (lastCol < 3) continue;
    var vals = sheet.getRange(row, 3, 1, lastCol - 2).getValues()[0];
    for (var ci = 0; ci < 5; ci++) {
      var cardNo = si * 5 + ci + 1;
      for (var ii = 0; ii < 4; ii++) {
        var colIdx = ci * 4 + ii;
        var v = String(vals[colIdx] || '').trim();
        if (v) result.nurlu['c' + cardNo + '_' + ii] = v;
      }
    }
  }

  // Sure (Ezber Takip)
  var sureSheet = getSheet_(SHEETS.sure);
  var sureRow = findStudentRow_(sureSheet, student, NAME_COL, TABLE_FIRST_ROW);
  if (sureRow) {
    var sureLast = sureSheet.getLastColumn();
    if (sureLast >= 3) {
      var sureVals = sureSheet.getRange(sureRow, 3, 1, sureLast - 2).getValues()[0];
      for (var i = 0; i < sureVals.length; i++) {
        var sv = String(sureVals[i] || '').trim();
        if (sv) result.sure[i] = sv;
      }
    }
  }

  // Elif-Ba
  var elifbaSheet = getSheet_(SHEETS.elifba);
  var elifbaRow = findStudentRow_(elifbaSheet, student, 1, 4);
  if (elifbaRow) {
    result.elifba = String(elifbaSheet.getRange(elifbaRow, 2).getValue() || '').trim();
  }

  // Namaz
  var namazSheet = getSheet_(SHEETS.namaz);
  var namazRow = findStudentRow_(namazSheet, student, NAME_COL, TABLE_FIRST_ROW);
  if (namazRow) {
    result.namaz = Number(namazSheet.getRange(namazRow, 3).getValue()) || 0;
  }

  try { cache.put(cacheKey, JSON.stringify(result), 120); } catch (e) {}
  return result;
}

function getSheet_(name) {
  const sheet = SpreadsheetApp.getActive().getSheetByName(name);
  if (!sheet) throw new Error('Sayfa bulunamadi: ' + name);
  return sheet;
}

function json_(body) {
  return ContentService
    .createTextOutput(JSON.stringify(body))
    .setMimeType(ContentService.MimeType.JSON);
}
