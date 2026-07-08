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
  kuran: 'Kuran Takip'
};

const TABLE_FIRST_ROW = 5;
const NAME_COL = 2;

function doGet() {
  return json_({ ok: true, message: 'Talebe Takip baglantisi hazir.' });
}

function doPost(e) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);

  try {
    const payload = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    const changes = Array.isArray(payload.changes) ? payload.changes : [];
    const results = changes.map(applyChange_);
    return json_({ ok: true, applied: results.length, results });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message ? err.message : err) });
  } finally {
    lock.releaseLock();
  }
}

function applyChange_(change) {
  if (!change || !change.type) return { ok: false, error: 'Bos degisiklik' };

  switch (change.type) {
    case 'att':
      return writeAttendance_(change.student, change.day, change.value);
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
    default:
      return { ok: false, type: change.type, error: 'Bilinmeyen degisiklik tipi' };
  }
}

function writeAttendance_(student, day, value) {
  // Yoklama sayfasinda gunler D sutunundan basliyor: day=0 -> D.
  return writeByStudent_(SHEETS.attendance, student, 4 + Number(day), value);
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
    if (String(values[i] || '').trim() === clean) return firstCol + i;
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

function findStudentRow_(sheet, student, nameCol, firstRow) {
  const needle = normalizeName_(student);
  if (!needle) return null;

  const lastRow = sheet.getLastRow();
  if (lastRow < firstRow) return null;

  const values = sheet.getRange(firstRow, nameCol, lastRow - firstRow + 1, 1).getValues();
  for (let i = 0; i < values.length; i++) {
    if (normalizeName_(values[i][0]) === needle) return firstRow + i;
  }
  return null;
}

function normalizeName_(value) {
  return String(value || '')
    .trim()
    .replace(/\s+/g, ' ')
    .replace(/İ/g, 'i')
    .replace(/I/g, 'ı')
    .toLowerCase();
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
