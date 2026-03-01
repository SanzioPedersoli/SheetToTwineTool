const COL_QUEST = 'Quest';
const COL_ID = 'ID';
const COL_TRIGGER = 'Trigger';
const COL_CHARACTER = 'Character';
const COL_DIALOG = 'Dialog';
const COL_JUMPTO = 'JumpTo';
const COL_NOTES = 'Artist Notes';
const COL_MACROS = 'Designer Notes';

function getSheetValues() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getActiveSheet();
  return sheet.getDataRange().getValues();
}

function validateData(values) {
  if (!values || values.length < 2) throw new Error('No data found.');
}

function getHeaderMap(headers) {
  const col = {};
  headers.forEach((h, i) => col[h] = i);
  return col;
}

function readCell_(row, index, fallback = '') {
  if (index == null) return fallback;
  return row[index] ?? fallback;
}

function isBlank_(v) {
  return v === '' || v === null || v === undefined;
}

function saveToDrive(output, fileName) {
  const folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
  const file = folder.createFile(fileName, output, MimeType.PLAIN_TEXT);
  Logger.log('Created file: ' + file.getUrl());
  return file.getUrl();
}