const BOOK_DOCUMENT_SCHEMA_VERSION_ = 1;
const BOOK_EDITOR_COLUMNS_ = 7;
const BOOK_EDITOR_MAX_ROWS_ = 5000;

function getBookDocument(sheetName) {
  const sheet = getTargetSheet_(sheetName);
  const lastRow = sheet.getLastRow();
  const rows = lastRow > 0
    ? sheet.getRange(1, 1, lastRow, BOOK_EDITOR_COLUMNS_).getDisplayValues()
    : [];

  return {
    schemaVersion: BOOK_DOCUMENT_SCHEMA_VERSION_,
    id: sheet.getSheetId(),
    name: sheet.getName(),
    columns: ['question', 'answer', 'studyCount', 'rank', 'comment1', 'comment2', 'comment3'],
    rows: rows.filter((row) => row.some((value) => String(value || '').trim())),
    updatedAt: new Date().toISOString(),
    storage: 'google-sheets',
  };
}

function saveBookDocument(rawDocument) {
  const document = normalizeBookDocument_(rawDocument);
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);

  try {
    const sheet = getTargetSheet_(document.name);
    const existingRows = Math.max(sheet.getLastRow(), 1);
    sheet.getRange(1, 1, existingRows, BOOK_EDITOR_COLUMNS_).clearContent();

    if (document.rows.length > 0) {
      sheet
        .getRange(1, 1, document.rows.length, BOOK_EDITOR_COLUMNS_)
        .setValues(document.rows);
    }

    SpreadsheetApp.flush();
    return {
      ok: true,
      name: sheet.getName(),
      rowCount: document.rows.length,
      schemaVersion: BOOK_DOCUMENT_SCHEMA_VERSION_,
      updatedAt: new Date().toISOString(),
      storage: 'google-sheets',
    };
  } finally {
    lock.releaseLock();
  }
}

function normalizeBookDocument_(rawDocument) {
  if (!rawDocument || typeof rawDocument !== 'object') {
    throw new Error('Book document is required.');
  }

  const name = String(rawDocument.name || '').trim();
  if (!name) {
    throw new Error('Book name is required.');
  }

  const sourceRows = Array.isArray(rawDocument.rows) ? rawDocument.rows : [];
  if (sourceRows.length > BOOK_EDITOR_MAX_ROWS_) {
    throw new Error(`A book can contain up to ${BOOK_EDITOR_MAX_ROWS_} rows.`);
  }

  const rows = sourceRows
    .map((rawRow) => {
      const row = Array.isArray(rawRow) ? rawRow : [];
      return Array.from({ length: BOOK_EDITOR_COLUMNS_ }, (_, index) =>
        sanitizeBookCell_(row[index])
      );
    })
    .filter((row) => row.some((value) => value !== ''));

  return {
    schemaVersion: BOOK_DOCUMENT_SCHEMA_VERSION_,
    name,
    rows,
  };
}

function sanitizeBookCell_(value) {
  const text = String(value === null || value === undefined ? '' : value)
    .replace(/\r\n?/g, '\n')
    .slice(0, 20000);

  // Prevent pasted spreadsheet values from becoming formulas unintentionally.
  return /^[=+\-@]/.test(text) ? `'${text}` : text;
}
