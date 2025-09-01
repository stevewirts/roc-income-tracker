function autoSizeAllColumns(sheet, extaPadding) {
  SpreadsheetApp.flush();         // Apply pending changes
  Utilities.sleep(500);           // Let the sheet render

  const lastColumn = sheet.getLastColumn();
  for (let col = 1; col <= lastColumn; col++) {
    sheet.autoResizeColumn(col);
    const currentWidth = sheet.getColumnWidth(col);
    sheet.setColumnWidth(col, currentWidth + extaPadding);  // Add 4 pixels of padding
  }
}

function freezeHeaders(sheet) {
  sheet.setFrozenRows(1);
}

function boldHeaders(sheet) {
  const lastColumn = sheet.getLastColumn();
  const headerRange = sheet.getRange(1, 1, 1, lastColumn);
  headerRange.setFontWeight("bold");
}

function filterHeaders(sheet) {
  const range = sheet.getDataRange();
  const filter = sheet.getFilter();
  if (filter) {
    filter.remove();
  }

  const lastColumn = sheet.getLastColumn();
  sheet.getRange(1, 1, 1, lastColumn).createFilter();
}

function addHeaderNotes(sheet, notesMap) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  headers.forEach((header, i) => {
    const note = notesMap[header];
    if (note) {
      sheet.getRange(1, i + 1).setNote(note);
    }
  });
}

function clearSheet(sheet) {
  if (!sheet) return;

  // Clear all content, formatting, notes, and filters
  sheet.clear(); // clears values + formats
  sheet.clearNotes();
  sheet.clearConditionalFormatRules();

  // Remove filter if present
  const filter = sheet.getFilter();
  if (filter) filter.remove();

  // Reset frozen rows/columns
  sheet.setFrozenRows(0);
  sheet.setFrozenColumns(0);

  // Reset column widths
  const lastCol = sheet.getMaxColumns();
  for (let i = 1; i <= lastCol; i++) {
    sheet.setColumnWidth(i, 100); // default width
  }

  // Optional: reset tab color
  sheet.setTabColor(null);
}

function insureClearedSheet(sheetName) {
  const sheet = getOrCreateSheet(sheetName);
  clearSheet(sheet);
  return sheet;
}

function getOrCreateSheet(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

//simple utility to extract sheet data as an array of objects keyed by column headers:
//ie:
//[
//  { Sym: "MSTW", TotInc: 123.45, ROCamt: 67.89 },
//  { Sym: "MSTY", TotInc: 234.56, ROCamt: 12.34 },
//  ...
//]
function getSheetData(sheet) {
  const range = sheet.getDataRange();
  const values = range.getValues();
  const headers = values[0];
  return values.slice(1).map(row => {
    const obj = {};
    headers.forEach((key, i) => {
      obj[key] = row[i];
    });
    return obj;
  });
}

/**
 * Returns the zero-based index of the first matching header in `headers`.
 * @param {string[]} headers   The sheet’s header row.
 * @param {string[]} options   Possible names, in priority order.
 * @returns {number}           Index of the first match, or -1 if none.
 */
function findHeaderIndex(headers, options) {
  for (let i = 0; i < options.length; i++) {
    const idx = headers.indexOf(options[i]);
    if (idx >= 0) return idx;
  }
  return -1;
}
/**
 * Looks up the current price for a given symbol.
 * Expects a sheet named "Prices" with headers in row 1:
 *   A: Sym   B: CurrPx
 */
function priceGet(sym) {
  const ss     = SpreadsheetApp.getActive();
  const pSheet = ss.getSheetByName('MarketData');
  if (!pSheet) throw new Error('Sheet "Prices" not found');
  
  // pull values once
  const rows   = pSheet.getRange(2, 1, pSheet.getLastRow() - 1, 2).getValues();
  const map    = rows.reduce((m, [s, p]) => {
    m[s] = +p;
    return m;
  }, {});
  
  return map[sym] || 0;
}

/**
 * Formats a Date object as yyyy-MM-dd.
 */
function formatDate_(dt) {
  if (!dt) return '';
  const tz = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  return Utilities.formatDate(new Date(dt), tz, 'yyyy-MM-dd');
}



