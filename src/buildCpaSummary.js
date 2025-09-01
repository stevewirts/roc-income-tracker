/**
 * Builds the CPA Summary sheet by reconciling BrokerROC, CpaNotes,
 * TrancheTracker and Transactions. Wraps CPA_Note and Summary
 * cells to avoid horizontal scrolling, and adds hover-notes to all headers.
 */
function buildCpaSummary() {
  const ss           = SpreadsheetApp.getActiveSpreadsheet();
  const brokerSheet  = ss.getSheetByName("BrokerROC");
  const cpaSheet     = ss.getSheetByName("CpaNotes");
  const synthSheet   = ss.getSheetByName("TrancheTracker");
  const txSheet      = ss.getSheetByName("Transactions");

  // Ensure CpaSummary exists and is cleared
  let summarySheet = insureClearedSheet("CpaSummary");
  if (!summarySheet) {
    summarySheet = ss.insertSheet("CpaSummary");
  }
  summarySheet.clear();

  // Build lookup maps
  const brokerBox3Map    = buildNoteMap(brokerSheet, "Sym",    "Box3");
  const brokerNoteMap    = buildNoteMap(brokerSheet, "Sym",    "Note");
  const cpaNoteMap       = buildCpaNoteMap(cpaSheet);
  const trancheTrackerIncome  = buildTrancheTrackerMap(synthSheet);
  const actualRocMap     = buildTransactionsRocMap(txSheet);

  // Union & sort all symbols
  const allSyms = Array.from(new Set([
    ...Object.keys(brokerBox3Map),
    ...Object.keys(brokerNoteMap),
    ...Object.keys(cpaNoteMap),
    ...Object.keys(trancheTrackerIncome),
    ...Object.keys(actualRocMap)
  ])).sort();

  // Build output rows
  const output = allSyms.map(sym => {
    const box3      = brokerBox3Map[sym]      || "";
    const synthTot  = trancheTrackerIncome[sym]?.Inc || 0;
    const synthRoc  = trancheTrackerIncome[sym]?.ROCamt || 0;
    const actualRoc = actualRocMap[sym]       || 0;
    const brokerNt  = brokerNoteMap[sym]      || "No broker note";
    const cpaNt     = cpaNoteMap[sym]         || "No CPA note";
    const flag      = (synthRoc > synthTot) ? "⚠️ ROC > Inc" : "";

    return [
      sym,
      box3,
      synthTot,
      synthRoc,
      actualRoc,
      flag,
      brokerNt,
      cpaNt,
      `${sym}: Box3=${box3}, SynthROC=${synthRoc}, ActualROC=${actualRoc}, ${brokerNt} | ${cpaNt}`
    ];
  });

  // Write headers + data
  const headers = [
    "Sym",             // Ticker symbol
    "Box3",            // Broker-reported ROC from BrokerROC tab
    "Inc",             // Total income from TrancheTracker
    "TrancheTrackerROC",// Total ROCAmt from TrancheTracker
    "TransactionsROC", // ROC computed from Transactions
    "BasisAdjFlag",    // Warning flag if ROC > income
    "BrokerNote",      // Notes from BrokerROC tab
    "CPA_Note",        // Notes from CpaNotes tab
    "Summary"          // Combined summary string
  ];
  summarySheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // Add hover-notes to headers
  const headerNotes = {
    Sym:             "Ticker symbol (Sym)",
    Box3:            "Box 3 non-dividend distributions from BrokerROC tab",
    Inc:             "Sum of Inc from TrancheTracker tab",
    TrancheTrackerROC:"Sum of ROCAmt from TrancheTracker tab",
    TransactionsROC:  "Sum of ROCAmt from Transactions tab",
    BasisAdjFlag:    "⚠️ if TrancheTracker ROC exceeds TrancheTracker income",
    BrokerNote:      "Note column from BrokerROC tab",
    CPA_Note:        "Aggregated CPA_Note values from CpaNotes tab",
    Summary:         "Concatenated summary of values and notes"
  };
  Object.entries(headerNotes).forEach(([h, note], i) => {
    summarySheet.getRange(1, i + 1).setNote(note);
  });

  // Write the data rows
  if (output.length) {
    summarySheet
      .getRange(2, 1, output.length, headers.length)
      .setValues(output);

    // Wrap text in the CPA_Note (col 8) and Summary (col 9) columns
    summarySheet
      .getRange(2, 8, output.length, 2)
      .setWrap(true);
  }

  // Apply filters & autosize
  filterHeaders(summarySheet);
  autoSizeAllColumns(summarySheet, headers.length);
  freezeHeaders(summarySheet);
}


/**
 * Generic map builder for single-value lookups.
 * If sumValues=true, sums numeric values per key.
 */
function buildNoteMap(sheet, keyColName, valueColName, sumValues = false) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const keyIdx  = headers.indexOf(keyColName);
  const valIdx  = headers.indexOf(valueColName);
  const data    = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();

  return data.reduce((map, row) => {
    const key = row[keyIdx];
    if (!key) return map;
    const raw = row[valIdx];
    if (sumValues) {
      map[key] = (map[key] || 0) + Number(raw || 0);
    } else {
      map[key] = raw;
    }
    return map;
  }, {});
}


/**
 * Aggregates multiple CPA notes per ticker into one string.
 * Expects headers: Sym, NoteDate, NoteType, CPA_Note.
 */
function buildCpaNoteMap(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const iSym     = headers.indexOf("Sym");
  const iDate    = headers.indexOf("NoteDate");
  const iType    = headers.indexOf("NoteType");
  const iNote    = headers.indexOf("CPA_Note");
  const rows     = sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues();

  const map = {};
  rows.forEach(r => {
    const sym = r[iSym];
    if (!sym) return;
    const entry = `${r[iDate]} (${r[iType]}): ${r[iNote]}`;
    (map[sym] = map[sym] || []).push(entry);
  });

  Object.keys(map).forEach(sym => {
    map[sym] = map[sym].join("; ");
  });
  return map;
}


/**
 * Reads { Sym, Inc, ROCamt } from TrancheTracker,
 * summing Inc and ROCamt per symbol.
 */
function buildTrancheTrackerMap(sheet) {
  const rows = getSheetData(sheet);
  return rows.reduce((map, r) => {
    const sym    = r["Sym"];
    const inc = Number(r["Inc"]) || 0;
    const rocAmt = Number(r["ROCAmt"]) || 0;
    if (!sym) return map;
    if (!map[sym]) map[sym] = { Inc: 0, ROCamt: 0 };
    map[sym].Inc += inc;
    map[sym].ROCamt += rocAmt;
    return map;
  }, {});
}


/**
 * Reads { Sym, RocAmount } from Transactions,
 * summing RocAmount per symbol.
 */
function buildTransactionsRocMap(sheet) {
  const rows = getSheetData(sheet);
  return rows.reduce((map, r) => {
    const sym = r["Sym"] || r["Symbol"];
    const v   = Number(r["ROCAmt"]) || 0;
    if (!sym) return map;
    map[sym] = (map[sym] || 0) + v;
    return map;
  }, {});
}


/**
 * Helper: returns an array of objects keyed by header.
 */
function getSheetData(sheet) {
  const vs = sheet.getDataRange().getValues();
  const hdr = vs.shift();
  return vs.map(r => {
    const o = {};
    hdr.forEach((h, i) => o[h] = r[i]);
    return o;
  });
}
