/**
 * updateHistoricalTrancheMetrics()
 *
 * Reads the Transactions sheet, inserts a spacer column after TIDOverride,
 * then rebuilds the metrics columns to the right—removing the redundant Inc
 * field, renaming TaxInc → Inc, placing RocPS after IncPS, with columns
 * ordered as requested.
 */
function updateHistoricalTrancheMetrics() {
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Transactions");

  // 1) Read all data & locate fixed-left columns
  const raw     = sheet.getDataRange().getValues();
  const headers = raw[0];
  const dateIdx     = headers.indexOf("Date");
  const typeIdx     = headers.indexOf("Type");
  const symIdx      = headers.indexOf("Sym");
  const rocPctIdx   = headers.indexOf("RocPct");
  const sharesIdx   = headers.indexOf("Shr");
  const priceIdx    = headers.indexOf("Price");
  const distIdx     = headers.indexOf("Dist");
  const overrideIdx = headers.indexOf("TIDOverride");

  if (
    [dateIdx, typeIdx, symIdx, rocPctIdx, sharesIdx, priceIdx, distIdx, overrideIdx]
      .some(i => i < 0)
  ) {
    throw new Error(
      "Missing one of the required headers: " +
      "Date, Type, Sym, ROCPct, Shr, Price, Dist, TIDOverride"
    );
  }

  // 2) Collect only the valid-date rows
  const rows = [];
  for (let i = 1; i < raw.length; i++) {
    const cell = raw[i][dateIdx];
    if (cell instanceof Date && !isNaN(cell)) {
      rows.push(raw[i]);
    } else {
      break;
    }
  }
  const numRows = rows.length;

  // 3) Insert blank spacer column after TIDOverride
  const spacerCol1 = overrideIdx + 2;  // 1-based index
  sheet.insertColumnAfter(overrideIdx + 1);
  sheet.getRange(1, spacerCol1).setValue(" ");
  sheet
    .getRange(2, spacerCol1, numRows, 1)
    .setBackground("#D0E7E5");

  // 4) Define rebuilt headers (to the right of spacer),
  //    in this order: WkStart, TrID, CostBasis, DistPS, IncPS, RocPS, Inc, ROCAmt, TotShr, RemShr, TStat
  const rebuilt = [
    "WkStart",   // WeekStart
    "TrID",      // TrancheID
    "CostBasis",
    "DistPS",    // distribution per share
    "IncPS",     // taxable income per share
    "RocPS",     // return-of-capital per share
    "Inc",       // taxable portion of distribution (renamed from TaxInc)
    "ROCAmt",    // return-of-capital amount
    "TotShr",    // running total shares
    "RemShr",    // remaining shares in tranche
    "TStat"      // tranche status
  ];

  // 5) Clear everything to the right of spacer, then write rebuilt headers
  const lastCol       = sheet.getLastColumn();
  const firstRebuilt1 = spacerCol1 + 1;
  const colsToClear   = lastCol - firstRebuilt1 + 1;
  if (colsToClear > 0) {
    sheet
      .getRange(1, firstRebuilt1, numRows + 1, colsToClear)
      .clearContent()
      .clearFormat()
      .setNumberFormat("General");
  }
  sheet
    .getRange(1, firstRebuilt1, 1, rebuilt.length)
    .setValues([rebuilt]);

  // 6) Notes map for headers
  const notes = {
    Date:        "Trade settlement date.",
    Type:        "Event type: buy, sell, or dividend.",
    Sym:         "Ticker symbol.",
    ROCPct:      "Return-of-capital % (0–1).",
    Shr:         "Number of shares.",
    Price:       "Price per share.",
    Dist:        "Distribution = Inc + ROCAmt.",
    TIDOverride:"Manual tranche ID override.",
    WkStart:     "Monday of the event’s week.",
    TrID:        "Calculated tranche identifier (YYMMDD_A...).",
    CostBasis:   "Shares × Price for buy/sell events.",
    DistPS:      "Distribution per share = Dist ÷ TotalShares.",
    IncPS:       "Taxable income per share = Inc ÷ TotalShares.",
    RocPS:       "Return-of-capital per share = ROCAmt ÷ TotalShares.",
    Inc:         "Taxable portion of the distribution.",
    ROCAmt:      "Return-of-capital portion of the distribution.",
    TotShr:      "Running total shares held.",
    RemShr:      "Remaining shares in this tranche.",
    TStat:       "Open, Partial, or Closed tranche status."
  };

  // 7) Apply notes to headers
  const allHdrs = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0];
  allHdrs.forEach((h, j) => {
    if (notes[h]) {
      sheet.getRange(1, j + 1).setNote(notes[h]);
    }
  });

  // 8) Build name→index map for rebuilt fields
  const updatedHdrs = sheet
    .getRange(1, 1, 1, sheet.getLastColumn())
    .getValues()[0];
  const fCols = {};
  rebuilt.forEach(name => {
    const idx = updatedHdrs.indexOf(name);
    if (idx < 0) throw new Error("Missing rebuilt header: " + name);
    fCols[name] = idx;
  });

  // 9) Set number/date formats for rebuilt columns
  sheet
    .getRange(2, fCols["WkStart"] + 1, numRows, 1)
    .setNumberFormat("yyyy-MM-dd");
  ["CostBasis", "ROCAmt", "Inc"].forEach(c => {
    sheet
      .getRange(2, fCols[c] + 1, numRows, 1)
      .setNumberFormat("$#,##0.00");
  });
  ["DistPS", "IncPS", "RocPS"].forEach(c => {
    sheet
      .getRange(2, fCols[c] + 1, numRows, 1)
      .setNumberFormat("$#,##0.0000");
  });
  ["TotShr", "RemShr"].forEach(c => {
    sheet
      .getRange(2, fCols[c] + 1, numRows, 1)
      .setNumberFormat("0.0000");
  });

  // 10) Populate WeekStart, CostBasis, running shares,
  //     and distribution breakdown fields
  const symbolRun = {};
  rows.forEach((r, i) => {
    const rowNum = i + 2;
    const dt     = new Date(r[dateIdx]);
    const type   = String(r[typeIdx] || "").toLowerCase();
    const sym    = String(r[symIdx]  || "").trim().toUpperCase();
    const sh     = parseFloat(r[sharesIdx]) || 0;
    const px     =
      parseFloat(String(r[priceIdx] || "").replace(/[^0-9.\-]/g, "")) || 0;
    const pct    = parseFloat(r[rocPctIdx]) || 0;
    const rawDist= parseFloat(r[distIdx])   || 0;

    // WeekStart
    sheet
      .getRange(rowNum, fCols["WkStart"] + 1)
      .setValue(getWeekStart(dt));

    // CostBasis for buy
    if (type === "buy") {
      sheet
        .getRange(rowNum, fCols["CostBasis"] + 1)
        .setValue(sh * px);
    }

    // Running total shares
    symbolRun[sym] = (symbolRun[sym] || 0)
                   + (type === "buy"  ?  sh
                      : type === "sell" ? -sh
                                         : 0);
    sheet
      .getRange(rowNum, fCols["TotShr"] + 1)
      .setValue(symbolRun[sym]);

    // On dividend rows, fill DistPS, IncPS, RocPS, Inc, ROCAmt
    if (type === "dividend") {
      const ts     = symbolRun[sym] || 0;
      const rocAmt = pct * rawDist;
      const incAmt = (1 - pct) * rawDist;
      const incPS  = ts ? incAmt  / ts : 0;
      const rocPS  = ts ? rocAmt  / ts : 0;
      const distPS = ts ? rawDist / ts : 0;

      sheet.getRange(rowNum, fCols["DistPS"] + 1).setValue(distPS);
      sheet.getRange(rowNum, fCols["IncPS"]  + 1).setValue(incPS);
      sheet.getRange(rowNum, fCols["RocPS"]  + 1).setValue(rocPS);
      sheet.getRange(rowNum, fCols["Inc"]    + 1).setValue(incAmt);
      sheet.getRange(rowNum, fCols["ROCAmt"] + 1).setValue(rocAmt);
    }
  });

  // 11) Assign Tranche IDs (skip dividends)
  const trancheMap = {};
  const assigned   = [];
  rows.forEach((r, i) => {
    const rowNum = i + 2;
    const type   = String(r[typeIdx] || "").toLowerCase();
    if (type === "dividend") {
      assigned[i] = null;
      return;
    }

    let tid = r[overrideIdx];
    if (!tid) {
      const key = String(r[symIdx]).trim().toUpperCase()
                + "_" + Utilities.formatDate(
                              new Date(r[dateIdx]),
                              ss.getSpreadsheetTimeZone(),
                              "yyMMdd"
                          );
      const cnt = trancheMap[key] || 0;
      tid = key + "_" + String.fromCharCode(65 + cnt);
      trancheMap[key] = cnt + 1;
    }
    sheet.getRange(rowNum, fCols["TrID"] + 1).setValue(tid);
    assigned[i] = tid;
  });

  // 12) Compute RemShr & TStat, highlight “Partial”
  const buys = {};
  rows.forEach((r, i) => {
    const type = String(r[typeIdx] || "").toLowerCase();
    const id   = assigned[i];
    if (type === "buy" && id) {
      const s = parseFloat(r[sharesIdx]) || 0;
      buys[id] = (buys[id] || 0) + s;
    }
  });

  rows.forEach((r, i) => {
    const rowNum = i + 2;
    const type   = String(r[typeIdx] || "").toLowerCase();
    const id     = assigned[i];
    if (!id || type === "dividend") return;

    // sum sells through this row
    const sold = rows.reduce((sum, r2, j) => {
      const t2 = String(r2[typeIdx]).toLowerCase();
      if (
        t2 === "sell" &&
        assigned[j] === id &&
        new Date(r2[dateIdx]) <= new Date(r[dateIdx])
      ) {
        return sum + (parseFloat(r2[sharesIdx]) || 0);
      }
      return sum;
    }, 0);

    const total     = buys[id] || 0;
    const remaining = total - sold;
    const status    =
      remaining === 0    ? "Closed" :
      remaining < total  ? "Partial" :
                           "Open";

    sheet
      .getRange(rowNum, fCols["RemShr"] + 1)
      .setValue(remaining);

    const sc = sheet.getRange(rowNum, fCols["TStat"] + 1);
    sc.setValue(status);
    sc.setBackground(status === "Partial" ? "#FFCCCC" : null);
  });

  // 13) Final housekeeping
  autoSizeAllColumns(sheet, 4);
  filterHeaders(sheet);
  freezeHeaders(sheet);

  // lock spacer column width
  sheet.setColumnWidth(typeIdx + 1, 100);
  sheet.setColumnWidth(spacerCol1, 10);
}




/**
 * Returns the Monday of the week for a given Date.
 */
function getWeekStart(dateObj) {
  const dt = new Date(dateObj);
  const day = dt.getDay();               // Sun=0 … Sat=6
  const shift = (day + 6) % 7;           // Mon→0 … Sun→6
  dt.setDate(dt.getDate() - shift);
  return dt;
}
