/**
 * Reads your Transactions sheet,
 * injects a blank spacer column after TrancheIDOverride,
 * computes WeekStart for each record,
 * writes everything back, auto‐resizes,
 * then locks the spacer column width to 10px.
 */
function updateHistoricalTrancheMetrics() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("Transactions");

  // 1) Read all data + find existing column indexes
  const raw = sheet.getDataRange().getValues();
  const headers     = raw[0];
  const dateIdx     = headers.indexOf("Date");
  const typeIdx     = headers.indexOf("Type");
  const symIdx      = headers.indexOf("Symbol");
  const rocPctIdx   = headers.indexOf("RocPct");
  const sharesIdx   = headers.indexOf("Shares");
  const priceIdx    = headers.indexOf("Price");
  const divIdx      = headers.indexOf("Dividend");
  const overrideIdx = headers.indexOf("TrancheIDOverride");

  if ([dateIdx, typeIdx, symIdx, rocPctIdx, sharesIdx, priceIdx, divIdx, overrideIdx]
      .some(i => i < 0)) {
    throw new Error(
      "Missing one of the required headers: Date, Type, Symbol, RocPct, Shares, Price, Dividend, TrancheIDOverride"
    );
  }

  // 2) Extract only rows with a valid Date
  const rows = [];
  for (let i = 1; i < raw.length; i++) {
    if (raw[i][dateIdx] instanceof Date && !isNaN(raw[i][dateIdx])) {
      rows.push(raw[i]);
    } else {
      break;
    }
  }
  const numRows = rows.length;

  // 3) Insert spacer column immediately after TrancheIDOverride
  const spacerCol1 = overrideIdx + 2;      // 1-based index of new spacer
  sheet.insertColumnAfter(overrideIdx + 1);
  sheet.getRange(1, spacerCol1).setValue(" ");
  sheet.getRange(2, spacerCol1, numRows, 1)
       .setBackground("#D0E7E5");

  // 4) Define which fields we’ll rebuild (no CurrentPrice)
  const rebuilt = [
    "TrancheID",
    "WeekStart",
    "CostBasis",
    "RocAmount",
    "TaxableIncome",
    "Income",
    "IncomePerShare",
    "TotalShares",
    "RemShares",
    "TrancheStatus"
  ];

  // 5) Clear everything to the right of the spacer
  const lastCol       = sheet.getLastColumn();
  const firstRebuilt1 = spacerCol1 + 1;
  const colsToClear   = lastCol - firstRebuilt1 + 1;
  if (colsToClear > 0) {
    sheet.getRange(1, firstRebuilt1, numRows + 1, colsToClear)
         .clearContent()
         .clearFormat()
         .setNumberFormat("General");
  }

  // 6) Rewrite the rebuilt headers row
  sheet.getRange(1, firstRebuilt1, 1, rebuilt.length).setValues([rebuilt]);

  // 7) Full hover-notes map for every header, old + new
  const notes = {
    "Date":              "Trade settlement date.",
    "Type":              "Event type: buy, sell, or dividend.",
    "Symbol":            "Ticker symbol of the security.",
    "RocPct":            "Return‐of‐capital percentage (0–1).",
    "Shares":            "Number of shares for this event.",
    "Price":             "Trade price per share.",
    "Dividend":          "Gross dividend per share.",
    "TrancheIDOverride": "Manual override for the tranche ID.",
    "TrancheID":         "Calculated tranche identifier (YYMMDD suffix).",
    "WeekStart":         "Monday date of the event’s week.",
    "CostBasis":         "Shares × Price for buy events.",
    "RocAmount":         "Return‐of‐capital portion of dividend.",
    "TaxableIncome":     "Taxable portion of dividend income.",
    "Income":            "Total dividend income (ROC + taxable).",
    "IncomePerShare":    "Dividend income per share (taxable only).",
    "TotalShares":       "Running cumulative shares held.",
    "RemShares":         "Remaining shares in this tranche.",
    "TrancheStatus":     "Open, Partial, or Closed status."
  };

  // 8) Apply notes to every header cell
  const allHdrs = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  allHdrs.forEach((h, j) => {
    if (notes[h]) {
      sheet.getRange(1, j + 1).setNote(notes[h]);
    }
  });

  // 9) Build a name→column map for our rebuilt fields
  const updatedHdrs = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const fCols = {};
  rebuilt.forEach(name => {
    const idx = updatedHdrs.indexOf(name);
    if (idx < 0) throw new Error("Missing rebuilt header: " + name);
    fCols[name] = idx;
  });

  // 10) Set number/date formats for rebuilt columns
  sheet.getRange(2, fCols["WeekStart"] + 1, numRows, 1)
       .setNumberFormat("yyyy-MM-dd");
  ["CostBasis", "RocAmount", "TaxableIncome", "Income"].forEach(c => {
    sheet.getRange(2, fCols[c] + 1, numRows, 1)
         .setNumberFormat("$#,##0.00");
  });
  sheet.getRange(2, fCols["IncomePerShare"] + 1, numRows, 1)
       .setNumberFormat("$#,##0.0000");
  ["TotalShares", "RemShares"].forEach(c => {
    sheet.getRange(2, fCols[c] + 1, numRows, 1)
         .setNumberFormat("0.0000");
  });

  // 11) Populate WeekStart, CostBasis, running TotalShares, and dividend metrics
  const symbolRun = {};
  rows.forEach((r, i) => {
    const rowNum = i + 2;
    const dt     = new Date(r[dateIdx]);
    const type   = String(r[typeIdx] || "").toLowerCase();
    const sym    = String(r[symIdx]  || "").trim().toUpperCase();
    const sh     = parseFloat(r[sharesIdx]) || 0;
    const px     = parseFloat(
                     String(r[priceIdx] || "").replace(/[^0-9.\-]/g, "")
                   ) || 0;

    // WeekStart
    sheet.getRange(rowNum, fCols["WeekStart"] + 1)
         .setValue(getWeekStart(dt));

    // CostBasis for buys
    if (type === "buy" && sh) {
      sheet.getRange(rowNum, fCols["CostBasis"] + 1)
           .setValue(sh * px);
    }

    // Running total shares
    symbolRun[sym] = (symbolRun[sym] || 0)
                   + (type === "buy" ? sh
                      : type === "sell" ? -sh
                      : 0);
    sheet.getRange(rowNum, fCols["TotalShares"] + 1)
         .setValue(symbolRun[sym]);

    // Dividend breakdown
    if (type === "dividend") {
      const pct    = parseFloat(r[rocPctIdx]) || 0;
      const dv     = parseFloat(r[divIdx])    || 0;
      const ts     = symbolRun[sym]           || 0;
      const rocAmt = pct * dv * ts;
      const taxInc = (1 - pct) * dv * ts;
      const ips    = (1 - pct) * dv;
      const inc    = rocAmt + taxInc || dv * ts;

      if (rocAmt) sheet.getRange(rowNum, fCols["RocAmount"] + 1).setValue(rocAmt);
      if (taxInc) sheet.getRange(rowNum, fCols["TaxableIncome"] + 1).setValue(taxInc);
      if (ips)    sheet.getRange(rowNum, fCols["IncomePerShare"] + 1).setValue(ips);
      if (inc)    sheet.getRange(rowNum, fCols["Income"] + 1).setValue(inc);
    }
  });

  // 12) Assign TrancheIDs (override or auto-generate), but skip dividends
  const trancheMap = {};
  const assigned   = [];
  rows.forEach((r, i) => {
    const rowNum = i + 2;
    const type   = String(r[typeIdx] || "").toLowerCase();

    // dividend rows get no TrancheID
    if (type === "dividend") {
      assigned[i] = null;
      // ensure the cell stays blank
      return;
    }

    // override or auto-generate ID
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

    sheet.getRange(rowNum, fCols["TrancheID"] + 1).setValue(tid);
    assigned[i] = tid;
  });

  // 13) Compute remaining shares, status & color “Partial” red
  const buys = {};
  rows.forEach((r, i) => {
    if (String(r[typeIdx]).toLowerCase() === "buy") {
      const s  = parseFloat(r[sharesIdx]) || 0;
      const id = assigned[i];
      if (id) buys[id] = (buys[id] || 0) + s;
    }
  });

  rows.forEach((r, i) => {
    const rowNum = i + 2;
    const type   = String(r[typeIdx] || "").toLowerCase();
    const id     = assigned[i];
    if (!id || type === "dividend") return;

    const sold = rows.reduce((sum, r2, j) => {
      if (
        String(r2[typeIdx]).toLowerCase() === "sell" &&
        assigned[j] === id &&
        new Date(r2[dateIdx]) <= new Date(r[dateIdx])
      ) {
        return sum + (parseFloat(r2[sharesIdx]) || 0);
      }
      return sum;
    }, 0);

    const total     = buys[id] || 0;
    const remaining = total - sold;
    const status    = remaining === 0 ? "Closed"
                    : remaining < total    ? "Partial"
                                           : "Open";

    sheet.getRange(rowNum, fCols["RemShares"] + 1).setValue(remaining);

    const sc = sheet.getRange(rowNum, fCols["TrancheStatus"] + 1);
    sc.setValue(status);
    sc.setBackground(status === "Partial" ? "#FFCCCC" : null);
  });

  // 14) Final housekeeping
  autoSizeAllColumns(sheet, 14);
  filterHeaders(sheet);
  freezeHeaders(sheet);

  // lock spacer column to exactly 10px
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
