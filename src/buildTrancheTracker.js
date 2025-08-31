/**
 * buildTrancheTracker()
 *
 * For each “Dividend” in Transactions, emit one row per open tranche
 * matching that symbol, preserving the by-week → by-tranche logic.
 * Applies group-based zebra striping and highlights any “Partial” status cells in red.
 *
 * Output columns (in order):
 *   WkStart, DistDt, TrID, CostBasis,
 *   DistPS, IncPS, RocPS, Inc, ROCAmt,
 *   TotShr, RemShr, TStat
 *
 * Assumes helper functions:
 *   findHeaderIndex(), insureClearedSheet(),
 *   filterHeaders(), autoSizeAllColumns(), freezeHeaders()
 */
function buildTrancheTracker() {
  const ss        = SpreadsheetApp.getActiveSpreadsheet();
  const txSheet   = ss.getSheetByName("Transactions");
  const outSheet  = insureClearedSheet("TrancheTracker");
  const tz        = ss.getSpreadsheetTimeZone();

  // 1) Read data and headers
  const data      = txSheet.getDataRange().getValues();
  const headersIn = data.shift();
  const rows      = data;

  // 2) Map headers to indices
  const idx = {
    type     : findHeaderIndex(headersIn, ["Type"]),
    sym      : findHeaderIndex(headersIn, ["Sym","Symbol"]),
    date     : findHeaderIndex(headersIn, ["Date","DistDt"]),
    wkStart  : findHeaderIndex(headersIn, ["WkStart","WeekStart","WeekStarting"]),
    trID     : findHeaderIndex(headersIn, ["TrID","TrancheID"]),
    costBase : findHeaderIndex(headersIn, ["CostBasis","Cost Basis"]),
    totShr   : findHeaderIndex(headersIn, ["TotShr","TotalShares"]),
    remShr   : findHeaderIndex(headersIn, ["RemShr","ShRem"]),
    distPS   : findHeaderIndex(headersIn, ["DistPS"]),
    incPS    : findHeaderIndex(headersIn, ["IncPS"]),
    rocPS    : findHeaderIndex(headersIn, ["RocPS"]),
    incAmt   : findHeaderIndex(headersIn, ["Inc","TaxInc"]),
    rocAmt   : findHeaderIndex(headersIn, ["ROCAmt"]),
    tStat    : findHeaderIndex(headersIn, ["TStat","TrStatus","TrStat"])
  };

  // 3) Validate headers
  Object.entries(idx).forEach(([key, col]) => {
    if (col < 0) {
      throw new Error(`Missing required header "${key}". Found: ${headersIn.join(", ")}`);
    }
  });

  // 4) Split buys vs. dividends
  const buys = rows
    .filter(r => String(r[idx.type]).toLowerCase() === "buy")
    .map(r => ({
      sym      : r[idx.sym],
      trID     : r[idx.trID],
      costBase : r[idx.costBase],
      totShr   : parseFloat(r[idx.totShr]) || 0,
      remShr   : parseFloat(r[idx.remShr]) || 0,
      tStat    : r[idx.tStat]
    }));

  const divs = rows.filter(r => String(r[idx.type]).toLowerCase() === "dividend");

  // 5) Cross‐join dividends with open tranches
  const output = [];
  divs.forEach(rDiv => {
    const sym     = rDiv[idx.sym];
    const wkStart = Utilities.formatDate(new Date(rDiv[idx.wkStart]), tz, "yyyy-MM-dd");
    const distDt   = Utilities.formatDate(new Date(rDiv[idx.date]),    tz, "yyyy-MM-dd");
    const distPS  = parseFloat(rDiv[idx.distPS]) || 0;
    const incPS   = parseFloat(rDiv[idx.incPS])  || 0;
    const rocPS   = parseFloat(rDiv[idx.rocPS])  || 0;

    buys.forEach(b => {
      if (b.sym === sym && b.remShr > 0 && String(b.tStat).toLowerCase() === "open") {
        const incAmt = incPS * b.remShr;
        const rocAmt = rocPS * b.remShr;
        output.push([
          wkStart,
          distDt,
          b.trID,
          b.costBase || 0,
          distPS,
          incPS,
          rocPS,
          incAmt,
          rocAmt,
          b.totShr,
          b.remShr,
          b.tStat || ""
        ]);
      }
    });
  });

  // 6) Sort by WkStart → DistDt → TrID
  output.sort((a, b) => {
    const wk = new Date(a[0]) - new Date(b[0]);
    if (wk) return wk;
    const dt = new Date(a[1]) - new Date(b[1]);
    if (dt) return dt;
    return String(a[2]).localeCompare(b[2]);
  });

  // 7) Write headers and data
  const headersOut = [
    "WkStart", "DistDt", "TrID", "CostBasis",
    "DistPS", "IncPS", "RocPS", "Inc", "ROCAmt",
    "TotShr", "RemShr", "TStat"
  ];
  outSheet.clearContents();
  outSheet.getRange(1, 1, 1, headersOut.length).setValues([headersOut]);
  if (output.length) {
    outSheet.getRange(2, 1, output.length, headersOut.length)
            .setValues(output);
  }

  // 8) Apply number/date formatting
  const formats = {
    WkStart   : "yyyy-MM-dd",
    DistDt     : "yyyy-MM-dd",
    CostBasis : "$#,##0.00",
    DistPS    : "$#,##0.0000",
    IncPS     : "$#,##0.0000",
    RocPS     : "$#,##0.0000",
    Inc       : "$#,##0.00",
    ROCAmt    : "$#,##0.00",
    TotShr    : "0.0000",
    RemShr    : "0.0000"
  };
  headersOut.forEach((h, i) => {
    if (formats[h] && output.length) {
      outSheet.getRange(2, i + 1, output.length).setNumberFormat(formats[h]);
    }
  });

  // 9) Zebra‐striping and highlight “Partial”
  applyDividendColoring(outSheet, output, headersOut.length);

  // 10) Final housekeeping
  filterHeaders(outSheet);
  autoSizeAllColumns(outSheet, 24);
  freezeHeaders(outSheet);
}


/**
 * Alternate row shading by WkStart groups and highlight Partial status cells.
 */
function applyDividendColoring(sheet, output, totalCols) {
  if (output.length === 0) return;

  let colorToggle = false;
  let currentDate = output[0][0];
  let blockStart  = 0;

  for (let i = 1; i <= output.length; i++) {
    const rowDate = output[i]?.[0];

    if (rowDate !== currentDate || i === output.length) {
      const bgColor = colorToggle ? "#cce6ff" : "#ffffff";

      for (let r = blockStart + 2; r <= i + 1; r++) {
        const status = sheet.getRange(r, totalCols).getValue();
        for (let c = 1; c <= totalCols; c++) {
          const cell = sheet.getRange(r, c);
          if (c === totalCols && status === "Partial") {
            cell.setBackground("#ffcccc");
          } else {
            cell.setBackground(bgColor);
          }
        }
      }

      currentDate  = rowDate;
      colorToggle  = !colorToggle;
      blockStart   = i;
    }
  }
}
