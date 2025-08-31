/**
 * buildIncomeTracker()
 *
 * Aggregates weekly and year-to-date distributions, taxable income, and
 * return of capital metrics per symbol, then writes to the IncomeTracker sheet
 * using the renamed headers and removed columns as requested.
 */
function buildIncomeTracker() {
  const ss           = SpreadsheetApp.getActiveSpreadsheet();
  const txSheet      = ss.getSheetByName("Transactions");
  const incSheetName = "IncomeTracker";
  const incSheet     = insureClearedSheet(incSheetName);
  const tz           = ss.getSpreadsheetTimeZone();
  const now          = new Date();

  // 1) Define output headers (renamed, with IncPS/TrCnt/UpdTS removed)
  const headers = [
    "Wk",         // Week start date (yyyy-MM-dd)
    "Sym",        // Ticker symbol
    "Dist",       // Distribution this week, by symbol
    "DistWkTot",  // Distribution this week, all symbols
    "DistYTD",    // YTD distribution, by symbol
    "DistYtdAll", // YTD distribution, all symbols
    "Inc",        // Taxable portion this week, by symbol
    "IncWkTot",   // Taxable portion this week, all symbols
    "IncYTD",     // YTD taxable portion, by symbol
    "IncYtdAll",  // YTD taxable portion, all symbols
    "Roc",        // ROC this week, by symbol
    "RocWkTot",   // ROC this week, all symbols
    "RocYTD",     // YTD ROC, by symbol
    "RocYtdAll",  // YTD ROC, all symbols
    "ShElig"      // Shares eligible
  ];

  // 2) Clear sheet & write headers
  incSheet.clearContents();
  incSheet.getRange(1, 1, 1, headers.length)
          .setValues([headers])
          .setFontWeight("bold")
          .setBackground("#d9e1f2");

  // 3) Header notes
  const notes = {
    Wk:         "Week start date (formatted yyyy-MM-dd)",
    Sym:        "Ticker symbol of the dividend event",
    Dist:       "Sum of dividend amounts for this symbol during the week",
    DistWkTot:  "Sum of dividend amounts across all symbols during the week",
    DistYTD:     "Cumulative dividends for this symbol year-to-date",
    DistYtdAll: "Cumulative dividends across all symbols year-to-date",
    Inc:        "Sum of taxable income for this symbol during the week",
    IncWkTot:   "Sum of taxable income across all symbols during the week",
    IncYTD:     "Cumulative taxable income for this symbol year-to-date",
    IncYtdAll:  "Cumulative taxable income across all symbols year-to-date",
    Roc:        "Sum of return of capital for this symbol during the week",
    RocWkTot:   "Sum of return of capital across all symbols during the week",
    RocYTD:     "Cumulative return of capital for this symbol year-to-date",
    RocYtdAll:  "Cumulative return of capital across all symbols year-to-date",
    ShElig:     "Total shares eligible at the moment of dividend"
  };
  headers.forEach((h, i) => {
    incSheet.getRange(1, i + 1).setNote(notes[h] || "");
  });

  // 4) Read Transactions data
  const raw  = txSheet.getDataRange().getValues();
  const hdrs = raw.shift().map(h => String(h).trim());
  const data = raw;

  // 5) Build a robust, lowercase header index
  const idx = hdrs.reduce((map, h, i) => {
    map[h.toLowerCase()] = i;
    return map;
  }, {});

  // 6) Reference the real headers from your sheet
  const iType          = idx["type"];      // "Type"
  const iWeekStart     = idx["wkstart"];   // "WkStart"
  const iSym           = idx["sym"];       // "Sym"
  const iDivTotal      = idx["dist"];      // "Dist"
  const iRocAmount     = idx["rocamt"];    // "ROCAmt"
  const iTaxableIncome = idx["inc"];       // "Inc"
  const iTotalShares   = idx["totshr"];    // "TotShr"

  // 7) Fail fast if any required column isn’t found
  [
    "type", "wkstart", "sym",
    "dist", "rocamt", "inc", "totshr"
  ].forEach(key => {
    if (idx[key] == null) {
      throw new Error(`Cannot find column “${key}” in Transactions headers.`);
    }
  });

  // 8) Aggregate per (Week, Symbol)
  const agg = {};
  data.forEach(row => {
    if (String(row[iType]).toLowerCase() !== "dividend") return;

    // normalize week
    let wkdt = row[iWeekStart];
    if (!(wkdt instanceof Date)) wkdt = new Date(wkdt);
    const wk = Utilities.formatDate(wkdt, tz, "yyyy-MM-dd");

    const sym   = String(row[iSym]).trim();
    const dist  = parseFloat(String(row[iDivTotal]).replace(/[^0-9.\-]/g, "")) || 0;
    const roc   = parseFloat(String(row[iRocAmount]).replace(/[^0-9.\-]/g, "")) || 0;
    const tax   = parseFloat(String(row[iTaxableIncome]).replace(/[^0-9.\-]/g, "")) || (dist - roc);
    const shEl  = parseFloat(row[iTotalShares]) || 0;

    const key = `${wk}|${sym}`;
    if (!agg[key]) {
      agg[key] = { Wk: wk, Sym: sym, Dist: 0, Inc: 0, Roc: 0, ShElig: 0 };
    }
    agg[key].Dist   += dist;
    agg[key].Inc    += tax;
    agg[key].Roc    += roc;
    agg[key].ShElig += shEl;
  });

  // 9) Weekly totals across all symbols
  const distWkAll = {};
  const incWkAll  = {};
  const rocWkAll  = {};
  Object.values(agg).forEach(e => {
    distWkAll[e.Wk] = (distWkAll[e.Wk] || 0) + e.Dist;
    incWkAll[e.Wk]  = (incWkAll[e.Wk]  || 0) + e.Inc;
    rocWkAll[e.Wk]  = (rocWkAll[e.Wk]  || 0) + e.Roc;
  });

  // 10) Year-to-date accumulators & build output
  const ytdDistAll  = {};
  const ytdIncAll   = {};
  const ytdRocAll   = {};
  const ytdDistSym  = {};
  const ytdIncSym   = {};
  const ytdRocSym   = {};
  const output      = [];

  Object.values(agg)
    .sort((a, b) => new Date(a.Wk) - new Date(b.Wk))
    .forEach(e => {
      const yr     = new Date(e.Wk).getFullYear();
      const symKey = `${yr}|${e.Sym}`;

      ytdDistAll[yr]      = (ytdDistAll[yr]   || 0) + e.Dist;
      ytdIncAll[yr]       = (ytdIncAll[yr]    || 0) + e.Inc;
      ytdRocAll[yr]       = (ytdRocAll[yr]    || 0) + e.Roc;

      ytdDistSym[symKey]  = (ytdDistSym[symKey] || 0) + e.Dist;
      ytdIncSym[symKey]   = (ytdIncSym[symKey]  || 0) + e.Inc;
      ytdRocSym[symKey]   = (ytdRocSym[symKey]  || 0) + e.Roc;

      output.push([
        e.Wk,
        e.Sym,
        e.Dist,
        distWkAll[e.Wk],
        ytdDistSym[symKey],
        ytdDistAll[yr],
        e.Inc,
        incWkAll[e.Wk],
        ytdIncSym[symKey],
        ytdIncAll[yr],
        e.Roc,
        rocWkAll[e.Wk],
        ytdRocSym[symKey],
        ytdRocAll[yr],
        e.ShElig
      ]);
    });

  // 11) Write output rows under headers
  if (output.length) {
    incSheet
      .getRange(2, 1, output.length, headers.length)
      .setValues(output);
  }

  // 12) Apply number/date formats
  const fmt = {
    Wk:         "yyyy-MM-dd",
    Dist:       "$#,##0.00", DistWkTot: "$#,##0.00",
    DistYTD:    "$#,##0.00", DistYtdAll: "$#,##0.00",
    Inc:        "$#,##0.00", IncWkTot:  "$#,##0.00",
    IncYTD:     "$#,##0.00", IncYtdAll:  "$#,##0.00",
    Roc:        "$#,##0.00", RocWkTot:  "$#,##0.00",
    RocYTD:     "$#,##0.00", RocYtdAll:  "$#,##0.00",
    ShElig:     "0.00"
  };
  headers.forEach((h, i) => {
    if (fmt[h] && output.length) {
      incSheet
        .getRange(2, i + 1, output.length)
        .setNumberFormat(fmt[h]);
    }
  });

// 13) Zebra-stripe rows by week (fixed to compare getTime() or string)
if (output.length) {
  // grab the raw week values
  const rawWks = incSheet
    .getRange(2, 1, output.length, 1)
    .getValues()
    .flat();

  // initialize with the first week’s timestamp or string
  let lastKey = rawWks[0] instanceof Date
    ? rawWks[0].getTime()
    : String(rawWks[0]);
  let toggle = false;

  // build a 2D array of backgrounds
  const bgs = rawWks.map(cell => {
    // normalize each cell to a comparable key
    const key = cell instanceof Date
      ? cell.getTime()
      : String(cell);

    // if we hit a new week, flip the toggle
    if (key !== lastKey) {
      toggle = !toggle;
      lastKey = key;
    }

    // fill this entire row with the chosen color
    return Array(headers.length)
      .fill(toggle ? "#cce6ff" : "#ffffff");
  });

  // apply the backgrounds in one go
  incSheet
    .getRange(2, 1, output.length, headers.length)
    .setBackgrounds(bgs);
}


  // 14) Final touches
  filterHeaders(incSheet);
  autoSizeAllColumns(incSheet, 4);
  freezeHeaders(incSheet);
}
