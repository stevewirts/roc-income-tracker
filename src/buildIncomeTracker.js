function buildIncomeTracker() {
  const ss           = SpreadsheetApp.getActiveSpreadsheet();
  const txSheet      = ss.getSheetByName("Transactions");
  const incSheetName = "IncomeTracker";
  const incSheet     = insureClearedSheet(incSheetName);
  const tz           = ss.getSpreadsheetTimeZone();
  const now          = new Date();

  // 1) Define headers
  const headers = [
    "Wk",          // Week start date (yyyy-MM-dd)
    "Sym",         // Ticker symbol
    "Dist_WkSym",  // Distribution this week, by symbol
    "Dist_WkAll",  // Distribution this week, all symbols
    "Dist_YTDSym", // YTD distribution, by symbol
    "Dist_YTDAll", // YTD distribution, all symbols
    "Tax_WkSym",   // Taxable portion this week, by symbol
    "Tax_WkAll",   // Taxable portion this week, all symbols
    "Tax_YTDSym",  // YTD taxable portion, by symbol
    "Tax_YTDAll",  // YTD taxable portion, all symbols
    "ROC_WkSym",   // ROC this week, by symbol
    "ROC_WkAll",   // ROC this week, all symbols
    "ROC_YTDSym",  // YTD ROC, by symbol
    "ROC_YTDAll",  // YTD ROC, all symbols
    "ShElig",      // Shares eligible
    "IncPS",       // Distribution per share (Dist_WkSym ÷ ShElig)
    "TrCnt",       // Count of dividend events
    "UpdTS"        // Timestamp of last update
  ];

  // 2) Clear & write headers
  incSheet.clearContents();
  incSheet.getRange(1, 1, 1, headers.length)
          .setValues([headers])
          .setFontWeight("bold")
          .setBackground("#d9e1f2");

  // 3) Header notes
  const notes = {
    Wk:          "Week start date (formatted yyyy-MM-dd)",
    Sym:         "Ticker symbol of the dividend event",
    Dist_WkSym:  "Sum of Dividend amounts for this symbol during the week",
    Dist_WkAll:  "Sum of Dividend amounts across all symbols during the week",
    Dist_YTDSym: "Cumulative Dividends for this symbol year-to-date",
    Dist_YTDAll: "Cumulative Dividends across all symbols year-to-date",
    Tax_WkSym:   "Sum of TaxableIncome for this symbol during the week",
    Tax_WkAll:   "Sum of TaxableIncome across all symbols during the week",
    Tax_YTDSym:  "Cumulative TaxableIncome for this symbol year-to-date",
    Tax_YTDAll:  "Cumulative TaxableIncome across all symbols year-to-date",
    ROC_WkSym:   "Sum of RocAmount for this symbol during the week",
    ROC_WkAll:   "Sum of RocAmount across all symbols during the week",
    ROC_YTDSym:  "Cumulative RocAmount for this symbol year-to-date",
    ROC_YTDAll:  "Cumulative RocAmount across all symbols year-to-date",
    ShElig:      "TotalShares at the moment of dividend",
    IncPS:       "Distribution per share = Dist_WkSym ÷ ShElig",
    TrCnt:       "Number of dividend transactions aggregated",
    UpdTS:       "When this row was last generated"
  };
  headers.forEach((h, i) => {
    incSheet.getRange(1, i + 1).setNote(notes[h] || "");
  });

  // 4) Read Transactions data
  const raw = txSheet.getDataRange().getValues();
  const hdrs = raw.shift().map(h => h.trim());
  const data = raw;

  // 5) Column index map
  const idx = hdrs.reduce((m, h, i) => {
    m[h] = i;
    return m;
  }, {});

  const iType          = idx["Type"];
  const iWeekStart     = idx["WeekStart"];
  const iSym           = idx["Symbol"];
  const iDivTotal      = idx["Dividend"];
  const iRocAmount     = idx["RocAmount"];
  const iTaxableIncome = idx["TaxableIncome"];
  const iTotalShares   = idx["TotalShares"];

  // 6) Aggregate per (Week, Symbol)
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
      agg[key] = {
        Wk:         wk,
        Sym:        sym,
        Dist_WkSym: 0,
        Tax_WkSym:  0,
        ROC_WkSym:  0,
        ShElig:     0,
        TrCnt:      0
      };
    }
    agg[key].Dist_WkSym += dist;
    agg[key].Tax_WkSym  += tax;
    agg[key].ROC_WkSym  += roc;
    agg[key].ShElig     += shEl;
    agg[key].TrCnt      += 1;
  });

  // 7) Compute weekly totals across symbols
  const distWkAll = {};
  const taxWkAll  = {};
  const rocWkAll  = {};
  Object.values(agg).forEach(e => {
    distWkAll[e.Wk] = (distWkAll[e.Wk] || 0) + e.Dist_WkSym;
    taxWkAll[e.Wk]  = (taxWkAll[e.Wk]  || 0) + e.Tax_WkSym;
    rocWkAll[e.Wk]  = (rocWkAll[e.Wk]  || 0) + e.ROC_WkSym;
  });

  // 8) Year‐to‐date accumulators
  const ytdDistAll  = {};
  const ytdTaxAll   = {};
  const ytdRocAll   = {};
  const ytdDistSym  = {};
  const ytdTaxSym   = {};
  const ytdRocSym   = {};
  const output      = [];

  Object.values(agg)
    .sort((a, b) => new Date(a.Wk) - new Date(b.Wk))
    .forEach(e => {
      const yr       = new Date(e.Wk).getFullYear();
      const symKey   = `${yr}|${e.Sym}`;

      ytdDistAll[yr]   = (ytdDistAll[yr]   || 0) + e.Dist_WkSym;
      ytdTaxAll[yr]    = (ytdTaxAll[yr]    || 0) + e.Tax_WkSym;
      ytdRocAll[yr]    = (ytdRocAll[yr]    || 0) + e.ROC_WkSym;

      ytdDistSym[symKey] = (ytdDistSym[symKey] || 0) + e.Dist_WkSym;
      ytdTaxSym[symKey]  = (ytdTaxSym[symKey]  || 0) + e.Tax_WkSym;
      ytdRocSym[symKey]  = (ytdRocSym[symKey]  || 0) + e.ROC_WkSym;

      output.push([
        e.Wk,
        e.Sym,
        e.Dist_WkSym,
        distWkAll[e.Wk],
        ytdDistSym[symKey],
        ytdDistAll[yr],
        e.Tax_WkSym,
        taxWkAll[e.Wk],
        ytdTaxSym[symKey],
        ytdTaxAll[yr],
        e.ROC_WkSym,
        rocWkAll[e.Wk],
        ytdRocSym[symKey],
        ytdRocAll[yr],
        e.ShElig,
        e.ShElig ? e.Dist_WkSym / e.ShElig : 0,
        e.TrCnt,
        now
      ]);
    });

  // 9) Write output
  if (output.length) {
    incSheet
      .getRange(2, 1, output.length, headers.length)
      .setValues(output);
  }

  // 10) Apply number formats
  const fmt = {
    Wk:           "yyyy-MM-dd",
    Dist_WkSym:   "$#,##0.00", Dist_WkAll: "$#,##0.00",
    Dist_YTDSym:  "$#,##0.00", Dist_YTDAll: "$#,##0.00",
    Tax_WkSym:    "$#,##0.00", Tax_WkAll:  "$#,##0.00",
    Tax_YTDSym:   "$#,##0.00", Tax_YTDAll:  "$#,##0.00",
    ROC_WkSym:    "$#,##0.00", ROC_WkAll:  "$#,##0.00",
    ROC_YTDSym:   "$#,##0.00", ROC_YTDAll:  "$#,##0.00",
    ShElig:       "0.00",
    IncPS:        "$0.0000",
    TrCnt:        "0",
    UpdTS:        "yyyy-MM-dd HH:mm"
  };
  headers.forEach((h, i) => {
    if (fmt[h]) {
      incSheet
        .getRange(2, i + 1, output.length)
        .setNumberFormat(fmt[h]);
    }
  });

  // 11) Zebra-strip rows by week
  if (output.length) {
    const wkCol = incSheet.getRange(2, 1, output.length, 1).getValues().flat();
    let last = wkCol[0].getTime(), toggle = false;
    const bgs = wkCol.map(d => {
      if (d.getTime() !== last) {
        toggle = !toggle;
        last = d.getTime();
      }
      return new Array(headers.length).fill(toggle ? "#cce6ff" : "#ffffff");
    });
    incSheet.getRange(2, 1, output.length, headers.length).setBackgrounds(bgs);
  }

  // 12) Final touches
  filterHeaders(incSheet);
  autoSizeAllColumns(incSheet, 4);
  freezeHeaders(incSheet);
}



