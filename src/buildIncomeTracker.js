function buildIncomeTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const divSheet = ss.getSheetByName("SyntheticDividends");
  const incSheetName = "IncomeTracker";
  const incSheet = insureClearedSheet(incSheetName);

  // 1) Write headers
  const headers = [
    "Wk",       // WeekStart
    "Sym",      // Symbol
    "WkInc",    // WeeklyIncome
    "YTDinc",   // Grand total income for the year
    "YTDsym",   // Symbol-specific YTD income
    "ROCamt",   // Return of capital
    "TaxInc",   // Taxable income
    "ShElig",   // Shares eligible
    "IncPS",    // Income per share
    "TrCnt",    // Tranches included
    "UpdTS"     // Timestamp
  ];
  incSheet.clearContents();
  incSheet.getRange(1, 1, 1, headers.length)
          .setValues([headers])
          .setFontWeight("bold")
          .setBackground("#d9e1f2");
  // hover notes
  const notesMap = {
    Wk:       "Week start date (typically Friday)",
    Sym:      "Ticker symbol of the income-generating asset",
    WkInc:    "Total income received for this symbol during the week",
    YTDinc:   "Running total income across all symbols for the calendar year",
    YTDsym:   "Running total income for this symbol during the calendar year",
    ROCamt:   "Return of capital portion (non-taxable)",
    TaxInc:   "Taxable income (WkInc minus ROCamt)",
    ShElig:   "Shares eligible for income during this period",
    IncPS:    "Income per share (WkInc ÷ ShElig)",
    TrCnt:    "Number of tranches contributing to this row",
    UpdTS:    "Timestamp of last update"
  };
  headers.forEach((h, i) => incSheet.getRange(1, i+1).setNote(notesMap[h] || ""));

  // 2) Pull SyntheticDividends data
  const divData = divSheet.getDataRange().getValues();
  const divHdrs  = divData.shift();
  const idx = divHdrs.reduce((m, h, i) => { m[h] = i; return m; }, {});
  const rows = divData;

  // 3) Aggregate by Week and Symbol
  const agg = {};
  rows.forEach(r => {
    const wk  = r[idx["DivDt"]];
    const sym = r[idx["Sym"]];
    const key = `${wk}|${sym}`;
    const tot = parseFloat(r[idx["TotInc"]]) || 0;
    const roc = parseFloat(r[idx["ROCamt"]]) || 0;
    const sh  = parseFloat(r[idx["ShRem"]]) || 0;
    if (!agg[key]) {
      agg[key] = { Wk: wk, Sym: sym, WkInc:0, ROCamt:0, ShElig:0, TrCnt:0 };
    }
    agg[key].WkInc   += tot;
    agg[key].ROCamt  += roc;
    agg[key].ShElig  += sh;
    agg[key].TrCnt   += 1;
  });

  // 4) Compute YTD totals and build output rows
  const ytdTotal = {};
  const ytdSym   = {};
  const now      = new Date();
  const output   = [];
  Object.values(agg)
        .sort((a,b) => new Date(a.Wk) - new Date(b.Wk))
        .forEach(entry => {
    const yearKey = new Date(entry.Wk).getFullYear();
    const symKey  = `${yearKey}|${entry.Sym}`;
    ytdTotal[yearKey] = (ytdTotal[yearKey]||0) + entry.WkInc;
    ytdSym[symKey]    = (ytdSym[symKey]||0) + entry.WkInc;

    output.push([
      entry.Wk,
      entry.Sym,
      entry.WkInc,
      ytdTotal[yearKey],
      ytdSym[symKey],
      entry.ROCamt,
      entry.WkInc - entry.ROCamt,
      entry.ShElig,
      entry.ShElig ? entry.WkInc/entry.ShElig : 0,
      entry.TrCnt,
      now
    ]);
  });

  // 5) Write aggregated rows
  if (output.length) {
    incSheet.getRange(2, 1, output.length, headers.length)
            .setValues(output);
  }

  // 6) Apply number formats
  const formatMap = {
    Wk:     "yyyy-MM-dd",
    WkInc:  "$#,##0.00",
    YTDinc: "$#,##0.00",
    YTDsym: "$#,##0.00",
    ROCamt: "$#,##0.00",
    TaxInc: "$#,##0.00",
    ShElig: "0.00",
    IncPS:  "$0.0000",
    TrCnt:  "0",
    UpdTS:  "yyyy-MM-dd HH:mm"
  };
  headers.forEach((h,i) => {
    const rng = incSheet.getRange(2, i+1, output.length);
    if (formatMap[h]) rng.setNumberFormat(formatMap[h]);
  });

  // —————————————
  // 7) ALTERNATE BACKGROUNDS BY WEEK
  // —————————————
  if (output.length) {
    const startRow = 2;
    const numRows   = output.length;
    const numCols   = headers.length;
    // 1) Grab the “Wk” column as Date objects
    const wkCol = incSheet
      .getRange(startRow, 1, numRows, 1)
      .getValues()
      .flat();

    // 2) Build a 2D array of background‐color strings
    const bgColors = [];
    let currentTime = wkCol[0].getTime();
    let toggle      = false;
    wkCol.forEach((cellDate, idx) => {
      const t = cellDate.getTime();
      if (t !== currentTime) {
        toggle      = !toggle;
        currentTime = t;
      }
      // choose two contrasting colors
      const color = toggle ? "#cce6ff" : "#ffffff";
      // one array per row, filled with the same color
      bgColors.push(new Array(numCols).fill(color));
    });

    // 3) Paint them all at once
    incSheet
      .getRange(startRow, 1, numRows, numCols)
      .setBackgrounds(bgColors);
  }

  // 8) Final touches
  filterHeaders(incSheet);
  autoSizeAllColumns(incSheet, 24);
  freezeHeaders(incSheet);
}
