/**
 * Builds a synthetic dividends report from the updated Transactions sheet.
 * Accommodates the new RocPct header and blank spacer column.
 */
function buildSyntheticDividends() {
  const dividendNotesMap = {
    DivDt:    "Date of dividend distribution",
    TrID:     "Tranche ID receiving dividend",
    Sym:      "Symbol of underlying ETF",
    ShRem:    "Remaining shares eligible for dividend",
    IncPS:    "Income per share (taxable slice)",
    DivPS:    "Full dividend per share (IncPS/(1−ROCpct))",
    TotInc:   "Total dividend (ShRem × DivPS)",
    ROCpct:   "Return of capital percentage",
    ROCamt:   "Dollar amount of ROC",
    TaxInc:   "Taxable income (TotInc − ROCamt)",
    TrStatus: "Tranche status at time of dividend"
  };

  const ss       = SpreadsheetApp.getActiveSpreadsheet();
  const txSheet  = ss.getSheetByName("Transactions");
  const divSheet = insureClearedSheet("SyntheticDividends");

  // 1) Grab all data from Transactions
  const txData  = txSheet.getDataRange().getValues();
  const headers = txData[0];
  const rows    = txData.slice(1);

  // 2) Map headers → indices (zero-based)
  const idx = {
    type:           headers.indexOf("Type"),
    date:           headers.indexOf("Date"),
    sym:            headers.indexOf("Symbol"),
    trancheID:      headers.indexOf("TrancheID"),
    shares:         headers.indexOf("Shares"),
    incomePerShare: headers.indexOf("IncomePerShare"),
    rocPercent:     headers.indexOf("RocPct"),
    trancheStatus:  headers.indexOf("TrancheStatus")
  };

  // 3) Validate required columns
  Object.entries(idx).forEach(([key, i]) => {
    if (i < 0) throw new Error(`Missing required header: ${key}`);
  });

  // 4) Partition into buys/sells/dividends
  const dividendRows = rows.filter(r => String(r[idx.type]).toLowerCase() === "dividend");
  const buyRows      = rows.filter(r => String(r[idx.type]).toLowerCase() === "buy");
  const sellRows     = rows.filter(r => String(r[idx.type]).toLowerCase() === "sell");

  // 5) Prepare output
  const headerKeys = Object.keys(dividendNotesMap);
  const output     = [];

  // 6) Build each synthetic‐dividend line
  dividendRows.forEach(div => {
    const divDate        = new Date(div[idx.date]);
    const formattedDate  = Utilities.formatDate(divDate, ss.getSpreadsheetTimeZone(), "yyyy-MM-dd");
    const sym            = String(div[idx.sym]).trim();
    const incomePerShare = parseFloat(div[idx.incomePerShare]) || 0;
    const rocPct         = parseFloat(div[idx.rocPercent])   || 0;
    const divisor        = 1 - rocPct;
    const divPerShare    = divisor > 0 ? incomePerShare / divisor : 0;

    buyRows.forEach(buy => {
      const buyDate      = new Date(buy[idx.date]);
      const buySym       = String(buy[idx.sym]).trim();
      const trancheID    = buy[idx.trancheID];
      const sharesBought = parseFloat(buy[idx.shares]) || 0;

      if (buySym !== sym || buyDate > divDate) return;

      // Compute sold shares up to divDate
      const soldShares = sellRows
        .filter(s => s[idx.trancheID] === trancheID && new Date(s[idx.date]) <= divDate)
        .reduce((sum, s) => sum + (parseFloat(s[idx.shares]) || 0), 0);

      const remShares = Math.max(0, sharesBought - soldShares);
      if (remShares === 0) return;

      // Determine tranche status
      const trancheStatus = soldShares === 0
        ? "Open"
        : remShares === 0
          ? "Closed"
          : "Partial";

      // Calculate full dividend, ROC, and taxable income
      const totalDividend = remShares * divPerShare;
      const rocAmount     = totalDividend * rocPct;
      const taxableIncome = totalDividend - rocAmount;

      // Assemble row
      const rowData = {
        DivDt:    formattedDate,
        TrID:     trancheID,
        Sym:      sym,
        ShRem:    remShares,
        IncPS:    incomePerShare,
        DivPS:    divPerShare,
        TotInc:   totalDividend,
        ROCpct:   rocPct,
        ROCamt:   rocAmount,
        TaxInc:   taxableIncome,
        TrStatus: trancheStatus
      };

      output.push(headerKeys.map(k => rowData[k]));
    });
  });

  // 7) Sort by dividend date
  output.sort((a, b) => new Date(a[0]) - new Date(b[0]));

  // 8) Write headers + data
  divSheet.clearContents();
  divSheet.getRange(1, 1, 1, headerKeys.length).setValues([headerKeys]);
  if (output.length) {
    divSheet.getRange(2, 1, output.length, headerKeys.length).setValues(output);
  }

  // 9) Add notes, formatting, coloring, filters, autosize, freeze
  addHeaderNotes(divSheet, dividendNotesMap);
  formatDividendSheet(divSheet, headerKeys);
  applyDividendColoring(divSheet, output, headerKeys.length);
  filterHeaders(divSheet);
  autoSizeAllColumns(divSheet, 24);
  freezeHeaders(divSheet);
}

/**
 * Alternate row shading and highlight Partial status cells.
 */
function applyDividendColoring(sheet, output, totalCols) {
  if (output.length === 0) return;

  let colorToggle = false;
  let currentDate = output[0][0];
  let blockStart  = 0;

  for (let i = 1; i <= output.length; i++) {
    const rowDate  = output[i]?.[0];
    const rowIndex = i + 1;

    if (rowDate !== currentDate || i === output.length) {
      const bgColor = colorToggle ? "#cce6ff" : "#ffffff";

      for (let r = blockStart + 2; r <= i + 1; r++) {
        const status = sheet.getRange(r, totalCols).getValue();
        for (let c = 1; c <= totalCols; c++) {
          if (!(c === totalCols && status === "Partial")) {
            sheet.getRange(r, c).setBackground(bgColor);
          }
        }
      }

      currentDate = rowDate;
      colorToggle = !colorToggle;
      blockStart  = i;
    }
  }

  // Highlight all Partial tranches in red
  for (let i = 0; i < output.length; i++) {
    const status = output[i][totalCols - 1];
    if (status === "Partial") {
      sheet.getRange(i + 2, totalCols).setBackground("#ffcccc");
    }
  }
}

/**
 * Formats columns: dates, numbers, percentages, and currency.
 */
function formatDividendSheet(sheet, headerKeys) {
  const colMap = {};
  headerKeys.forEach((name, i) => colMap[name] = i + 1);
  const numRows = sheet.getLastRow();

  const currencyCols = ["IncPS", "DivPS", "TotInc", "ROCamt", "TaxInc"];
  const percentCols  = ["ROCpct"];
  const numberCols   = ["ShRem"];
  const dateCols     = ["DivDt"];

  currencyCols.forEach(col => {
    if (colMap[col]) {
      sheet.getRange(2, colMap[col], numRows - 1).setNumberFormat("$#,##0.00");
    }
  });

  percentCols.forEach(col => {
    if (colMap[col]) {
      sheet.getRange(2, colMap[col], numRows - 1).setNumberFormat("0.00%");
    }
  });

  numberCols.forEach(col => {
    if (colMap[col]) {
      sheet.getRange(2, colMap[col], numRows - 1).setNumberFormat("0.00");
    }
  });

  dateCols.forEach(col => {
    if (colMap[col]) {
      sheet.getRange(2, colMap[col], numRows - 1).setNumberFormat("yyyy-mm-dd");
    }
  });
}
