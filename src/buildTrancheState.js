function buildTrancheState() {
  const notesMap = {
    ID:            "Unique tranche identifier",
    Sym:           "Ticker symbol",
    BuyDt:         "Date shares were purchased",
    ShBuy:         "Number of shares bought",
    BuyPx:         "Purchase price per share",
    ShSold:        "Number of shares sold",
    SellPx:        "Latest sale price",
    ShRem:         "Remaining shares",
    CurrPx:        "Current market price",
    CostBasis:     "Original cost basis (PurchasePrice × SharesBought)",
    AdjBasis:      "Adjusted cost basis (after ROC)",
    ConBasis:      "Consumed basis",
    ConBasisPct:   "Consumed basis as percentage of cost basis",
    MktVal:        "Market value of remaining shares",
    UnrlGain:      "Unrealized gain",
    ExitGain:      "Projected gain if sold now",
    ExitFlg:       "Exit readiness flag",
    RlzGain:       "Realized gain from sold shares",
    Status:        "Tranche status (Open, Partial, Closed)",
    HeldDays:      "Days held since purchase",
    LTGDays:       "Days until long-term gain threshold"
  };

  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const txSheet     = ss.getSheetByName("Transactions");
  const marketSheet = ss.getSheetByName("MarketData");
  const stateSheet  = insureClearedSheet("TrancheState");
  clearSheet(stateSheet, 8);

  // 1) build price lookup
  const mktData = marketSheet.getDataRange().getValues();
  const mktHdr  = mktData[0].map(String);
  const tkrI    = mktHdr.indexOf("Ticker");
  const pxI     = mktHdr.indexOf("CurrentPrice");
  if (tkrI < 0 || pxI < 0) {
    throw new Error("MarketData must have 'Ticker' and 'CurrentPrice'");
  }
  const marketMap = {};
  mktData.slice(1).forEach(r => {
    const sym = r[tkrI], px = parseFloat(r[pxI]) || 0;
    if (sym) marketMap[sym] = px;
  });

  // 2) read & index transactions
  const txData = txSheet.getDataRange().getValues();
  const hdr    = txData[0].map(String);
  const rows   = txData.slice(1);

  const idx = {
    type:      hdr.indexOf("Type"),
    trancheID: hdr.indexOf("TrID"),
    sym:       hdr.indexOf("Sym"),
    date:      hdr.indexOf("Date"),
    shares:    hdr.indexOf("Shr"),
    price:     hdr.indexOf("Price"),
    rocAmt:    hdr.indexOf("ROCAmt"),
    rocPct:    hdr.indexOf("RocPct"),
    dividend:  hdr.indexOf("Dist")
  };
  Object.entries(idx).forEach(([k,v]) => {
    if (v < 0) throw new Error("Missing header '" + k + "'");
  });

  // 3) init buy tranches
  const trancheMap = {};
  const latestSale = {};
  const rocMap     = {};

  rows.forEach(r => {
    if (String(r[idx.type]).toLowerCase() !== "buy") return;
    const id  = r[idx.trancheID];
    const sym = r[idx.sym];
    trancheMap[id] = {
      TrancheID:     id,
      Symbol:        sym,
      BuyDate:       r[idx.date],
      SharesBought:  parseFloat(r[idx.shares])  || 0,
      PurchasePrice: parseFloat(r[idx.price])   || 0,
      CurrentPrice:  marketMap[sym]             || 0,
      SharesSold:    0
    };
  });

  // 4) process sells + dividends→ROC
  rows.forEach(r => {
    const txType = String(r[idx.type]).toLowerCase();
    const id     = r[idx.trancheID];
    const t      = trancheMap[id];
    if (!t && txType !== "dividend") return;

    // 4a) sell
    if (t && txType === "sell") {
      const sold = parseFloat(r[idx.shares]) || 0;
      const px   = parseFloat(r[idx.price])  || 0;
      t.SharesSold += sold;
      const dt = new Date(r[idx.date]);
      if (!latestSale[id] || dt > latestSale[id].date) {
        latestSale[id] = { price: px, date: dt };
      }
    }

    // 4b) dividend ⇒ pro-rata ROC
    if (txType === "dividend") {
      const sym = r[idx.sym];
      // collect all open tranches for this ticker
      const openIDs = Object.keys(trancheMap)
                            .filter(key => trancheMap[key].Symbol === sym);
      // compute total remaining shares
      const totalRem = openIDs
        .reduce((sum, key) => {
          const tr = trancheMap[key];
          return sum + (tr.SharesBought - tr.SharesSold);
        }, 0);
      if (totalRem <= 0) return;

      // clean-parse per-share dividend
      const rawDiv = r[idx.dividend];
      const divPS  = parseFloat(String(rawDiv).replace(/[^0-9.\-]/g, "")) || 0;

      // clean-parse explicit RocAmount
      const rawAmt = r[idx.rocAmt];
      const amtVal = parseFloat(String(rawAmt).replace(/[^0-9.\-]/g, ""));

      // clean-parse RocPct
      const pctVal = parseFloat(String(r[idx.rocPct])
                     .replace(/[^0-9.\-]/g, "")) / 100 || 0;

      openIDs.forEach(key => {
        const tr   = trancheMap[key];
        const rem  = tr.SharesBought - tr.SharesSold;
        if (rem <= 0) return;

        // use explicit amount if you set one,
        // otherwise pctVal * (divPS * rem)
        const rocVal = !isNaN(amtVal) && amtVal > 0
                     ? amtVal * (rem / totalRem)    // split amt pro-rata
                     : pctVal * divPS * rem;

        rocMap[key] = (rocMap[key] || 0) + rocVal;
      });
    }
  });

  // 5) build output
  const keys   = Object.keys(notesMap);
  const output = [ keys ];
  const today  = new Date();

  Object.values(trancheMap).forEach(t => {
    const rem  = t.SharesBought - t.SharesSold;
    const cost = t.PurchasePrice * t.SharesBought;
    const roc  = rocMap[t.TrancheID] || 0;
    const adj  = cost - roc;
    const con  = roc;
    const pct  = cost ? con / cost : 0;

    const sale = latestSale[t.TrancheID] || {};
    const sellPx = sale.price || "";

    const mktVal   = rem * t.CurrentPrice;
    const unrlGain = mktVal - adj;
    const exitGain = unrlGain;
    let exitFlg = "";
    if (rem > 0) {
      exitFlg = mktVal >= adj           ? "Yes"
              : mktVal >= 0.98 * adj     ? "Partial"
              : "";
    }
    const rlzGain = (t.SharesSold && sale.price)
                  ? (sale.price - t.PurchasePrice) * t.SharesSold
                  : "";

    let heldDays = "", ltgDays = "";
    if (t.BuyDate) {
      const diff = today - new Date(t.BuyDate);
      heldDays = Math.floor(diff / (1000*60*60*24));
      ltgDays  = Math.max(0, 365 - heldDays);
    }

    const rowObj = {
      ID:           t.TrancheID,
      Sym:          t.Symbol,
      BuyDt:        t.BuyDate,
      ShBuy:        t.SharesBought,
      BuyPx:        t.PurchasePrice,
      ShSold:       t.SharesSold || "",
      SellPx:       sellPx,
      ShRem:        rem || "",
      CurrPx:       t.CurrentPrice,
      CostBasis:    cost,
      AdjBasis:     adj,
      ConBasis:     con,
      ConBasisPct:  pct,
      MktVal:       mktVal,
      UnrlGain:     unrlGain,
      ExitGain:     exitGain,
      ExitFlg:      exitFlg,
      RlzGain:      rlzGain,
      Status:       rem === 0 ? "Closed" : (rem < t.SharesBought ? "Partial" : "Open"),
      HeldDays:     heldDays,
      LTGDays:      ltgDays
    };

    output.push(keys.map(k => rowObj[k]));
  });

  // 6) write & format
  stateSheet.clearContents();
  stateSheet
    .getRange(1,1,output.length, output[0].length)
    .setValues(output);

  formatTrancheStateSheet(stateSheet);
  addHeaderNotes(stateSheet, notesMap);
  filterHeaders(stateSheet);

  // percent-format ConBasisPct
  const pctCol = output[0].indexOf("ConBasisPct") + 1;
  if (pctCol > 0) {
    stateSheet
      .getRange(2, pctCol, output.length - 1, 1)
      .setNumberFormat("0.00%");
  }

  autoSizeAllColumns(stateSheet, 4);
  freezeHeaders(stateSheet);

  // reapply your red gradient
  addConBasisPctColorScale(stateSheet);
}


function addConBasisPctColorScale(sheet) {
  // 1) Find header row and locate ConBasisPct
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colIndex = headers.indexOf("ConBasisPct") + 1; // 1-based
  
  if (colIndex < 1) return;  // no ConBasisPct column, bail out
  
  // 2) Define the data range (excluding header)
  const lastRow = sheet.getLastRow();
  const dataRange = sheet.getRange(2, colIndex, lastRow - 1, 1);
  
  // 3) Build gradient rule: white→pink→dark red
  const gradientRule = SpreadsheetApp
    .newConditionalFormatRule()
    .setGradientMinpointWithValue("#FFFFFF", SpreadsheetApp.InterpolationType.NUMBER, 0)
    .setGradientMidpointWithValue("#FFCCCC", SpreadsheetApp.InterpolationType.NUMBER, 0.5)
    .setGradientMaxpointWithValue("#FF0000", SpreadsheetApp.InterpolationType.NUMBER, 1)
    .setRanges([dataRange])
    .build();
  
  // 4) Append and reapply all rules
  const rules = sheet.getConditionalFormatRules();
  rules.push(gradientRule);
  sheet.setConditionalFormatRules(rules);
}



/**
 * Formats the TrancheState sheet: dates, numbers, currency.
 */
function formatTrancheStateSheet(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const colMap  = headers.reduce((m,h,i) => (m[h]=i+1, m), {});
  const numRows = sheet.getLastRow();

  const dateCols     = ["BuyDt"];
  const currencyCols = ["BuyPx","SellPx","CurrPx","RemBasis","AdjBasis","ConBasis","MktVal","UnrlGain","ExitGain","RlzGain"];
  const numCols      = ["ShBuy","ShSold","ShRem"];
  const intCols      = ["HeldDays","LTGDays"];

  dateCols.forEach(c => {
    if (colMap[c]) sheet.getRange(2, colMap[c], numRows-1).setNumberFormat("yyyy-MM-dd");
  });
  currencyCols.forEach(c => {
    if (colMap[c]) sheet.getRange(2, colMap[c], numRows-1).setNumberFormat("$#,##0.00");
  });
  numCols.forEach(c => {
    if (colMap[c]) sheet.getRange(2, colMap[c], numRows-1).setNumberFormat("0.00");
  });
  intCols.forEach(c => {
    if (colMap[c]) sheet.getRange(2, colMap[c], numRows-1).setNumberFormat("0");
  });
}