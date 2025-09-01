/**
 * buildTrancheState.gs
 *
 * Populates the “TrancheState” sheet with up-to-date position data,
 * adjusts basis for return-of-capital, computes unrealized gain/loss,
 * and applies conditional formatting.
 */
function buildTrancheState() {
  const ss         = SpreadsheetApp.getActive();
  const txSheet    = ss.getSheetByName('Transactions');
  const stateSheet = insureClearedSheet('TrancheState');
  const today      = new Date();

  // 1) Read & normalize headers
  const allRows = txSheet.getDataRange().getValues();
  const rawHdr  = allRows.shift();
  const hdr     = rawHdr.map(h => String(h || '').trim());
  const idx     = hdr.reduce((map, h, i) => ((map[h] = i), map), {});

  // 2) Validate required columns
  ['Type','TrID','Sym','Date','Shr','Price','ROCAmt','Inc']
    .forEach(col => {
      if (!(col in idx)) {
        throw new Error(
          `Missing Transactions header "${col}". Found: [${Object.keys(idx).join(', ')}]`
        );
      }
    });

  // 3) Initialize tranche & distribution maps
  const trancheMap = {};  // { TrID → tranche state }
  const rocMap     = {};  // { TrID → cumulative ROC }
  const incMap     = {};  // { TrID → cumulative non-ROC income }

  // 4) Bucket each transaction row
  allRows.forEach(row => {
    const type = String(row[idx.Type] || '').toLowerCase();

    // 4a) Dividends distribute ROC + income across all open tranches of that symbol
    if (type === 'dividend') {
      const rawROC = parseFloat(String(row[idx.ROCAmt]).replace(/[^0-9.\-]/g, '')) || 0;
      const rawInc = parseFloat(String(row[idx.Inc]).replace(/[^0-9.\-]/g, ''))   || 0;
      const sym    = row[idx.Sym];

      const openIDs = Object.keys(trancheMap).filter(id => {
        const t   = trancheMap[id];
        const rem = t.ShBuy - t.ShSold;
        return t.Sym === sym && rem > 0;
      });
      const totalRem = openIDs.reduce(
        (sum, id) => sum + (trancheMap[id].ShBuy - trancheMap[id].ShSold),
        0
      );

      if (totalRem > 0) {
        openIDs.forEach(id => {
          const t     = trancheMap[id];
          const rem   = t.ShBuy - t.ShSold;
          const share = rem / totalRem;
          rocMap[id]  = (rocMap[id]  || 0) + rawROC * share;
          incMap[id]  = (incMap[id]  || 0) + rawInc * share;
        });
      }
      return;
    }

    // 4b) Buys & sells (requires TrID)
    const tid = row[idx.TrID];
    if (!tid) return;

    if (!trancheMap[tid]) {
      trancheMap[tid] = {
        ID:        tid,
        Sym:       row[idx.Sym],
        BuyDt:     null,
        ShBuy:     0,
        BuyPx:     0,
        ShSold:    0,
        SellPx:    0,
        CostBasis: 0,
        Status:    ''
      };
    }
    const t = trancheMap[tid];

    if (type === 'buy') {
      const qty   = +row[idx.Shr] || 0;
      const price = parseFloat(String(row[idx.Price]).replace(/[^0-9.\-]/g, '')) || 0;
      t.ShBuy      += qty;
      t.CostBasis  += qty * price;
      t.BuyPx       = price || t.BuyPx;
      const dt      = new Date(row[idx.Date]);
      t.BuyDt       = !t.BuyDt || dt < t.BuyDt ? dt : t.BuyDt;
      t.Status      = row[idx.TStat] || t.Status;
    }
    else if (type === 'sell') {
      const qty   = +row[idx.Shr] || 0;
      const price = parseFloat(String(row[idx.Price]).replace(/[^0-9.\-]/g, '')) || 0;
      t.ShSold    += qty;
      t.SellPx     = price || t.SellPx;
      t.Status     = row[idx.TStat] || t.Status;
    }
  });

  // 5) Assemble output rows
  const notesMap = {
    ID:                "Unique tranche identifier",
    Sym:               "Ticker symbol",
    BuyDt:             "Date shares were purchased",
    ShBuy:             "Number of shares bought",
    BuyPx:             "Average purchase price",
    ShSold:            "Shares sold",
    SellPx:            "Average sale price",
    ShRem:             "Remaining shares",
    CurrPx:            "Current market price",
    CostBasis:         "PurchasePrice × SharesBought",
    ROC:               "Return of capital allocated to tranche",
    AdjBasis:          "CostBasis − ROC",
    CumIncome:         "Non-ROC distributions allocated to tranche",
    MktValue:          "ShRem × CurrPx",
    UnrealizedGainLoss:"ShRem × CurrPx − AdjBasis",
    PctToExit:         "PctToExit = (AdjBasis − MktValue) / AdjBasis; 0% = break-even, 100% = full loss of principal",
    Status:            "Open, Partial, or Closed",
    HeldDays:          "Days since BuyDt"
  };
  const keys = Object.keys(notesMap);
  const out  = [keys];

  Object.values(trancheMap).forEach(t => {
    const rem        = t.ShBuy - t.ShSold;
    const roc        = rocMap[t.ID] || 0;
    const inc        = incMap[t.ID] || 0;
    const costBasis  = t.CostBasis;
    const adjBasis   = costBasis - roc;
    const currPx     = priceGet(t.Sym);
    const mktValue   = rem * currPx;
    const unrealGain = mktValue - adjBasis;

    // PctToExit ignores CumIncome, clamped 0–1
    const rawPct    = adjBasis ? (adjBasis - mktValue) / adjBasis : 0;
    const pctToExit = Math.max(0, Math.min(1, rawPct));

    const heldDays = t.BuyDt
      ? Math.floor((today - t.BuyDt) / 86400000)
      : '';

    const row = {
      ID:                t.ID,
      Sym:               t.Sym,
      BuyDt:             formatDate_(t.BuyDt),
      ShBuy:             t.ShBuy,
      BuyPx:             t.BuyPx,
      ShSold:            t.ShSold,
      SellPx:            t.SellPx || '',
      ShRem:             rem,
      CurrPx:            currPx,
      CostBasis:         costBasis,
      ROC:               roc,
      AdjBasis:          adjBasis,
      CumIncome:         inc,
      MktValue:          mktValue,
      UnrealizedGainLoss:unrealGain,
      PctToExit:         pctToExit,
      Status:            t.Status,
      HeldDays:          heldDays
    };
    out.push(keys.map(k => row[k]));
  });

  // 6) Write & format output
  stateSheet.clearContents();
  stateSheet
    .getRange(1, 1, out.length, keys.length)
    .setValues(out);

  addHeaderNotes(stateSheet, notesMap);
  filterHeaders(stateSheet);

  const colOf   = name => keys.indexOf(name) + 1;
  const fmtDate = "yyyy-MM-dd";
  const fmtCur  = "$#,##0.00";
  const fmtPct  = "0.00%";
  const fmtInt  = "0";

  // Date formatting
  stateSheet
    .getRange(2, colOf("BuyDt"), out.length - 1)
    .setNumberFormat(fmtDate);

  // Currency formatting
  [
    "BuyPx", "SellPx", "CostBasis", "ROC",
    "AdjBasis", "CumIncome", "MktValue", "UnrealizedGainLoss"
  ].forEach(name =>
    stateSheet
      .getRange(2, colOf(name), out.length - 1)
      .setNumberFormat(fmtCur)
  );

  // Integer formatting
  ["ShBuy","ShSold","ShRem","HeldDays"].forEach(name =>
    stateSheet
      .getRange(2, colOf(name), out.length - 1)
      .setNumberFormat(fmtInt)
  );

  // Percentage formatting
  stateSheet
    .getRange(2, colOf("PctToExit"), out.length - 1)
    .setNumberFormat(fmtPct);

  // Layout refinements
  autoSizeAllColumns(stateSheet, 4);
  freezeHeaders(stateSheet);

  // 7) Apply green-gradient to PctToExit (white @ 0%, green @ 100%)
  const pctCol = colOf("PctToExit");
  const lastRow = stateSheet.getLastRow();
  if (lastRow > 1) {
    const range = stateSheet.getRange(2, pctCol, lastRow - 1, 1);
    const rules = stateSheet.getConditionalFormatRules();
    const pctRule = SpreadsheetApp.newConditionalFormatRule()
      .setGradientMinpointWithValue(
        "#ffffff",
        SpreadsheetApp.InterpolationType.NUMBER,
        0
      )
      .setGradientMaxpointWithValue(
        "#1a9850",
        SpreadsheetApp.InterpolationType.NUMBER,
        1
      )
      .setRanges([range])
      .build();
    rules.push(pctRule);
    stateSheet.setConditionalFormatRules(rules);
  }
}
