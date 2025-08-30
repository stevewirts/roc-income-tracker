function checkPriceUpdates() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("MarketData");
  const dataRange = sheet.getRange("B2:B"); // CurrentPrice column
  const lastUpdateRange = sheet.getRange("C2:C"); // LastUpdate column
  const prevRange = sheet.getRange("AA2:AA"); // Stored previous prices

  const currentPrices = dataRange.getValues();
  const prevPrices = prevRange.getValues();

  const now = new Date();

  for (let i = 0; i < currentPrices.length; i++) {
    const current = currentPrices[i][0];
    const prev = prevPrices[i][0];

    if (current !== "" && current !== prev) {
      lastUpdateRange.getCell(i + 1, 1).setValue(now);
      prevRange.getCell(i + 1, 1).setValue(current);
    }
  }
}
