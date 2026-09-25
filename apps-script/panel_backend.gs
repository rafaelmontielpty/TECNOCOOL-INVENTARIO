function openPanel() {
  const html = HtmlService.createHtmlOutputFromFile("panel").setWidth(400).setHeight(600);
  SpreadsheetApp.getUi().showSidebar(html);
}

function processSale(sale) {
  try {
    registerSale(sale);
    return "Venta registrada correctamente.";
  } catch (err) {
    throw new Error(err);
  }
}
