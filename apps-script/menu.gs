function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("TECNOCOOL")
    .addItem("Abrir Panel de Ventas", "openPanel")
    .addItem("Buscar Repuesto", "openSearchPanel")
    .addSeparator()
    .addItem("Cargar Inventario Inicial", "importInitialInventory")
    .addToUi();
}

function openSearchPanel() {
  const html = HtmlService.createHtmlOutputFromFile("panel_busqueda")
    .setWidth(420)
    .setHeight(650);
  SpreadsheetApp.getUi().showSidebar(html);
}
