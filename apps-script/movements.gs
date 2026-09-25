/*******************************
 * MOVIMIENTOS DE INVENTARIO
 * CAMBIO: antes, cuando se creaba stock por primera vez (producto
 * nuevo en ese almacén), la función terminaba con un "return" antes
 * de llegar a checkReorder() — es decir, nunca se validaba alerta de
 * mínimo/reorden en la primera carga. Ahora checkReorder() se llama
 * siempre, tanto en creación como en actualización de stock.
 *******************************/
function registerMovement(movement) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stockSheet = ss.getSheetByName('STOCK');
  const movementSheet = ss.getSheetByName('MOVEMENTS');
  const auditSheet = ss.getSheetByName('AUDIT_LOG');

  const {
    movement_id,
    movement_type,
    product_id,
    warehouse_id,
    quantity,
    unit_cost,
    total_cost,
    reference,
    user,
    created_at
  } = movement;

  // Registrar movimiento en MOVEMENTS
  movementSheet.appendRow([
    movement_id,
    movement_type,
    product_id,
    warehouse_id,
    quantity,
    unit_cost,
    total_cost,
    reference,
    user,
    created_at
  ]);

  // Buscar stock existente
  const stockData = stockSheet.getDataRange().getValues();
  let stockRow = null;

  for (let i = 1; i < stockData.length; i++) {
    if (stockData[i][1] === product_id && stockData[i][2] === warehouse_id) {
      stockRow = i + 1;
      break;
    }
  }

  let newQuantity;

  if (!stockRow) {
    // No existía stock: se crea con la cantidad del movimiento
    newQuantity = quantity;

    stockSheet.appendRow([
      'STOCK-' + new Date().getTime(),
      product_id,
      warehouse_id,
      newQuantity,
      created_at
    ]);

    auditSheet.appendRow([
      'AUD-' + new Date().getTime(),
      user,
      'CREATE_STOCK',
      'STOCK',
      product_id + '|' + warehouse_id,
      '',
      newQuantity,
      created_at
    ]);

  } else {
    // Ya existía stock: se actualiza la cantidad
    const oldQuantity = stockSheet.getRange(stockRow, 4).getValue();
    newQuantity = oldQuantity + quantity;

    stockSheet.getRange(stockRow, 4).setValue(newQuantity);
    stockSheet.getRange(stockRow, 5).setValue(created_at);

    auditSheet.appendRow([
      'AUD-' + new Date().getTime(),
      user,
      'UPDATE_STOCK',
      'STOCK',
      product_id + '|' + warehouse_id,
      oldQuantity,
      newQuantity,
      created_at
    ]);
  }

  // ANTES: esta llamada quedaba fuera de alcance cuando el stock se
  // creaba por primera vez, porque el bloque de arriba tenía un
  // "return" propio. Ahora siempre se ejecuta.
  checkReorder(product_id, warehouse_id, newQuantity, user);
}


function checkReorder(product_id, warehouse_id, newQuantity, user) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const productsSheet = ss.getSheetByName('PRODUCTS');
  const alertsSheet = ss.getSheetByName('REORDER_ALERTS');

  const productsData = productsSheet.getDataRange().getValues();

  let minStock = null;
  let reorderPoint = null;

  for (let i = 1; i < productsData.length; i++) {
    if (productsData[i][0] === product_id) {
      minStock = productsData[i][12];        // min_stock
      reorderPoint = productsData[i][14];    // reorder_point
      break;
    }
  }

  if (minStock === null) return; // Producto no encontrado

  const now = new Date();

  if (newQuantity < minStock) {
    alertsSheet.appendRow([
      'ALERT-' + new Date().getTime(),
      product_id,
      warehouse_id,
      newQuantity,
      minStock,
      reorderPoint,
      'MIN_STOCK',
      now,
      'Pendiente'
    ]);
  }

  if (newQuantity < reorderPoint) {
    alertsSheet.appendRow([
      'ALERT-' + new Date().getTime(),
      product_id,
      warehouse_id,
      newQuantity,
      minStock,
      reorderPoint,
      'REORDER_POINT',
      now,
      'Pendiente'
    ]);
  }
}
