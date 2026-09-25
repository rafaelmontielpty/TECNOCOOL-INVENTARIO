/*******************************
 * CARGA DE INVENTARIO INICIAL (Paso 4, v3)
 *
 * CAMBIO vs v2: v2 se quedaba a medio camino (timeout de 6 min de
 * Apps Script) porque releía las hojas PRODUCTS/STOCK/MOVEMENTS
 * completas en cada fila (manufacturerCodeExists_,
 * getNextProductIdNumber_, registerMovement). Con 142+ filas eso
 * son cientos de lecturas contra el servidor.
 *
 * v3 lee PRODUCTS y WAREHOUSES UNA sola vez en memoria, valida y
 * arma ahí los nuevos productos/stock/movimientos, y escribe todo
 * en bloque (una sola escritura por hoja) al final. Es mucho más
 * rápido y, si aun así se acerca al límite de tiempo, se detiene
 * de forma segura y avisa que se debe volver a correr: como
 * respeta el estado "Procesado" de cada fila, re-ejecutar el menú
 * continúa donde quedó, sin duplicar nada.
 *
 * Mantiene el mismo comportamiento de negocio que v2:
 * - cost / price_customer son opcionales -> status "Pendiente Precio"
 *   si faltan (y bloqueado para venta, ver sales.gs).
 *******************************/

const IMPORT_SHEET_NAME = 'IMPORTAR_STOCK_INICIAL';

const IMPORT_COLS = {
  manufacturer_code: 0,
  name: 1,
  description: 2,
  category: 3,
  subcategory: 4,
  brand: 5,
  model: 6,
  cost: 7,
  price_customer: 8,
  price_technician: 9,
  price_wholesale: 10,
  min_stock: 11,
  max_stock: 12,
  reorder_point: 13,
  has_serial: 14,
  warehouse_id: 15,
  quantity_initial: 16,
  estado: 17,
  product_id_generado: 18
};

// Margen de seguridad bajo el límite real de 6 minutos de Apps Script.
const IMPORT_MAX_MS = 4.5 * 60 * 1000;

function importInitialInventory() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  const startTime = new Date().getTime();

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const importSheet = ss.getSheetByName(IMPORT_SHEET_NAME);
    if (!importSheet) {
      throw new Error('No existe la hoja "' + IMPORT_SHEET_NAME + '". Impórtala primero (ver plantilla).');
    }

    const data = importSheet.getDataRange().getValues();

    // ---- Lecturas únicas (en memoria de aquí en adelante) ----
    const warehousesSheet = ss.getSheetByName('WAREHOUSES');
    const warehouseData = warehousesSheet.getDataRange().getValues();
    const warehouseIds = [];
    for (let w = 1; w < warehouseData.length; w++) {
      if (warehouseData[w][0]) warehouseIds.push(String(warehouseData[w][0]).trim());
    }

    const productsSheet = ss.getSheetByName('PRODUCTS');
    const productsData = productsSheet.getDataRange().getValues();
    const existingCodes = {}; // manufacturer_code normalizado -> true
    let maxSkuNumber = 0;
    for (let p = 1; p < productsData.length; p++) {
      const code = String(productsData[p][1] || '').trim().toUpperCase();
      if (code) existingCodes[code] = true;
      const match = String(productsData[p][0] || '').match(/^SKU-(\d+)$/);
      if (match) {
        const n = parseInt(match[1], 10);
        if (n > maxSkuNumber) maxSkuNumber = n;
      }
    }

    const stockSheet = ss.getSheetByName('STOCK');
    const movementsSheet = ss.getSheetByName('MOVEMENTS');
    const auditSheet = ss.getSheetByName('AUDIT_LOG');
    const alertsSheet = ss.getSheetByName('REORDER_ALERTS');
    const user = Session.getActiveUser().getEmail() || 'Sistema';

    // ---- Acumuladores para escritura en bloque ----
    const newProductRows = [];
    const newStockRows = [];
    const newMovementRows = [];
    const newAuditRows = [];
    const newAlertRows = [];
    const estadoUpdates = []; // { row, estado, productId }

    let procesados = 0;
    let pendientesPrecio = 0;
    let errores = 0;
    let saltados = 0;
    let timedOut = false;

    for (let i = 1; i < data.length; i++) {
      if (new Date().getTime() - startTime > IMPORT_MAX_MS) {
        timedOut = true;
        break;
      }

      const row = data[i];
      const rowNum = i + 1;

      const estadoActual = String(row[IMPORT_COLS.estado] || '').trim();
      if (estadoActual === 'Procesado') { saltados++; continue; }

      const isBlank = row.every(function (v) { return v === '' || v === null; });
      if (isBlank) continue;

      const name = String(row[IMPORT_COLS.name] || '').trim();
      const category = String(row[IMPORT_COLS.category] || '').trim();
      const brand = String(row[IMPORT_COLS.brand] || '').trim();
      const warehouseId = String(row[IMPORT_COLS.warehouse_id] || '').trim();
      const quantityInitial = Number(row[IMPORT_COLS.quantity_initial]);

      const costRaw = row[IMPORT_COLS.cost];
      const priceRaw = row[IMPORT_COLS.price_customer];
      const hasCost = costRaw !== '' && costRaw !== null && !isNaN(Number(costRaw));
      const hasPrice = priceRaw !== '' && priceRaw !== null && !isNaN(Number(priceRaw));
      const cost = hasCost ? Number(costRaw) : 0;
      const priceCustomer = hasPrice ? Number(priceRaw) : 0;

      const faltantes = [];
      if (!name) faltantes.push('name');
      if (!category) faltantes.push('category');
      if (!brand) faltantes.push('brand');
      if (!warehouseId) faltantes.push('warehouse_id');
      if (isNaN(quantityInitial)) faltantes.push('quantity_initial');

      if (faltantes.length) {
        estadoUpdates.push({ row: rowNum, estado: 'Error: falta ' + faltantes.join(', ') });
        errores++;
        continue;
      }

      if (warehouseIds.indexOf(warehouseId) === -1) {
        estadoUpdates.push({ row: rowNum, estado: 'Error: warehouse_id "' + warehouseId + '" no existe en WAREHOUSES' });
        errores++;
        continue;
      }

      const manufacturerCode = String(row[IMPORT_COLS.manufacturer_code] || '').trim();
      const normCode = manufacturerCode.toUpperCase();
      if (manufacturerCode && existingCodes[normCode]) {
        estadoUpdates.push({ row: rowNum, estado: 'Error: manufacturer_code "' + manufacturerCode + '" ya existe en PRODUCTS' });
        errores++;
        continue;
      }

      maxSkuNumber++;
      const productId = 'SKU-' + String(maxSkuNumber).padStart(4, '0');

      const priceTechnician = hasPrice && row[IMPORT_COLS.price_technician] !== '' && !isNaN(Number(row[IMPORT_COLS.price_technician]))
        ? Number(row[IMPORT_COLS.price_technician]) : priceCustomer;
      const priceWholesale = hasPrice && row[IMPORT_COLS.price_wholesale] !== '' && !isNaN(Number(row[IMPORT_COLS.price_wholesale]))
        ? Number(row[IMPORT_COLS.price_wholesale]) : priceCustomer;
      const minStock = row[IMPORT_COLS.min_stock] !== '' && !isNaN(Number(row[IMPORT_COLS.min_stock]))
        ? Number(row[IMPORT_COLS.min_stock]) : 0;
      const maxStock = row[IMPORT_COLS.max_stock] !== '' && !isNaN(Number(row[IMPORT_COLS.max_stock]))
        ? Number(row[IMPORT_COLS.max_stock]) : '';
      const reorderPoint = row[IMPORT_COLS.reorder_point] !== '' && !isNaN(Number(row[IMPORT_COLS.reorder_point]))
        ? Number(row[IMPORT_COLS.reorder_point]) : minStock;
      const hasSerialRaw = String(row[IMPORT_COLS.has_serial] || '').trim().toLowerCase();
      const hasSerial = (hasSerialRaw === 'si' || hasSerialRaw === 'sí' || hasSerialRaw === 'yes') ? 'Sí' : 'No';

      const status = (hasCost && hasPrice) ? 'Activo' : 'Pendiente Precio';
      if (status === 'Pendiente Precio') pendientesPrecio++;

      const now = new Date();

      newProductRows.push([
        productId, manufacturerCode, name,
        String(row[IMPORT_COLS.description] || '').trim(),
        category,
        String(row[IMPORT_COLS.subcategory] || '').trim(),
        brand,
        String(row[IMPORT_COLS.model] || '').trim(),
        cost, priceCustomer, priceTechnician, priceWholesale,
        minStock, maxStock, reorderPoint, hasSerial, status, now, now
      ]);

      if (manufacturerCode) existingCodes[normCode] = true;

      if (quantityInitial > 0) {
        newStockRows.push([
          'STOCK-' + productId, productId, warehouseId, quantityInitial, now
        ]);
        newMovementRows.push([
          'MOV-' + productId, 'Ajuste Inventario Inicial', productId, warehouseId,
          quantityInitial, cost, cost * quantityInitial, 'CARGA_INICIAL', user, now
        ]);
        newAuditRows.push([
          'AUD-' + productId, user, 'CREATE_STOCK', 'STOCK',
          productId + '|' + warehouseId, '', quantityInitial, now
        ]);
        if (quantityInitial < minStock) {
          newAlertRows.push([
            'ALERT-' + productId + '-MIN', productId, warehouseId,
            quantityInitial, minStock, reorderPoint, 'MIN_STOCK', now, 'Pendiente'
          ]);
        }
        if (quantityInitial < reorderPoint) {
          newAlertRows.push([
            'ALERT-' + productId + '-REO', productId, warehouseId,
            quantityInitial, minStock, reorderPoint, 'REORDER_POINT', now, 'Pendiente'
          ]);
        }
      }

      estadoUpdates.push({ row: rowNum, estado: 'Procesado', productId: productId });
      procesados++;
    }

    // ---- Escrituras en bloque ----
    if (newProductRows.length) {
      productsSheet.getRange(
        productsSheet.getLastRow() + 1, 1,
        newProductRows.length, newProductRows[0].length
      ).setValues(newProductRows);
    }
    if (newStockRows.length) {
      stockSheet.getRange(
        stockSheet.getLastRow() + 1, 1,
        newStockRows.length, newStockRows[0].length
      ).setValues(newStockRows);
    }
    if (newMovementRows.length) {
      movementsSheet.getRange(
        movementsSheet.getLastRow() + 1, 1,
        newMovementRows.length, newMovementRows[0].length
      ).setValues(newMovementRows);
    }
    if (newAuditRows.length) {
      auditSheet.getRange(
        auditSheet.getLastRow() + 1, 1,
        newAuditRows.length, newAuditRows[0].length
      ).setValues(newAuditRows);
    }
    if (newAlertRows.length) {
      alertsSheet.getRange(
        alertsSheet.getLastRow() + 1, 1,
        newAlertRows.length, newAlertRows[0].length
      ).setValues(newAlertRows);
    }
    estadoUpdates.forEach(function (u) {
      importSheet.getRange(u.row, IMPORT_COLS.estado + 1).setValue(u.estado);
      if (u.productId) {
        importSheet.getRange(u.row, IMPORT_COLS.product_id_generado + 1).setValue(u.productId);
      }
    });

    let resumen = 'Procesados: ' + procesados +
      ' (de los cuales ' + pendientesPrecio + ' quedaron "Pendiente Precio"). ' +
      'Errores: ' + errores + '. Ya procesados antes (saltados): ' + saltados + '.';

    if (timedOut) {
      resumen += '\n\nSe llegó al tiempo máximo de ejecución antes de terminar todas las filas. ' +
        'Vuelve a correr "Cargar Inventario Inicial" desde el menú TECNOCOOL para continuar ' +
        'con las filas restantes (no se duplica nada, ya que las filas procesadas quedan marcadas).';
    }

    SpreadsheetApp.getUi().alert(resumen);
    return resumen;

  } finally {
    lock.releaseLock();
  }
}

function getValidWarehouseIds_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('WAREHOUSES');
  const data = sheet.getDataRange().getValues();
  const ids = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][0]) ids.push(String(data[i][0]).trim());
  }
  return ids;
}

function manufacturerCodeExists_(code) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('PRODUCTS');
  const data = sheet.getDataRange().getValues();
  const target = code.trim().toUpperCase();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1] || '').trim().toUpperCase() === target) return true;
  }
  return false;
}

function getNextProductIdNumber_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('PRODUCTS');
  const data = sheet.getDataRange().getValues();
  let maxNumber = 0;
  for (let i = 1; i < data.length; i++) {
    const id = String(data[i][0] || '');
    const match = id.match(/^SKU-(\d+)$/);
    if (match) {
      const n = parseInt(match[1], 10);
      if (n > maxNumber) maxNumber = n;
    }
  }
  return maxNumber + 1;
}
