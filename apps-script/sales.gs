/*******************************
 * AUTOMATIZACIÓN DE CLIENTES
 * (sin cambios)
 *******************************/
function ensureCustomer(customer_id, user) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const customersSheet = ss.getSheetByName('CUSTOMERS');
  const data = customersSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === customer_id) {
      return data[i][2]; // type
    }
  }

  const now = new Date();
  const defaultType = "Cliente final";

  customersSheet.appendRow([
    customer_id,
    "Cliente creado automáticamente",
    defaultType,
    "",
    "",
    "",
    now
  ]);

  return defaultType;
}


/*******************************
 * AUTOMATIZACIÓN DE PRODUCTOS
 * CAMBIO (v2): ahora también devuelve "status". registerSale() usa
 * ese status para bloquear la venta si el producto quedó
 * "Pendiente Precio" (por ejemplo, tras una carga masiva sin costo
 * ni precio todavía completados).
 *******************************/
function ensureProduct(product_id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const productsSheet = ss.getSheetByName('PRODUCTS');
  const data = productsSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === product_id) {
      return {
        cost: data[i][8],
        priceCustomer: data[i][9],
        priceTechnician: data[i][10],
        priceWholesale: data[i][11],
        status: data[i][16]
      };
    }
  }

  throw new Error(
    "El producto '" + product_id + "' no existe en PRODUCTS. " +
    "Debe crearse (nombre, marca, modelo, categoría) antes de poder venderse."
  );
}


/*******************************
 * AUTOMATIZACIÓN DE STOCK
 * (sin cambios)
 *******************************/
function ensureStock(product_id, warehouse_id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const stockSheet = ss.getSheetByName('STOCK');
  const data = stockSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][1] === product_id && data[i][2] === warehouse_id) {
      return {
        row: i + 1,
        quantity: data[i][3]
      };
    }
  }

  const now = new Date();
  const stock_id = "STOCK-" + new Date().getTime();

  stockSheet.appendRow([
    stock_id,
    product_id,
    warehouse_id,
    0,
    now
  ]);

  return {
    row: stockSheet.getLastRow(),
    quantity: 0
  };
}


/*******************************
 * REGISTRO DE VENTAS
 * CAMBIO (v2): si el producto existe pero su status no es "Activo"
 * (ej. "Pendiente Precio" recién cargado sin costo/precio), la venta
 * se bloquea con un mensaje claro, igual que cuando el producto no
 * existe del todo.
 *******************************/
function registerSale(sale) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const {
      sale_id,
      sale_type,
      customer_id,
      product_id,
      quantity,
      warehouse_id,
      user,
      created_at
    } = sale;

    /********** CLIENTE **********/
    let customerType = ensureCustomer(customer_id, user);

    /********** PRODUCTO **********/
    const product = ensureProduct(product_id); // lanza error si no existe

    if (product.status && product.status !== 'Activo') {
      throw new Error(
        "El producto '" + product_id + "' está en estado '" + product.status +
        "' y no se puede vender todavía. Completa su costo y precio en PRODUCTS y cambia su status a 'Activo'."
      );
    }

    /********** SELECCIÓN DE PRECIO **********/
    let unit_price = 0;
    if (customerType === "Cliente final") unit_price = product.priceCustomer;
    if (customerType === "Técnico") unit_price = product.priceTechnician;
    if (customerType === "Mayorista") unit_price = product.priceWholesale;

    const total_price = unit_price * quantity;
    const unit_cost = product.cost;
    const total_cost = unit_cost * quantity;

    /********** STOCK **********/
    const stockInfo = ensureStock(product_id, warehouse_id);
    const currentStock = stockInfo.quantity;

    if (currentStock < quantity) {
      throw new Error("Stock insuficiente. Disponible: " + currentStock);
    }

    /********** REGISTRAR VENTA **********/
    const salesSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('SALES');
    salesSheet.appendRow([
      sale_id,
      sale_type,
      customer_id,
      product_id,
      quantity,
      unit_price,
      total_price,
      warehouse_id,
      "",
      "",
      created_at
    ]);

    /********** REGISTRAR MOVIMIENTO **********/
    registerMovement({
      movement_id: "MOV-" + new Date().getTime(),
      movement_type: "Venta",
      product_id,
      warehouse_id,
      quantity: -quantity,
      unit_cost: unit_cost,
      total_cost: total_cost,
      reference: sale_id,
      user,
      created_at
    });

  } finally {
    lock.releaseLock();
  }
}
