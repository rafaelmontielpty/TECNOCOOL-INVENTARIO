/*******************************
 * BÚSQUEDA DE CATÁLOGO (Paso 2)
 * Busca por modelo de equipo o por número/nombre de repuesto en las
 * hojas EQUIPMENT_MODELS / PARTS / MODEL_PART_MAP (catálogo Lennox
 * importado), y cruza cada resultado contra el inventario real de
 * Tecnocool: si PRODUCTS.manufacturer_code coincide con el part_code
 * del catálogo, muestra si lo tienes en existencia y cuánto stock
 * hay (sumado entre todos los almacenes).
 *
 * NOTA (Fase 1 / Sheets): cada búsqueda relee las hojas completas.
 * Con ~51,000 filas en MODEL_PART_MAP esto puede tardar unos
 * segundos, no es instantáneo. Es una limitación conocida de Sheets
 * y una de las razones para migrar a una base de datos real en la
 * Fase 2 del proyecto.
 *******************************/

function normalizeCode_(v) {
  return String(v || '').trim().toUpperCase();
}

/** Índice de PRODUCTS por manufacturer_code normalizado */
function buildProductIndexByManufacturerCode_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('PRODUCTS');
  const data = sheet.getDataRange().getValues();
  const index = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const manufacturerCode = normalizeCode_(row[1]); // manufacturer_code
    if (!manufacturerCode) continue;
    index[manufacturerCode] = {
      product_id: row[0],
      name: row[2],
      brand: row[6],
      model: row[7],
      status: row[16]
    };
  }
  return index;
}

/** Índice de STOCK total (suma todos los almacenes) por product_id */
function buildStockIndexByProductId_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('STOCK');
  const data = sheet.getDataRange().getValues();
  const index = {};
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const productId = row[1];
    const qty = Number(row[3]) || 0;
    if (!productId) continue;
    index[productId] = (index[productId] || 0) + qty;
  }
  return index;
}

/** Cruza un part_code del catálogo contra el inventario real de Tecnocool */
function crossReferenceStock_(partCode, productIndex, stockIndex) {
  const code = normalizeCode_(partCode);
  const product = productIndex[code];
  if (!product) {
    return { en_catalogo_tecnocool: false, stock: 0 };
  }
  return {
    en_catalogo_tecnocool: true,
    product_id: product.product_id,
    nombre_interno: product.name,
    status: product.status,
    stock: stockIndex[product.product_id] || 0
  };
}

/**
 * Busca modelos de equipo por coincidencia parcial (case-insensitive)
 * en model_code. Devuelve como máximo 30 resultados.
 */
function searchEquipmentModels(query) {
  const q = normalizeCode_(query);
  if (!q) return [];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('EQUIPMENT_MODELS');
  const data = sheet.getDataRange().getValues();

  const results = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const modelCode = normalizeCode_(row[2]);
    if (modelCode.indexOf(q) !== -1) {
      results.push({
        model_id: row[0],
        model_code: row[2],
        category_code: row[3],
        range_description: row[4],
        line: row[5]
      });
      if (results.length >= 30) break;
    }
  }
  return results;
}

/**
 * Dado un model_id (de EQUIPMENT_MODELS), devuelve todos los
 * repuestos de su despiece (MODEL_PART_MAP + PARTS), cruzados con
 * el stock real, ordenados por posición en el diagrama de despiece.
 */
function getPartsForModel(model_id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const partsSheet = ss.getSheetByName('PARTS');
  const partsData = partsSheet.getDataRange().getValues();
  const partsIndex = {};
  for (let i = 1; i < partsData.length; i++) {
    const row = partsData[i];
    partsIndex[row[0]] = { part_code: row[1], part_name: row[2], category: row[4] };
  }

  const mapSheet = ss.getSheetByName('MODEL_PART_MAP');
  const mapData = mapSheet.getDataRange().getValues();

  const productIndex = buildProductIndexByManufacturerCode_();
  const stockIndex = buildStockIndexByProductId_();

  const results = [];
  for (let i = 1; i < mapData.length; i++) {
    const row = mapData[i];
    const rowModelId = row[1];
    if (rowModelId !== model_id) continue;

    const partId = row[2];
    const part = partsIndex[partId];
    if (!part) continue;

    const stockInfo = crossReferenceStock_(part.part_code, productIndex, stockIndex);

    results.push({
      exploded_view_no: row[3],
      qty: row[4],
      part_code: part.part_code,
      part_name: part.part_name,
      category: part.category,
      en_catalogo_tecnocool: stockInfo.en_catalogo_tecnocool,
      stock: stockInfo.stock || 0
    });
  }

  results.sort(function (a, b) {
    const na = parseFloat(a.exploded_view_no);
    const nb = parseFloat(b.exploded_view_no);
    if (isNaN(na) || isNaN(nb)) return 0;
    return na - nb;
  });

  return results;
}

/**
 * Busca repuestos por coincidencia parcial en part_code o part_name.
 * Devuelve como máximo 50 resultados, cruzados con stock real.
 */
function searchParts(query) {
  const q = normalizeCode_(query);
  if (!q) return [];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('PARTS');
  const data = sheet.getDataRange().getValues();

  const productIndex = buildProductIndexByManufacturerCode_();
  const stockIndex = buildStockIndexByProductId_();

  const results = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const partCode = normalizeCode_(row[1]);
    const partName = normalizeCode_(row[2]);
    if (partCode.indexOf(q) !== -1 || partName.indexOf(q) !== -1) {
      const stockInfo = crossReferenceStock_(row[1], productIndex, stockIndex);
      results.push({
        part_id: row[0],
        part_code: row[1],
        part_name: row[2],
        category: row[4],
        en_catalogo_tecnocool: stockInfo.en_catalogo_tecnocool,
        stock: stockInfo.stock || 0
      });
      if (results.length >= 50) break;
    }
  }
  return results;
}
