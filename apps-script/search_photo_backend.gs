/*******************************
 * BÚSQUEDA POR FOTO (OCR) — Paso 3 (v2)
 * Ahora reconoce en la foto tanto:
 *  a) un MODELO de equipo (ej. placa del equipo) → muestra su
 *     despiece completo, o
 *  b) un NÚMERO DE PARTE (ej. caja/etiqueta del repuesto en tu
 *     inventario) → muestra a qué modelo(s) de equipo pertenece.
 *
 * Sube la foto al OCR nativo de Google Drive, extrae el texto, y
 * busca coincidencias en EQUIPMENT_MODELS y en PARTS. Los archivos
 * temporales que genera el OCR se borran automáticamente al
 * terminar.
 *
 * REQUIERE: servicio avanzado "Drive API" activado en este proyecto
 * de Apps Script (Servicios → + → Drive API).
 *******************************/

function searchByPhoto(base64Data, mimeType, fileName) {
  const text = extractTextFromImage_(base64Data, mimeType, fileName);

  const matchedModels = matchModelsInText_(text);
  const matchedParts = matchPartsInText_(text);

  return {
    texto_extraido: text,
    modelos_encontrados: matchedModels,
    partes_encontradas: matchedParts
  };
}

/** Sube la imagen a Drive pidiendo conversión con OCR, lee el texto resultante, y borra los archivos temporales. */
function extractTextFromImage_(base64Data, mimeType, fileName) {
  const bytes = Utilities.base64Decode(base64Data);
  const blob = Utilities.newBlob(bytes, mimeType, fileName || 'foto.jpg');

  const resource = {
    title: 'OCR_TEMP_' + new Date().getTime(),
    mimeType: MimeType.GOOGLE_DOCS
  };

  const options = {
    ocr: true,
    ocrLanguage: 'es'
  };

  const ocrFile = Drive.Files.insert(resource, blob, options);

  let text = '';
  try {
    const doc = DocumentApp.openById(ocrFile.id);
    text = doc.getBody().getText();
  } finally {
    try { Drive.Files.remove(ocrFile.id); } catch (e) { /* no crítico */ }
  }

  return text;
}

/** Revisa si algún model_code de EQUIPMENT_MODELS aparece dentro del texto detectado. */
function matchModelsInText_(text) {
  const normalizedText = normalizeCode_(text).replace(/[\s\-]/g, '');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('EQUIPMENT_MODELS');
  const data = sheet.getDataRange().getValues();

  const matches = [];
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const modelCode = normalizeCode_(row[2]).replace(/[\s\-]/g, '');
    if (modelCode.length < 5) continue;
    if (normalizedText.indexOf(modelCode) !== -1) {
      matches.push({
        model_id: row[0],
        model_code: row[2],
        category_code: row[3],
        range_description: row[4],
        line: row[5]
      });
    }
  }
  return matches;
}

/** Revisa si algún part_code de PARTS aparece dentro del texto detectado, cruzando con stock real. */
function matchPartsInText_(text) {
  const normalizedText = normalizeCode_(text).replace(/[\s\-]/g, '');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const partsSheet = ss.getSheetByName('PARTS');
  const partsData = partsSheet.getDataRange().getValues();

  const productIndex = buildProductIndexByManufacturerCode_();
  const stockIndex = buildStockIndexByProductId_();

  const matches = [];
  for (let i = 1; i < partsData.length; i++) {
    const row = partsData[i];
    const partCode = normalizeCode_(row[1]).replace(/[\s\-]/g, '');
    if (partCode.length < 7) continue; // evita falsos positivos con códigos muy cortos
    if (normalizedText.indexOf(partCode) !== -1) {
      const stockInfo = crossReferenceStock_(row[1], productIndex, stockIndex);
      matches.push({
        part_id: row[0],
        part_code: row[1],
        part_name: row[2],
        category: row[4],
        en_catalogo_tecnocool: stockInfo.en_catalogo_tecnocool,
        stock: stockInfo.stock || 0
      });
    }
  }
  return matches;
}

/**
 * Dado un part_id, devuelve todos los modelos de equipo que usan
 * ese repuesto (búsqueda inversa sobre MODEL_PART_MAP). Tope de 50
 * modelos por seguridad, ya que algunos repuestos comunes aplican a
 * cientos de modelos.
 */
function getModelsForPart(part_id) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const modelsSheet = ss.getSheetByName('EQUIPMENT_MODELS');
  const modelsData = modelsSheet.getDataRange().getValues();
  const modelIndex = {};
  for (let i = 1; i < modelsData.length; i++) {
    const row = modelsData[i];
    modelIndex[row[0]] = {
      model_code: row[2],
      category_code: row[3],
      range_description: row[4],
      line: row[5]
    };
  }

  const mapSheet = ss.getSheetByName('MODEL_PART_MAP');
  const mapData = mapSheet.getDataRange().getValues();

  const seen = {};
  const results = [];
  let totalCount = 0;

  for (let i = 1; i < mapData.length; i++) {
    const row = mapData[i];
    if (row[2] !== part_id) continue; // columna part_id de MODEL_PART_MAP

    const modelId = row[1];
    if (seen[modelId]) continue;
    seen[modelId] = true;
    totalCount++;

    if (results.length < 50) {
      const model = modelIndex[modelId];
      if (model) {
        results.push({
          model_id: modelId,
          model_code: model.model_code,
          range_description: model.range_description,
          line: model.line
        });
      }
    }
  }

  return { modelos: results, total: totalCount };
}
