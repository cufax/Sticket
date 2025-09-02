// ===============================================================
// GESTOR DE HOJAS DE CÁLCULO 
// ===============================================================


/**
 * Función central y robusta para obtener la hoja de cálculo "Tickets_Master".
 * Si la hoja o la pestaña no existen, las crea y añade los encabezados.
 */
function getOrCreateSheet() {
  const HEADERS = [
    'Timestamp', 'Creado Por', 'Usuario Cliente', 'N° Caso Yoizen',
    'Planilla (Cliente)', 'URL Adjunto', 'Herramienta', 'Pedido de escalamiento (Asunto)',
    'ID Derivacion FAN', 'N° Ticket Escalamiento', 'Ticket Interno', 'Estado', 'Observaciones Auditor', 'Prioridad', 'Tomado Por',
    'Mal Derivado', 'Resolucion Primer Contacto', 'Seguimiento Especial', 'Detalle Seguimiento', 'Fecha Resolucion', 'Fecha Resolucion Real'
  ];

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName("Tickets_Master");

    if (!sheet) {
      sheet = ss.insertSheet("Tickets_Master");
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
      Logger.log('Hoja "Tickets_Master" creada con encabezados.');
      return sheet;
    }

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastRow === 0 || lastCol === 0) {
      sheet.clear();
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight("bold");
      Logger.log('Hoja "Tickets_Master" recreada con estructura completa.');
      return sheet;
    }

    try {
      const existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      if (!existingHeaders.includes('Fecha Resolucion Real')) {
        Logger.log('Falta la columna "Fecha Resolucion Real", ejecutando repairSheetColumns()');
        repairSheetColumns();
        sheet = ss.getSheetByName("Tickets_Master");
      }
    } catch (headerError) {
      Logger.log(`Error verificando encabezados: ${headerError.message}`);
      repairSheetColumns();
      sheet = ss.getSheetByName("Tickets_Master");
    }
    return sheet;
  } catch (e) {
    Logger.log(`Error crítico al acceder o crear la hoja: ${e.message}`);
    throw new Error(`No se pudo acceder a la Hoja de Cálculo. Verifica el ID y los permisos.`);
  }
}

/**
 * 🔧 FUNCIÓN DE REPARACIÓN ROBUSTA para la hoja "Tickets_Master".
 */
function repairSheetColumns() {
  try {
    Logger.log('🔧 Iniciando reparación robusta de columnas...');
    if (!SPREADSHEET_ID || SPREADSHEET_ID.trim() === '') throw new Error('SPREADSHEET_ID no está configurado');
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    Logger.log(`✅ Acceso a hoja de cálculo: "${ss.getName()}"`);
    let sheet = ss.getSheetByName("Tickets_Master");
    if (!sheet) {
      Logger.log('📝 Creando nueva hoja "Tickets_Master"...');
      sheet = ss.insertSheet("Tickets_Master");
    }
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    Logger.log(`📊 Estado actual - Filas: ${lastRow}, Columnas: ${lastCol}`);
    const HEADERS = [
      'Timestamp', 'Creado Por', 'Usuario Cliente', 'N° Caso Yoizen',
      'Planilla (Cliente)', 'URL Adjunto', 'Herramienta', 'Pedido de escalamiento (Asunto)',
      'ID Derivacion FAN', 'N° Ticket Escalamiento', 'Ticket Interno', 'Estado', 'Observaciones Auditor', 'Prioridad', 'Tomado Por',
      'Mal Derivado', 'Resolucion Primer Contacto', 'Seguimiento Especial', 'Detalle Seguimiento', 'Fecha Resolucion', 'Fecha Resolucion Real'
    ];
    if (lastRow === 0 || lastCol === 0) {
      Logger.log('📝 Hoja vacía o sin columnas, creando estructura completa...');
      sheet.clear();
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight("bold");
      Logger.log('✅ Estructura completa creada');
      return { status: "success", message: "Estructura creada desde cero" };
    }
    let existingHeaders = [];
    if (lastCol > 0) existingHeaders = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    Logger.log(`📋 Encabezados existentes: ${existingHeaders.join(', ')}`);
    const missingColumns = HEADERS.filter(h => !existingHeaders.includes(h));

    if (missingColumns.length > 0) {
      Logger.log(`➕ Añadiendo columnas faltantes: ${missingColumns.join(', ')}`);
      let currentCol = lastCol;
      missingColumns.forEach(header => {
        currentCol++;
        sheet.getRange(1, currentCol).setValue(header).setFontWeight("bold");
        if (lastRow > 1) {
          const defaultValue = header === 'Prioridad' ? 'Normal' : '';
          const range = sheet.getRange(2, currentCol, lastRow - 1, 1);
          const values = Array(lastRow - 1).fill([defaultValue]);
          range.setValues(values);
        }
        Logger.log(`✅ Columna "${header}" añadida`);
      });
    } else {
      Logger.log('✅ Todas las columnas necesarias ya existen');
    }
    Logger.log('✅ Reparación completada exitosamente');
    return { status: "success", message: "Columnas reparadas correctamente" };
  } catch (error) {
    Logger.log(`❌ Error en reparación robusta: ${error.message}`);
    return { status: "error", message: error.message };
  }
}

/**
 * Función central y robusta para obtener la hoja de cálculo "Ticket_Comments".
 */

function getOrCreateCommentsSheet() {
  const HEADERS = ['Ticket Interno', 'Timestamp', 'Autor', 'Comentario'];
  const SHEET_NAME = "Ticket_Comments";
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(HEADERS);
      const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
      headerRange.setFontWeight("bold").setBackground("#e8f4fd").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);
      sheet.setColumnWidth(1, 120);
      sheet.setColumnWidth(2, 150);
      sheet.setColumnWidth(3, 120);
      sheet.setColumnWidth(4, 400);
      sheet.setFrozenRows(1);
      sheet.getRange(2, 4, sheet.getMaxRows() - 1, 1).setWrap(true).setVerticalAlignment("top");
      Logger.log(`✅ Hoja ${SHEET_NAME} creada exitosamente`);
    } else {
      if (sheet.getLastRow() === 0) {
        sheet.clear();
        sheet.appendRow(HEADERS);
        // Re-apply formatting if needed
      }
      validateCommentsSheet(sheet); // Auto-validation
    }
    return sheet;
  } catch (error) {
    Logger.log(`❌ Error crítico en getOrCreateCommentsSheet: ${error.message}`);
    throw new Error(`No se pudo acceder a la Hoja de Cálculo de comentarios: ${error.message}`);
  }
}


/**
 * Valida la estructura de la hoja de comentarios.
 */
function validateCommentsSheet(sheet) {
  try {
    const EXPECTED_HEADERS = ['Ticket Interno', 'Timestamp', 'Autor', 'Comentario'];
    if (!sheet) return { valid: false, error: 'Hoja no proporcionada' };
    if (sheet.getLastRow() < 1) return { valid: false, error: 'La hoja no tiene encabezados' };
    if (sheet.getLastColumn() < EXPECTED_HEADERS.length) return { valid: false, error: `Faltan columnas` };
    const headers = sheet.getRange(1, 1, 1, EXPECTED_HEADERS.length).getValues()[0];
    for (let i = 0; i < EXPECTED_HEADERS.length; i++) {
      if (headers[i] !== EXPECTED_HEADERS[i]) {
        return { valid: false, error: `Encabezado incorrecto en columna ${i + 1}` };
      }
    }
    return { valid: true };
  } catch (error) {
    return { valid: false, error: `Error al validar: ${error.message}` };
  }
}

/**
 * Función central y robusta para obtener la hoja de cálculo "Ticket_History".
 */

function getOrCreateHistorySheet() {
  const HEADERS = ['Ticket Interno', 'Timestamp', 'Cambio', 'Usuario'];
  const SHEET_NAME = "Ticket_History";
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(HEADERS);
      const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
      headerRange.setFontWeight("bold").setBackground("#f0f0f0").setHorizontalAlignment("center");
      for (let i = 1; i <= HEADERS.length; i++) sheet.autoResizeColumn(i);
      sheet.setFrozenRows(1);
    } else {
      validateHistorySheet(sheet); // Auto-validation
    }
    return sheet;
  } catch (error) {
    console.error(`Error en getOrCreateHistorySheet: ${error.message}`);
    throw error;
  }
}

/**
 * Valida la estructura de la hoja de historial.
 */
function validateHistorySheet(sheet) {
  try {
    const EXPECTED_HEADERS = ['Ticket Interno', 'Timestamp', 'Cambio', 'Usuario'];
    if (!sheet) return { valid: false, error: 'Hoja no proporcionada' };
    if (sheet.getLastRow() < 1) return { valid: false, error: 'La hoja no tiene encabezados' };
    const headers = sheet.getRange(1, 1, 1, EXPECTED_HEADERS.length).getValues()[0];
    for (let i = 0; i < EXPECTED_HEADERS.length; i++) {
      if (headers[i] !== EXPECTED_HEADERS[i]) {
        return { valid: false, error: `Encabezado incorrecto en columna ${i + 1}` };
      }
    }
    return { valid: true };
  } catch (error) {
    return { valid: false, error: `Error al validar: ${error.message}` };
  }
}


/**
 * Función central y robusta para obtener la hoja de cálculo "Planillas_Master".
 */

function getOrCreatePlanillasMasterSheet() {
  const HEADERS = ['Motivo operativo', 'Planilla operativa', 'Ruta operativa'];
  const SHEET_NAME = "Planillas_Master";
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName(SHEET_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME);
      sheet.appendRow(HEADERS);
      const headerRange = sheet.getRange(1, 1, 1, HEADERS.length);
      headerRange.setFontWeight("bold").setBackground("#d9ead3").setHorizontalAlignment("center").setBorder(true, true, true, true, true, true);
      sheet.setColumnWidth(1, 300);
      sheet.setColumnWidth(2, 400);
      sheet.setFrozenRows(1);
    }
    return sheet;
  } catch (error) {
    Logger.log(`❌ Error crítico en getOrCreatePlanillasMasterSheet: ${error.message}`);
    throw new Error(`No se pudo acceder a la Hoja de Cálculo de Planillas_Master.`);
  }
}



/**
 * 🔧 FUNCIÓN DE REPARACIÓN COMPLETA - Repara todas las hojas del sistema
 */
function repairAllSheetsLogic() {
  try {
    Logger.log('🔧 === REPARACIÓN COMPLETA DEL SISTEMA ===');
    const mainResult = repairSheetColumns();
    Logger.log(`1️⃣ Resultado hoja principal: ${mainResult.message}`);
    getOrCreateCommentsSheet();
    Logger.log('2️⃣ Hoja de comentarios verificada');
    getOrCreateHistorySheet();
    Logger.log('3️⃣ Hoja de historial verificada');
    repairNotificationsStructure();
    Logger.log('4️⃣ Hoja de notificaciones verificada y reparada');
    getOrCreatePlanillasMasterSheet();
    Logger.log('5️⃣ Hoja de Planillas_Master verificada y reparada');
    const testNotif = getUnreadNotificationsLogic();
    Logger.log(`6️⃣ Resultado test notificaciones: ${testNotif.status}`);
    Logger.log('🎉 === REPARACIÓN COMPLETA FINALIZADA ===');
    return { status: "success", message: "Todas las hojas reparadas correctamente" };
  } catch (error) {
    Logger.log(`❌ Error en reparación completa: ${error.message}`);
    return { status: "error", message: error.message };
  }
}
