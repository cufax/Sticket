// ===============================================================
// PUNTO DE ENTRADA WEB Y API PÚBLICA PARA EL FRONTEND
// ===============================================================

/**
 * Punto de entrada principal para la aplicación web.
 */
function doGet(e) {
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Sticket')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/**
 * Función de ayuda para incluir otros archivos (CSS, JS) dentro del HTML principal.
 */
function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}


// --- API LLAMADA DESDE EL FRONTEND (google.script.run) ---

function getUserProfile() {
  return getUserProfileLogic();
}

function getDashboardData(options = {}) {
  return getDashboardDataLogic(options);
}

function createTicket(submissionData) {
    return createTicketLogic(submissionData);
}

function submitGbRequest(requestData) {
    return submitGbRequestLogic(requestData);
}

function getGbRequestsHistory() {
    return getGbRequestsHistoryLogic();
}

function updateGbRequestAuditStatus(params) {
    return updateGbRequestAuditStatusLogic(params);
}

function updateTicketStatus(params) {
  return updateTicketStatusLogic(params);
}

function takeTicket(params) {
  return takeTicketLogic(params);
}

function exportTicketsToExcel(exportParams = {}) {
  return exportTicketsToExcelLogic(exportParams);
}

function addTicketComment(commentData) {
  return addTicketCommentLogic(commentData);
}

function getTicketComments(ticketId) {
  return getTicketCommentsLogic(ticketId);
}

function getTicketHistory(ticketId) {
  return getTicketHistoryLogic(ticketId);
}

function getUnreadNotifications() {
  return getUnreadNotificationsLogic();
}

function markNotificationAsRead(row) {
  return markNotificationAsReadLogic(row);
}

function getBotResponse(userQuery) {
  return getBotResponseLogic(userQuery);
}

function getAllMotivosOperativos() {
  return getAllMotivosOperativosLogic();
}

function getPlanillaOperativaByMotivo(motivoOperativo) {
  return getPlanillaOperativaByMotivoLogic(motivoOperativo);
}

function getRutaOperativaByMotivo(motivoOperativo) {
  return getRutaOperativaByMotivoLogic(motivoOperativo);
}
function generateCustomReport(Config) {
  return generateCustomReportLogic(Config);
}
// --- NUEVA FUNCIÓN EXPUESTA ---
function updateTicketAuditMetrics(params) {
  return updateTicketAuditMetricsLogic(params);
}
function updateRealResolutionDate(params) {
  return updateRealResolutionDateLogic(params);
}

// ===============================================================
// FUNCIONES DE CACHÉ EXPUESTAS AL FRONTEND
// ===============================================================

function getCacheStatistics() {
  return getCacheStats();
}

function clearCache() {
  invalidateCache();
  return { status: "success", message: "Caché limpiado correctamente" };
}

// ✅ NUEVA FUNCIÓN OPTIMIZADA (COEXISTE CON LA ORIGINAL)
function getDashboardDataOptimized(options = {}) {
  return getDashboardDataOptimizedLogic(options);
}


// --- FUNCIONES DE ADMINISTRACIÓN (Ejecutar desde el editor) ---

function repairAllSheets() {
  return repairAllSheetsLogic();
}

// -----------Lógica de logueo equipo------------------------------

/**
 * Obtiene el nombre del superior del usuario actual desde la hoja 'Ingreso_Master'.
 */
function getUserSuperior() {
  try {
    const userEmail = Session.getActiveUser().getEmail();
    Logger.log('Usuario actual: ' + userEmail); // Debug: Imprime el email del usuario

    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheetName = "Ingreso_Master";
    let sheet = ss.getSheetByName(sheetName);

    if (!sheet) {
      Logger.log(`ERROR: La hoja "${sheetName}" no existe.`); // Debug: Hoja no encontrada
      throw new Error(`La hoja "${sheetName}" no existe en la hoja de cálculo.`);
    }

    const data = sheet.getDataRange().getValues();
    Logger.log('Datos de la hoja (primera fila): ' + JSON.stringify(data[0])); // Debug: Imprime los encabezados

    if (data.length <= 1) {
      Logger.log('ERROR: La hoja Ingreso_Master está vacía o solo tiene encabezados.'); // Debug: Hoja vacía
      return { status: "error", message: "La hoja 'Ingreso_Master' está vacía o solo tiene encabezados." };
    }

    const headers = data[0];
    const emailColIdx = headers.indexOf('Correo corpo');
    const superiorColIdx = headers.indexOf('Superior');

    Logger.log('Índice de columna Email: ' + emailColIdx); // Debug: Índice de Email
    Logger.log('Índice de columna Superior: ' + superiorColIdx); // Debug: Índice de Superior

    if (emailColIdx === -1 || superiorColIdx === -1) {
      Logger.log("ERROR: Las columnas 'Email' o 'Superior' no se encontraron."); // Debug: Columnas no encontradas
      throw new Error("Las columnas 'Email' o 'Superior' no se encontraron en la hoja 'Ingreso_Master'.");
    }

    for (let i = 1; i < data.length; i++) {
      if (data[i][emailColIdx] === userEmail) {
        Logger.log('Superior encontrado: ' + data[i][superiorColIdx]); // Debug: Superior encontrado
        return { status: "success", superior: data[i][superiorColIdx] };
      }
    }

    Logger.log('INFO: Usuario no encontrado en la hoja Ingreso_Master.'); // Debug: Usuario no encontrado
    return { status: "info", message: "Usuario no encontrado en la hoja 'Ingreso_Master'." };

  } catch (e) {
    Logger.log(`ERROR en getUserSuperior (catch): ${e.message}`); // Debug: Error en el catch
    return { status: "error", message: `Error al obtener el superior del usuario: ${e.message}` };
  }
}

// ----------- termina Lógica de logueo equipo------------------------------

// -----------Lógica de filtro equipo------------------------------

/**
 * Obtiene todos los superiores únicos de la hoja 'Ingreso_Master'.
 */
function getAllSuperiores() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName("Ingreso_Master");

    if (!sheet) {
      throw new Error("La hoja 'Ingreso_Master' no existe.");
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return { status: "error", message: "La hoja 'Ingreso_Master' está vacía." };
    }

    const headers = data[0];
    const superiorColIdx = headers.indexOf('Superior');

    if (superiorColIdx === -1) {
      throw new Error("La columna 'Superior' no se encontró en la hoja 'Ingreso_Master'.");
    }

    // Obtener todos los superiores únicos (excluyendo valores vacíos)
    const superiores = [...new Set(
      data.slice(1)
        .map(row => (row[superiorColIdx] || '').toString().trim())
        .filter(superior => superior !== '')
    )].sort();

    return { status: "success", superiores: superiores };

  } catch (e) {
    Logger.log(`Error en getAllSuperiores: ${e.message}`);
    return { status: "error", message: `Error al obtener superiores: ${e.message}` };
  }
}

/**
 * Obtiene todos los correos de un equipo específico (superior).
 */
function getEmailsByEquipo(superiorName) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName("Ingreso_Master");

    if (!sheet) {
      throw new Error("La hoja 'Ingreso_Master' no existe.");
    }

    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) {
      return [];
    }

    const headers = data[0];
    const emailColIdx = headers.indexOf('Correo corpo');
    const superiorColIdx = headers.indexOf('Superior');

    if (emailColIdx === -1 || superiorColIdx === -1) {
      throw new Error("Las columnas 'Correo corpo' o 'Superior' no se encontraron.");
    }

    // Filtrar todos los correos que pertenecen al superior especificado
    const emails = data.slice(1)
      .filter(row => (row[superiorColIdx] || '').toString().trim() === superiorName)
      .map(row => (row[emailColIdx] || '').toString().trim())
      .filter(email => email !== '');

    return emails;

  } catch (e) {
    Logger.log(`Error en getEmailsByEquipo: ${e.message}`);
    return [];
  }
}

// ----------- termina Lógica de filtro equipo------------------------------

// ===============================================================
// FUNCIONES DE CACHÉ EXPUESTAS AL FRONTEND
// ===============================================================

/**
 * Obtiene estadísticas del caché (para debugging)
 */
function getCacheStatistics() {
  return getCacheStats();
}

/**
 * Limpia el caché manualmente (para testing)
 */
function clearCache() {
  invalidateCache();
  return { status: "success", message: "Caché limpiado correctamente" };
}
