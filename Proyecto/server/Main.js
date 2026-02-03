// ==============================================================================================================================================
// PUNTO DE ENTRADA WEB Y API PÚBLICA PARA EL FRONTEND - PROPIETARIOS DEL CÓDIGO: FACUNDO ALVAREZ  Y ALEJANDRO DAVID GARCIA PARDO
// ==============================================================================================================================================

/**
 * Punto de entrada principal para la aplicación web.
 */
function doGet(e) {
    return HtmlService.createTemplateFromFile('index')
        .evaluate()
        .setTitle('Sticket')
        .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

// =============================================================================================================================================================================================

/**
 * Función de ayuda para incluir otros archivos (CSS, JS) dentro del HTML principal.
 */
function include(filename) {
    return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

// =============================================================================================================================================================================================
// --- API LLAMADA DESDE EL FRONTEND (google.script.run) ---

function getUserProfile() {
    return getUserProfileLogic();
}
// =============================================================================================================================================================================================

function getDashboardData(options = {}) {
    return getDashboardDataLogic(options);
}
// =============================================================================================================================================================================================

function createTicket(submissionData) {
    return createTicketLogic(submissionData);
}
// =============================================================================================================================================================================================

function updateTicketDetails(params) {
    return updateTicketDetailsLogic(params);
}
// =============================================================================================================================================================================================

function takeTicket(params) {
    return takeTicketLogic(params);
}

function exportTicketsToExcel(exportParams = {}) {
    return exportTicketsToExcelLogic(exportParams);
}
// =============================================================================================================================================================================================

function addTicketComment(commentData) {
    return addTicketCommentLogic(commentData);
}
// =============================================================================================================================================================================================

function getTicketComments(ticketId) {
    return getTicketCommentsLogic(ticketId);
}
// =============================================================================================================================================================================================

function getTicketHistory(ticketId) {
    return getTicketHistoryLogic(ticketId);
}
// =============================================================================================================================================================================================

function getUnreadNotifications() {
    return getUnreadNotificationsLogic();
}
// =============================================================================================================================================================================================

function markNotificationAsRead(row) {
    return markNotificationAsReadLogic(row);
}
// =============================================================================================================================================================================================

function getBotResponse(userQuery) {
    return getBotResponseLogic(userQuery);
}
// =============================================================================================================================================================================================

function getAllMotivosOperativos() {
    return getAllMotivosOperativosLogic();
}
// =============================================================================================================================================================================================

function getPlanillaOperativaByMotivo(motivoOperativo) {
    return getPlanillaOperativaByMotivoLogic(motivoOperativo);
}
// =============================================================================================================================================================================================

function getRutaOperativaByMotivo(motivoOperativo) {
    return getRutaOperativaByMotivoLogic(motivoOperativo);
}
// =============================================================================================================================================================================================

// --- FUNCIONES DE ADMINISTRACIÓN (Ejecutar desde el editor) ---

function repairAllSheets() {
    return repairAllSheetsLogic();
}

// =============================================================================================================================================================================================
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
// =============================================================================================================================================================================================
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

// =============================================================================================================================================================================================

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
// =============================================================================================================================================================================================

/**
 * 🧪 FUNCIÓN DE TEST PARA BACKEND - Ejecutar desde el editor de Apps Script
 */
function testBackendAuditFlow() {
    try {
        Logger.log('🧪 INICIANDO TEST DE BACKEND');

        // 1. Verificar estructura de hoja
        const repairResult = repairSheetColumns();
        Logger.log('1️⃣ Reparación de estructura: ' + JSON.stringify(repairResult));

        // 2. Obtener un ticket de prueba
        const sheet = getOrCreateSheet();
        const data = sheet.getDataRange().getValues();

        if (data.length <= 1) {
            Logger.log('❌ No hay tickets para probar');
            return;
        }

        const headers = data[0];
        const firstTicket = data[1];
        const ticketId = firstTicket[headers.indexOf('Ticket Interno')];

        Logger.log('2️⃣ Ticket de prueba: ' + ticketId);

        // 3. Mostrar valores actuales
        const currentValues = {
            malEscalado: firstTicket[headers.indexOf('Mal Escalado')],
            seguimientoEspecial: firstTicket[headers.indexOf('Seguimiento Especial')],
            primerContacto: firstTicket[headers.indexOf('Resolucion Primer Contacto')],
            observaciones: firstTicket[headers.indexOf('Observaciones Auditor')]
        };

        Logger.log('3️⃣ Valores actuales: ' + JSON.stringify(currentValues));

        // 4. Simular actualización con cambios reales
        const testUpdates = {
            'Mal Escalado': !Boolean(currentValues.malEscalado), // Cambiar el valor actual
            'Seguimiento Especial': !Boolean(currentValues.seguimientoEspecial), // Cambiar el valor actual
            'Resolucion Primer Contacto': !Boolean(currentValues.primerContacto), // Cambiar el valor actual
            'Observaciones Auditor': 'Test Backend - ' + new Date().toLocaleString()
        };

        Logger.log('4️⃣ Updates a aplicar: ' + JSON.stringify(testUpdates));

        const updateResult = updateTicketDetailsLogic({
            ticketId: ticketId,
            updates: testUpdates
        });

        Logger.log('5️⃣ Resultado de actualización: ' + JSON.stringify(updateResult));

        // 5. Verificar que se guardó correctamente
        const updatedData = sheet.getDataRange().getValues();
        const updatedTicket = updatedData.find(row => row[headers.indexOf('Ticket Interno')] === ticketId);

        const finalValues = {
            malEscalado: updatedTicket[headers.indexOf('Mal Escalado')],
            seguimientoEspecial: updatedTicket[headers.indexOf('Seguimiento Especial')],
            primerContacto: updatedTicket[headers.indexOf('Resolucion Primer Contacto')],
            observaciones: updatedTicket[headers.indexOf('Observaciones Auditor')]
        };

        Logger.log('6️⃣ Valores finales verificados: ' + JSON.stringify(finalValues));
        Logger.log('✅ TEST DE BACKEND COMPLETADO');

        return {
            status: 'success',
            ticketId: ticketId,
            updateResult: updateResult,
            finalValues: finalValues
        };

    } catch (error) {
        Logger.log('❌ ERROR EN TEST DE BACKEND: ' + error.message);
        Logger.log('❌ Stack: ' + error.stack);
        return {
            status: 'error',
            message: error.message
        };
    }
}

// =============================================================================================================================================================================================

/**
 * 🔧 FUNCIÓN PARA LIMPIAR LOGS ANTIGUOS
 */
function clearLogs() {
    console.clear();
    Logger.log('🧹 Logs limpiados - ' + new Date().toLocaleString());
}
// =============================================================================================================================================================================================
/**
 * 🔧 FUNCIÓN PRINCIPAL: Obtener datos de tickets con filtros y paginación
 * @param {Object} params - Parámetros de filtrado y paginación
 * @returns {Object} Datos de tickets y metadatos
 */
function getTicketsData(params = {}) {
    try {
        Logger.log('📊 getTicketsData llamada con parámetros: ' + JSON.stringify(params));

        const {
            page = 1,
            pageSize = 50,
            filters = {}
        } = params;

        // Obtener todos los tickets
        const allTickets = getTicketsLogic();

        if (allTickets.status !== 'ok') {
            Logger.log('❌ Error obteniendo tickets: ' + allTickets.message);
            return allTickets;
        }

        // Aplicar filtros si existen
        let filteredTickets = allTickets.tickets;

        if (filters.planilla && filters.planilla !== '') {
            filteredTickets = filteredTickets.filter(ticket =>
                ticket['Planilla (Cliente)'] &&
                ticket['Planilla (Cliente)'].toLowerCase().includes(filters.planilla.toLowerCase())
            );
        }

        if (filters.estado && filters.estado !== '') {
            filteredTickets = filteredTickets.filter(ticket =>
                ticket['Estado'] === filters.estado
            );
        }

        if (filters.creadoPor && filters.creadoPor !== '') {
            filteredTickets = filteredTickets.filter(ticket =>
                ticket['Creado Por'] &&
                ticket['Creado Por'].toLowerCase().includes(filters.creadoPor.toLowerCase())
            );
        }

        // Calcular paginación
        const totalTickets = filteredTickets.length;
        const startIndex = (page - 1) * pageSize;
        const endIndex = startIndex + pageSize;
        const paginatedTickets = filteredTickets.slice(startIndex, endIndex);

        // Obtener listas para filtros
        const allPlanillas = [...new Set(allTickets.tickets
            .map(t => t['Planilla (Cliente)'])
            .filter(Boolean)
        )].sort();

        const allCreators = [...new Set(allTickets.tickets
            .map(t => t['Creado Por'])
            .filter(Boolean)
        )].sort();

        const result = {
            status: 'ok',
            tickets: paginatedTickets,
            allTicketsForStats: allTickets.tickets, // Todos los tickets para estadísticas
            totalTickets: totalTickets,
            currentPage: page,
            pageSize: pageSize,
            totalPages: Math.ceil(totalTickets / pageSize),
            allPlanillas: allPlanillas,
            allCreators: allCreators,
            allSuperiores: [] // Si tienes esta lista, añádela aquí
        };

        Logger.log('✅ getTicketsData completada exitosamente');
        return result;

    } catch (error) {
        Logger.log('❌ Error en getTicketsData: ' + error.message);
        Logger.log('❌ Stack: ' + error.stack);
        return {
            status: 'error',
            message: 'Error del servidor: ' + error.message,
            tickets: [],
            allTicketsForStats: [],
            totalTickets: 0
        };
    }
}
// =============================================================================================================================================================================================

