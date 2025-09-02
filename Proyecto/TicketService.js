// ===============================================================
// SERVICIO DE TICKETS (LÓGICA DE NEGOCIO)
// ===============================================================

/**
 * Obtiene los datos para el dashboard, aplicando filtros y paginación.
 */
/**
 * Obtiene los datos para el dashboard, aplicando filtros y paginación.
 */
function getDashboardDataLogic(options = {}) {
  try {
    const userProfile = getUserProfileLogic();
    const sheet = getOrCreateSheet();

    if (sheet.getLastRow() <= 1) {
      return { tickets: [], allTicketsForStats: [], totalTickets: 0, allPlanillas: [], allCreators: [], allMotivosEscalamiento: [], allSuperiores: [], status: "ok" };
    }

    const rawData = sheet.getDataRange().getValues();
    const headers = rawData.shift();
    const allTickets = sanitizeTicketData(rawData, headers);

    const getUniqueSortedValues = (key) => [...new Set(allTickets.map(t => (t[key] || '').toString().trim()).filter(Boolean))].sort();
    const allPlanillas = getUniqueSortedValues('Planilla (Cliente)');
    const allCreators = getUniqueSortedValues('Creado Por');
    const allMotivosEscalamiento = getUniqueSortedValues('Pedido de escalamiento (Asunto)');
    const superioresResponse = getAllSuperiores();
    const allSuperiores = superioresResponse.status === 'success' ? superioresResponse.superiores : [];

    let filteredTickets = allTickets;

    if (options.searchQuery) {
      const query = options.searchQuery.toLowerCase();
      filteredTickets = filteredTickets.filter(t => Object.values(t).some(val => String(val).toLowerCase().includes(query)));
    }
    if (options.filterStatus) filteredTickets = filteredTickets.filter(t => t['Estado'] === options.filterStatus);
    if (options.filterTool) filteredTickets = filteredTickets.filter(t => t['Herramienta'] === options.filterTool);
    if (options.filterMotivoEscalamiento) filteredTickets = filteredTickets.filter(t => (t['Pedido de escalamiento (Asunto)'] || '') === options.filterMotivoEscalamiento);
    if (options.filterCreatedBy) filteredTickets = filteredTickets.filter(t => t['Creado Por'] === options.filterCreatedBy);
    if (options.filterPriority) filteredTickets = filteredTickets.filter(t => (t['Prioridad'] || 'Normal') === options.filterPriority);
    if (options.startDate) filteredTickets = filteredTickets.filter(t => t['Timestamp'] && new Date(t['Timestamp']) >= new Date(options.startDate));
    if (options.filterEquipo) {
      const emailsDelEquipo = getEmailsByEquipo(options.filterEquipo);
      if (emailsDelEquipo.length > 0) {
        filteredTickets = filteredTickets.filter(t => emailsDelEquipo.includes(t['Creado Por']));
      }
    }
    if (options.endDate) {
      const end = new Date(options.endDate);
      end.setHours(23, 59, 59, 999);
      filteredTickets = filteredTickets.filter(t => t['Timestamp'] && new Date(t['Timestamp']) <= end);
    }
    if (options.filterMyTickets) {
      filteredTickets = filteredTickets.filter(t => t['Creado Por'] === userProfile.email);
    }

    filteredTickets.sort((a, b) => new Date(b['Timestamp']) - new Date(a['Timestamp']));

    const totalFilteredTickets = filteredTickets.length;
    const pageSize = options.pageSize || 10;
    const page = options.page || 1;
    const startIndex = (page - 1) * pageSize;
    const paginatedTickets = filteredTickets.slice(startIndex, startIndex + pageSize);

    return {
      tickets: paginatedTickets,
      allTicketsForStats: filteredTickets,
      totalTickets: totalFilteredTickets,
      allPlanillas,
      allCreators,
      allMotivosEscalamiento,
      allSuperiores,
      status: "ok"
    };
  } catch (error) {
    Logger.log(`ERROR GRAVE al leer datos: ${error.message}. Stack: ${error.stack}`);
    return { status: "error", message: `⚠️ No se pudieron cargar los datos. Detalle: ${error.message}` };
  }
}


/**
 * Crea un nuevo ticket con los datos enviados desde el formulario del frontend.
 */
function createTicketLogic(submissionData) {
  try {
    const sheet = getOrCreateSheet();
    const requiredFields = {
      usuario: "Usuario (U)",
      casoYoizen: "Número de caso Yoizen",
      herramienta: "Herramienta",
      planilla: "Planilla Escalamiento",
      motivoEscalamiento: "Motivo Escalamiento",
      idDerivacion: "ID Derivación FAN"
    };
    for (const key in requiredFields) {
      if (!submissionData[key] || String(submissionData[key]).trim() === '') {
        throw new Error(`El campo "${requiredFields[key]}" es obligatorio.`);
      }
    }

    let imageUrl = "No se adjuntó imagen";
    if (submissionData && submissionData.fileContent) {
      const parts = submissionData.fileContent.split(",");
      const blob = Utilities.newBlob(Utilities.base64Decode(parts[1]), submissionData.fileType, submissionData.fileName);
      const folder = DriveApp.getFolderById(SUPERUSER_DRIVE_FOLDER_ID);
      const file = folder.createFile(blob);
      imageUrl = file.getUrl();
    }

    const timestamp = new Date();
    const currentUser = Session.getActiveUser().getEmail();
    const ticketInterno = `TK-${Math.floor(100000 + Math.random() * 900000)}`;
    const newRow = [
      timestamp, currentUser, submissionData.usuario.trim(), submissionData.casoYoizen.trim(),
      submissionData.planilla.trim(), imageUrl, submissionData.herramienta.trim(),
      submissionData.motivoEscalamiento.trim(), (submissionData.idDerivacion || '').trim(),
      (submissionData.numeroTicket || '').trim(), ticketInterno, 'Registrado', '', 'Normal', '',
      false, false, false, '', '' // Nuevos campos de auditoría
    ];
    sheet.appendRow(newRow);
    logTicketChange(ticketInterno, `Ticket creado por ${currentUser}`);

    const notificationMessage = `Nuevo ticket #${ticketInterno} creado por ${currentUser} - Asunto: ${submissionData.motivoEscalamiento}`;
    addNotificationToSheet('SUPERUSER', notificationMessage, ticketInterno);
    addNotificationToSheet('AUDITOR', notificationMessage, ticketInterno);

    return { status: "success", message: "Escalamiento recibido.", ticket: ticketInterno };
  } catch (error) {
    Logger.log(`ERROR GRAVE al crear ticket: ${error.message}. Stack: ${error.stack}`);
    return { status: "error", message: `Error en el servidor al crear el ticket: ${error.message}` };
  }
}


/**
 * Actualiza el estado, observaciones, prioridad y/o 'Tomado Por' de un ticket.
 * ✅ CORREGIDO: Se elimina la lógica heredada que poblaba automáticamente la columna 'Fecha Resolucion'.
 */
function updateTicketStatusLogic(params) {
  try {
    const { ticketId, newStatus, newObservations, newPriority } = params;
    const userProfile = getUserProfileLogic();
    if (userProfile.role !== 'AUDITOR' && userProfile.role !== 'SUPERUSER') {
      throw new Error('Acceso denegado: No tienes permisos para modificar tickets.');
    }
    const sheet = getOrCreateSheet();
    if (sheet.getLastRow() <= 1) throw new Error("No hay tickets para actualizar.");

    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const ticketIdColIdx = headers.indexOf('Ticket Interno');
    const statusColIdx = headers.indexOf('Estado');
    const obsColIdx = headers.indexOf('Observaciones Auditor');
    const prioridadColIdx = headers.indexOf('Prioridad');
    const creadoPorColIdx = headers.indexOf('Creado Por');

    if (ticketIdColIdx === -1) throw new Error('Columna "Ticket Interno" no encontrada.');

    const rowIndex = data.findIndex(row => row[ticketIdColIdx] === ticketId);
    if (rowIndex === -1) throw new Error(`Ticket con ID ${ticketId} no encontrado.`);

    const row = data[rowIndex];
    const sheetRowIndex = rowIndex + 2;
    let changes = [];
    let ticketCreatorEmail = row[creadoPorColIdx];

    if (newStatus !== undefined && newStatus !== row[statusColIdx]) {
      sheet.getRange(sheetRowIndex, statusColIdx + 1).setValue(newStatus);
      changes.push('estado');
      logTicketChange(ticketId, `Estado cambiado de '${row[statusColIdx]}' a '${newStatus}'`);
      addNotificationToSheet(ticketCreatorEmail, `Tu ticket #${ticketId} ha cambiado de estado: ${row[statusColIdx]} -> ${newStatus}`, ticketId);
    }

    if (newObservations !== undefined && obsColIdx > -1 && newObservations.trim() !== (row[obsColIdx] || '').trim()) {
      sheet.getRange(sheetRowIndex, obsColIdx + 1).setValue(newObservations.trim());
      changes.push('observaciones');
      logTicketChange(ticketId, `Observaciones actualizadas.`);
      addNotificationToSheet(ticketCreatorEmail, `Nueva observación del auditor en tu ticket #${ticketId}`, ticketId);
    }

    if (newPriority !== undefined && prioridadColIdx > -1 && newPriority !== (row[prioridadColIdx] || 'Normal')) {
      sheet.getRange(sheetRowIndex, prioridadColIdx + 1).setValue(newPriority);
      changes.push('prioridad');
      logTicketChange(ticketId, `Prioridad cambiada de '${row[prioridadColIdx] || 'Normal'}' a '${newPriority}'`);
      addNotificationToSheet(ticketCreatorEmail, `La prioridad de tu ticket #${ticketId} ha cambiado a: ${newPriority}`, ticketId);
    }

    if (changes.length > 0) {
      return { status: "success", message: `Ticket actualizado correctamente (${changes.join(', ')}).` };
    }
    return { status: "info", message: "No se realizaron cambios en el ticket." };
  } catch (error) {
    Logger.log(`ERROR al actualizar ticket: ${error.message}. Stack: ${error.stack}`);
    return { status: "error", message: `Error en el servidor al actualizar: ${error.message}` };
  }
}

/**
 * Nueva función para actualizar las métricas de auditoría de un ticket.
 */
function updateTicketAuditMetricsLogic(params) {
  try {
    const { ticketId, malDerivado, primerContacto, seguimientoEspecial, detalleSeguimiento } = params;
    const userProfile = getUserProfileLogic();
    if (userProfile.role !== 'AUDITOR' && userProfile.role !== 'SUPERUSER') {
      throw new Error('Acceso denegado: No tienes permisos para modificar métricas de auditoría.');
    }

    const sheet = getOrCreateSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const ticketIdColIdx = headers.indexOf('Ticket Interno');

    if (ticketIdColIdx === -1) throw new Error('Columna "Ticket Interno" no encontrada.');

    const rowIndex = data.findIndex(row => row[ticketIdColIdx] === ticketId);
    if (rowIndex === -1) throw new Error(`Ticket con ID ${ticketId} no encontrado.`);

    const sheetRowIndex = rowIndex + 2;
    let changes = [];

    const updateColumn = (colName, newValue) => {
      const colIdx = headers.indexOf(colName);
      if (colIdx !== -1 && newValue !== undefined && newValue !== data[rowIndex][colIdx]) {
        sheet.getRange(sheetRowIndex, colIdx + 1).setValue(newValue);
        changes.push(colName);
      }
    };

    updateColumn('Mal Derivado', malDerivado);
    updateColumn('Resolucion Primer Contacto', primerContacto);
    updateColumn('Seguimiento Especial', seguimientoEspecial);
    updateColumn('Detalle Seguimiento', detalleSeguimiento);


    if (changes.length > 0) {
      logTicketChange(ticketId, `Métricas de auditoría actualizadas: ${changes.join(', ')}.`);
      return { status: "success", message: "Métricas de auditoría actualizadas correctamente." };
    }

    return { status: "info", message: "No se realizaron cambios en las métricas." };

  } catch (error) {
    Logger.log(`ERROR al actualizar métricas de auditoría: ${error.message}. Stack: ${error.stack}`);
    return { status: "error", message: `Error al actualizar métricas: ${error.message}` };
  }
}
/**
 * Permite a un auditor/superuser "tomar" un ticket.
 */
function takeTicketLogic(params) {
  try {
    const { ticketId } = params;
    const userProfile = getUserProfileLogic();
    if (userProfile.role !== 'AUDITOR' && userProfile.role !== 'SUPERUSER') {
      throw new Error('Acceso denegado: Solo auditores o superusuarios pueden tomar tickets.');
    }
    const sheet = getOrCreateSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const ticketIdColIdx = headers.indexOf('Ticket Interno');
    const tomadoPorColIdx = headers.indexOf('Tomado Por');
    const creadoPorColIdx = headers.indexOf('Creado Por');

    if (ticketIdColIdx === -1 || tomadoPorColIdx === -1) throw new Error('Columnas críticas no encontradas.');

    const rowIndex = data.findIndex(row => row[ticketIdColIdx] === ticketId);
    if (rowIndex === -1) throw new Error(`Ticket con ID ${ticketId} no encontrado.`);

    const row = data[rowIndex];
    const sheetRowIndex = rowIndex + 2;
    const currentTomadoPor = (row[tomadoPorColIdx] || '').trim();
    const currentUserShortName = userProfile.email.split('@')[0];

    if (currentTomadoPor === currentUserShortName) {
      return { status: "info", message: `Ya has tomado el ticket ${ticketId}.` };
    } else if (currentTomadoPor !== '') {
      return { status: "info", message: `El ticket ${ticketId} ya ha sido tomado por ${currentTomadoPor}.` };
    }

    sheet.getRange(sheetRowIndex, tomadoPorColIdx + 1).setValue(currentUserShortName);
    logTicketChange(ticketId, `Ticket tomado por ${currentUserShortName}`);

    const ticketCreatorEmail = row[creadoPorColIdx];
    if (ticketCreatorEmail && ticketCreatorEmail !== userProfile.email) {
      addNotificationToSheet(ticketCreatorEmail, `Tu ticket #${ticketId} ha sido tomado por ${currentUserShortName}.`, ticketId);
    }

    return { status: "success", message: `Ticket ${ticketId} tomado por ti.` };
  } catch (error) {
    Logger.log(`ERROR al tomar ticket: ${error.message}. Stack: ${error.stack}`);
    return { status: "error", message: `Error en el servidor al tomar el ticket: ${error.message}` };
  }
}

/**
 * Exporta tickets filtrados a un archivo Excel (.xlsx) en Google Drive.
 */
function exportTicketsToExcelLogic(exportParams = {}) {
  try {
    const userProfile = getUserProfileLogic();
    if (userProfile.role !== 'SUPERUSER' && userProfile.role !== 'AUDITOR') {
      return { status: "error", message: "Acceso denegado: No tienes permisos para exportar datos." };
    }
    const sheet = getOrCreateSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return { status: "error", message: "No hay datos de tickets para exportar." };

    const headers = data[0];
    let filteredData = data.slice(1);

    if (exportParams.startDate) {
      const start = new Date(exportParams.startDate);
      filteredData = filteredData.filter(row => new Date(row[headers.indexOf('Timestamp')]) >= start);
    }
    if (exportParams.endDate) {
      const end = new Date(exportParams.endDate);
      end.setHours(23, 59, 59, 999);
      filteredData = filteredData.filter(row => new Date(row[headers.indexOf('Timestamp')]) <= end);
    }

    const selectedColumns = exportParams.columns || headers;
    const selectedHeaderIndices = selectedColumns.map(colName => headers.indexOf(colName)).filter(idx => idx !== -1);

    const exportData = [selectedColumns];
    filteredData.forEach(row => {
      exportData.push(selectedHeaderIndices.map(idx => row[idx]));
    });

    const timestamp = new Date().toISOString().slice(0, 19).replace(/:/g, '-');
    const workbook = SpreadsheetApp.create(`Tickets_Export_${timestamp}`);
    workbook.getSheets()[0].getRange(1, 1, exportData.length, exportData[0].length).setValues(exportData);

    const file = DriveApp.getFileById(workbook.getId());
    const folder = DriveApp.getFolderById(SUPERUSER_DRIVE_FOLDER_ID);
    folder.addFile(file);
    DriveApp.getRootFolder().removeFile(file);
    const fileUrl = file.getUrl();

    Logger.log(`Exportación completada por ${userProfile.email}. Archivo: ${fileUrl}`);
    return { status: "success", fileUrl: fileUrl, message: "Tickets exportados a Excel correctamente." };
  } catch (error) {
    Logger.log(`ERROR al exportar a Excel: ${error.message}`);
    return { status: "error", message: `Error en el servidor al exportar a Excel: ${error.message}` };
  }
}

/**
 * Actualiza la prioridad de los tickets a "Urgente" si hay más de 5 tickets activos para una misma planilla.
 */
function updateTicketPrioritiesByVolume() {
  const sheet = getOrCreateSheet();
  if (sheet.getLastRow() <= 1) return;
  const data = sheet.getDataRange().getValues();
  const headers = data.shift();
  const planillaColIdx = headers.indexOf('Planilla (Cliente)');
  const estadoColIdx = headers.indexOf('Estado');
  const prioridadColIdx = headers.indexOf('Prioridad');
  const ticketInternoColIdx = headers.indexOf('Ticket Interno');
  if ([planillaColIdx, estadoColIdx, prioridadColIdx, ticketInternoColIdx].includes(-1)) return;

  const planillaCounts = data.reduce((acc, row) => {
    const planilla = row[planillaColIdx];
    const estado = row[estadoColIdx];
    if (planilla && ['Registrado', 'En Progreso'].includes(estado)) {
      if (!acc[planilla]) acc[planilla] = [];
      acc[planilla].push({
        rowNum: acc.length + 2, // approximation, better to pass index
        id: row[ticketInternoColIdx],
        currentPriority: row[prioridadColIdx] || 'Normal'
      });
    }
    return acc;
  }, {});

  for (const planilla in planillaCounts) {
    if (planillaCounts[planilla].length > 5) {
      planillaCounts[planilla].forEach((ticketInfo, index) => {
        if (ticketInfo.currentPriority !== 'Urgente') {
          const realRowNum = data.findIndex(r => r[ticketInternoColIdx] === ticketInfo.id) + 2;
          sheet.getRange(realRowNum, prioridadColIdx + 1).setValue('Urgente');
          logTicketChange(ticketInfo.id, `Prioridad cambiada a 'Urgente' automáticamente por volumen.`);
        }
      });
    }
  }
}

/**
 * 🆕 Nueva función para actualizar la fecha de resolución real de un ticket.
 * ✅ VERSIÓN DEFINITIVA: Se corrige el manejo de la fecha para evitar problemas de zona horaria.
 */
function updateRealResolutionDateLogic(params) {
  try {
    const { ticketId, resolutionDate } = params;
    const userProfile = getUserProfileLogic();

    if (userProfile.role !== 'AUDITOR' && userProfile.role !== 'SUPERUSER') {
      throw new Error('Acceso denegado: No tienes permisos para modificar esta fecha.');
    }
    if (!ticketId || !resolutionDate) {
      throw new Error('El ID del ticket y la fecha de resolución son obligatorios.');
    }

    const sheet = getOrCreateSheet();
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const ticketIdColIdx = headers.indexOf('Ticket Interno');
    const realResolutionDateColIdx = headers.indexOf('Fecha Resolucion Real');

    if (ticketIdColIdx === -1 || realResolutionDateColIdx === -1) {
      repairSheetColumns(); // Intenta reparar la hoja si faltan columnas
      throw new Error('Columnas críticas no encontradas. Se intentó una reparación, por favor reintente.');
    }

    const rowIndex = data.findIndex(row => row[ticketIdColIdx] === ticketId);
    if (rowIndex === -1) {
      throw new Error(`Ticket con ID ${ticketId} no encontrado.`);
    }

    const sheetRowIndex = rowIndex + 2;

    // ✅ CORRECCIÓN CLAVE: Parsear la fecha correctamente
    const parts = resolutionDate.split('-');
    if (parts.length !== 3) {
      throw new Error('El formato de la fecha no es válido. Se esperaba YYYY-MM-DD.');
    }
    const year = parseInt(parts[0], 10);
    const month = parseInt(parts[1], 10) - 1; // El mes es 0-indexado en JS
    const day = parseInt(parts[2], 10);

    // Crear la fecha al final del día para asegurar que sea posterior a la creación
    const newDateUTC = new Date(Date.UTC(year, month, day, 23, 59, 59, 999));

    if (isNaN(newDateUTC.getTime())) {
      throw new Error('La fecha proporcionada no es válida.');
    }

    // Guardar la fecha en la hoja
    sheet.getRange(sheetRowIndex, realResolutionDateColIdx + 1).setValue(newDateUTC);
    logTicketChange(ticketId, `Fecha de resolución real actualizada a: ${resolutionDate}`);

    Logger.log(`✅ Fecha de resolución real guardada para ticket ${ticketId}: ${newDateUTC.toISOString()}`);

    return { status: "success", message: "Fecha de resolución real actualizada correctamente." };

  } catch (error) {
    Logger.log(`ERROR al actualizar fecha de resolución real: ${error.message}. Stack: ${error.stack}`);
    return { status: "error", message: `Error al actualizar la fecha: ${error.message}` };
  }
}

// ===============================================================
// VERSIÓN OPTIMIZADA CON CACHÉ
// ===============================================================

function getDashboardDataOptimizedLogic(options = {}) {
  try {
    Logger.log(`🚀 Iniciando getDashboardDataOptimized`);
    const startTime = Date.now();

    // ✅ MEJORA: Intentar obtener datos del caché
    const cacheKey = 'tickets';
    let allTickets = getCachedData(cacheKey);
    let fromCache = true;

    if (!allTickets) {
      fromCache = false;
      Logger.log(`📊 Leyendo datos desde Google Sheets...`);

      const userProfile = getUserProfileLogic();
      const sheet = getOrCreateSheet();

      if (sheet.getLastRow() <= 1) {
        return {
          tickets: [],
          allTicketsForStats: [],
          totalTickets: 0,
          allPlanillas: [],
          allCreators: [],
          allMotivosEscalamiento: [],
          allSuperiores: [],
          status: "ok",
          performance: { loadTime: Date.now() - startTime, fromCache: false }
        };
      }

      const rawData = sheet.getDataRange().getValues();
      const headers = rawData.shift();
      allTickets = sanitizeTicketData(rawData, headers);
      setCachedData(cacheKey, allTickets);
      Logger.log(`💾 ${allTickets.length} tickets guardados en caché`);
    }

    // Resto de la lógica igual que la función original
    const getUniqueSortedValues = (key) => [...new Set(allTickets.map(t => (t[key] || '').toString().trim()).filter(Boolean))].sort();
    const allPlanillas = getUniqueSortedValues('Planilla (Cliente)');
    const allCreators = getUniqueSortedValues('Creado Por');
    const allMotivosEscalamiento = getUniqueSortedValues('Pedido de escalamiento (Asunto)');
    const superioresResponse = getAllSuperiores();
    const allSuperiores = superioresResponse.status === 'success' ? superioresResponse.superiores : [];

    let filteredTickets = allTickets;

    // Aplicar todos los filtros (copiado de la función original)
    if (options.searchQuery) {
      const query = options.searchQuery.toLowerCase();
      filteredTickets = filteredTickets.filter(t => Object.values(t).some(val => String(val).toLowerCase().includes(query)));
    }
    if (options.filterStatus) filteredTickets = filteredTickets.filter(t => t['Estado'] === options.filterStatus);
    if (options.filterTool) filteredTickets = filteredTickets.filter(t => t['Herramienta'] === options.filterTool);
    if (options.filterMotivoEscalamiento) filteredTickets = filteredTickets.filter(t => (t['Pedido de escalamiento (Asunto)'] || '') === options.filterMotivoEscalamiento);
    if (options.filterCreatedBy) filteredTickets = filteredTickets.filter(t => t['Creado Por'] === options.filterCreatedBy);
    if (options.filterPriority) filteredTickets = filteredTickets.filter(t => (t['Prioridad'] || 'Normal') === options.filterPriority);
    if (options.startDate) filteredTickets = filteredTickets.filter(t => t['Timestamp'] && new Date(t['Timestamp']) >= new Date(options.startDate));
    if (options.filterEquipo) {
      const emailsDelEquipo = getEmailsByEquipo(options.filterEquipo);
      if (emailsDelEquipos.length > 0) {
        filteredTickets = filteredTickets.filter(t => emailsDelEquipo.includes(t['Creado Por']));
      }
    }
    if (options.endDate) {
      const end = new Date(options.endDate);
      end.setHours(23, 59, 59, 999);
      filteredTickets = filteredTickets.filter(t => t['Timestamp'] && new Date(t['Timestamp']) <= end);
    }
    if (options.filterMyTickets) {
      const userProfile = getUserProfileLogic();
      filteredTickets = filteredTickets.filter(t => t['Creado Por'] === userProfile.email);
    }

    filteredTickets.sort((a, b) => new Date(b['Timestamp']) - new Date(a['Timestamp']));

    const totalFilteredTickets = filteredTickets.length;
    const pageSize = options.pageSize || 10;
    const page = options.page || 1;
    const startIndex = (page - 1) * pageSize;
    const paginatedTickets = filteredTickets.slice(startIndex, startIndex + pageSize);

    const loadTime = Date.now() - startTime;
    Logger.log(`⚡ getDashboardDataOptimized completado en ${loadTime}ms (caché: ${fromCache})`);

    return {
      tickets: paginatedTickets,
      allTicketsForStats: filteredTickets,
      totalTickets: totalFilteredTickets,
      allPlanillas,
      allCreators,
      allMotivosEscalamiento,
      allSuperiores,
      status: "ok",
      performance: { loadTime, fromCache }
    };
  } catch (error) {
    Logger.log(`ERROR en getDashboardDataOptimized: ${error.message}`);
    return { status: "error", message: `Error: ${error.message}` };
  }
}

function invalidateTicketsCache() {
  invalidateCache('tickets');
  invalidateCache('filters');
  Logger.log(`🗑️ Caché invalidado por modificación de datos`);
}

/**
 * ✅ FUNCIÓN CORREGIDA Y ROBUSTA: Genera datos para el reporte personalizado.
 * La lógica de fechas ha sido reescrita para ser inmune a problemas de zona horaria.
 */
function generateCustomReportLogic(config) {
  try {
    if (!config || typeof config !== 'object') {
      throw new Error("La configuración del reporte no fue recibida por el servidor.");
    }

    const sheet = getOrCreateSheet();
    if (sheet.getLastRow() <= 1) {
      return { status: 'success', data: { labels: [], data: [] } };
    }

    const rawData = sheet.getDataRange().getValues();
    const headers = rawData.shift();
    let allTickets = sanitizeTicketData(rawData, headers);

    // ----> INICIO DE LA CORRECCIÓN DE FECHAS <----
    if (config.startDate) {
      const startParts = config.startDate.split('-').map(Number);
      // Se crea la fecha en UTC para evitar ambigüedades de zona horaria.
      const startDate = new Date(Date.UTC(startParts[0], startParts[1] - 1, startParts[2], 0, 0, 0, 0));
      allTickets = allTickets.filter(t => {
        const ticketDate = new Date(t['Timestamp']);
        return ticketDate >= startDate;
      });
    }
    if (config.endDate) {
      const endParts = config.endDate.split('-').map(Number);
      // Se crea la fecha de fin al ÚLTIMO milisegundo del día en UTC.
      const endDate = new Date(Date.UTC(endParts[0], endParts[1] - 1, endParts[2], 23, 59, 59, 999));
      allTickets = allTickets.filter(t => {
        const ticketDate = new Date(t['Timestamp']);
        return ticketDate <= endDate;
      });
    }
    // ----> FIN DE LA CORRECCIÓN DE FECHAS <----

    // Agrupar y contar
    const groupedData = allTickets.reduce((acc, ticket) => {
      const dimensionValue = ticket[config.dimension] || 'No especificado';
      if (config.metric === 'Ticket Interno') {
        acc[dimensionValue] = (acc[dimensionValue] || 0) + 1;
      }
      return acc;
    }, {});

    const sortedEntries = Object.entries(groupedData).sort(([, a], [, b]) => b - a);

    const labels = sortedEntries.map(entry => entry[0]);
    const data = sortedEntries.map(entry => entry[1]);

    return { status: 'success', data: { labels, data } };

  } catch (error) {
    Logger.log(`ERROR al generar reporte personalizado: ${error.message}`);
    return { status: "error", message: `Error al generar el reporte: ${error.message}` };
  }
}
