// ==============================================================================================================================
// SERVICIO DE TICKETS (LÓGICA DE NEGOCIO) - PROPIETARIOS DEL CÓDIGO: FACUNDO ALVAREZ  Y ALEJANDRO DAVID GARCIA PARDO
// ==============================================================================================================================

/**
 * Obtiene los datos para el dashboard con CACHÉ INTELIGENTE.
 */
function getDashboardDataLogic(options = {}) {
    try {
        const userProfile = getUserProfileLogic();

        // 1. DEFINIMOS LA CLAVE ÚNICA DE CACHÉ
        const CACHE_KEY = "TICKETS_MASTER_RAW_DATA";

        // 2. OBTENER DATOS (DESDE RAM O HOJA)
        // CacheManager debe estar definido en CacheManager.js
        const rawDataPackage = CacheManager.getOrSet(CACHE_KEY, () => {
            const sheet = getOrCreateSheet();
            if (sheet.getLastRow() <= 1) return { empty: true };

            // Leemos TODA la hoja de una vez
            const values = sheet.getDataRange().getValues();
            return { empty: false, values: values };
        });

        // 3. VALIDACIÓN Y EXTRACCIÓN
        if (!rawDataPackage || rawDataPackage.empty) {
            return { tickets: [], allTicketsForStats: [], totalTickets: 0, allPlanillas: [], allCreators: [], allSuperiores: [], pendingTeamTickets: 0, status: "ok" };
        }

        // ⚠️ AQUÍ ESTABA EL ERROR: Debemos definir rawData antes de usarla
        const rawData = rawDataPackage.values;

        // Importante: Trabajamos con una copia para no mutar el objeto original del caché si fuera por referencia
        // Aunque al venir de JSON.parse es una copia nueva, es buena práctica.
        const workingData = rawData.slice();

        const headers = workingData.shift().map(header => header.toString().trim());

        const tickets = workingData.map(row => {
            const ticket = {};
            headers.forEach((header, index) => {
                let value = row[index];
                if (header === 'Timestamp' && value) {
                    // Asegurar que sea fecha válida string para JSON
                    value = new Date(value).toISOString();
                }
                ticket[header] = (value === undefined || value === null) ? '' : value;
            });
            return ticket;
        });

        // --- LOGICA DE RADAR DE EQUIPO ---
        let pendingTeamCount = 0;
        if (['AUDITOR', 'SUPERUSER'].includes(userProfile.role)) {
            try {
                const superiorData = getUserSuperior();
                if (superiorData.status === 'success' && superiorData.superior) {
                    const teamEmails = getEmailsByEquipo(superiorData.superior);
                    pendingTeamCount = tickets.filter(t =>
                        teamEmails.includes(t['Creado Por']) &&
                        (!t['Tomado Por'] || t['Tomado Por'].toString().trim() === '') &&
                        !['Resuelto', 'Cerrado'].includes(t['Estado'])
                    ).length;
                }
            } catch (radarError) {
                Logger.log("⚠️ Error en Radar: " + radarError.message);
            }
        }

        // --- FILTROS Y PROCESAMIENTO (Código Original) ---
        const getUniqueSortedValues = (key) => [...new Set(tickets.map(t => (t[key] || '').toString().trim()).filter(Boolean))].sort();
        const allPlanillas = getUniqueSortedValues('Planilla (Cliente)');
        const allCreators = getUniqueSortedValues('Creado Por');
        const allMotivosEscalamiento = getUniqueSortedValues('Pedido de escalamiento (Asunto)');
        const superioresResponse = getAllSuperiores();
        const allSuperiores = superioresResponse.status === 'success' ? superioresResponse.superiores : [];

        let filteredTickets = tickets;

        if (options.searchQuery) {
            const query = options.searchQuery.toLowerCase();
            filteredTickets = filteredTickets.filter(t =>
                Object.values(t).some(val => String(val).toLowerCase().includes(query))
            );
        }
        if (options.filterStatus) filteredTickets = filteredTickets.filter(t => t['Estado'] === options.filterStatus);
        if (options.filterTool) filteredTickets = filteredTickets.filter(t => t['Herramienta'] === options.filterTool);
        if (options.filterMotivoEscalamiento) filteredTickets = filteredTickets.filter(t => (t['Pedido de escalamiento (Asunto)'] || '') === options.filterMotivoEscalamiento);
        if (options.filterCreatedBy) filteredTickets = filteredTickets.filter(t => t['Creado Por'] === options.filterCreatedBy);
        if (options.filterPriority) filteredTickets = filteredTickets.filter(t => (t['Prioridad'] || 'Normal') === options.filterPriority);
        if (options.startDate) filteredTickets = filteredTickets.filter(t => new Date(t['Timestamp']) >= new Date(options.startDate));
        if (options.filterEquipo) {
            const emailsDelEquipo = getEmailsByEquipo(options.filterEquipo);
            if (emailsDelEquipo.length > 0) {
                filteredTickets = filteredTickets.filter(t => emailsDelEquipo.includes(t['Creado Por']));
            }
        }
        if (options.endDate) {
            const end = new Date(options.endDate);
            end.setHours(23, 59, 59, 999);
            filteredTickets = filteredTickets.filter(t => new Date(t['Timestamp']) <= end);
        }

        let allTicketsForUser = filteredTickets;
        if (options.filterMyTickets) {
            allTicketsForUser = filteredTickets.filter(t => t['Creado Por'] === userProfile.email);
        }

        allTicketsForUser.sort((a, b) => new Date(b['Timestamp']) - new Date(a['Timestamp']));

        const totalFilteredTickets = allTicketsForUser.length;
        const pageSize = options.pageSize || 10;
        const page = options.page || 1;
        const startIndex = (page - 1) * pageSize;
        const paginatedTickets = allTicketsForUser.slice(startIndex, startIndex + pageSize);

        return {
            tickets: sanitizeTicketData(paginatedTickets),
            allTicketsForStats: sanitizeTicketData(allTicketsForUser),
            totalTickets: totalFilteredTickets,
            allPlanillas,
            allCreators,
            allMotivosEscalamiento,
            allSuperiores,
            pendingTeamTickets: pendingTeamCount,
            status: "ok"
        };
    } catch (error) {
        Logger.log(`ERROR GRAVE al leer datos: ${error.message} | Stack: ${error.stack}`);
        // Devolvemos un error estructurado para que el frontend lo entienda
        return { status: "error", message: `Error de servidor: ${error.message}` };
    }
}
// =============================================================================================================================================================================================

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
            (submissionData.numeroTicket || '').trim(), ticketInterno, 'Registrado', '', 'Normal', ''
        ];
        sheet.appendRow(newRow);
        // Forzamos a Google a escribir FÍSICAMENTE en la hoja antes de seguir.
        // Esto evita que el refresh inmediato lea datos viejos.
        SpreadsheetApp.flush();
        logTicketChange(ticketInterno, `Ticket creado por ${currentUser}`);

        const notificationMessage = `Nuevo ticket #${ticketInterno} creado por ${currentUser} - Asunto: ${submissionData.motivoEscalamiento}`;
        addNotificationToSheet('SUPERUSER', notificationMessage, ticketInterno);
        addNotificationToSheet('AUDITOR', notificationMessage, ticketInterno);
        // 👇 AGREGA ESTO: ¡IMPORTANTE!
        // Borramos la caché para que el próximo refresh muestre el nuevo ticket
        CacheManager.invalidate("TICKETS_MASTER_RAW_DATA");
        return { status: "success", message: "Escalamiento recibido.", ticket: ticketInterno };
    } catch (error) {
        Logger.log(`ERROR GRAVE al crear ticket: ${error.message}. Stack: ${error.stack}`);
        return { status: "error", message: `Error en el servidor al crear el ticket: ${error.message}` };
    }
}

// =============================================================================================================================================================================================

/**
 * FUNCIÓN MEJORADA: Actualiza detalles de ticket con detección precisa de cambios
 */
function updateTicketDetailsLogic(params) {
    try {
        const { ticketId, updates } = params;
        const userProfile = getUserProfileLogic();

        Logger.log(`🔍 Iniciando actualización para ticket: ${ticketId}`);
        Logger.log(`📝 Updates recibidos: ${JSON.stringify(updates)}`);

        if (!['AUDITOR', 'SUPERUSER'].includes(userProfile.role)) {
            throw new Error('Acceso denegado.');
        }

        if (!ticketId || !updates || Object.keys(updates).length === 0) {
            return { status: "info", message: "No se proporcionaron datos para actualizar." };
        }

        const sheet = getOrCreateSheet();
        const dataRange = sheet.getDataRange();
        const values = dataRange.getValues();
        const headers = values.shift();
        const ticketIdColIdx = headers.indexOf('Ticket Interno');

        if (ticketIdColIdx === -1) throw new Error('Columna "Ticket Interno" no encontrada.');

        const rowIndex = values.findIndex(row => row[ticketIdColIdx] === ticketId);
        if (rowIndex === -1) throw new Error(`Ticket ${ticketId} no encontrado.`);

        const sheetRowIndex = rowIndex + 2;
        let changes = [];
        const originalRow = values[rowIndex];

        Logger.log(`📊 Fila original encontrada en índice: ${rowIndex}`);

        for (const fieldName in updates) {
            const colIdx = headers.indexOf(fieldName);
            if (colIdx === -1) {
                Logger.log(`⚠️ Columna "${fieldName}" no encontrada, saltando...`);
                continue;
            }

            const currentValue = originalRow[colIdx];
            let newValue = updates[fieldName];
            let hasChanged = false;

            Logger.log(`🔍 Procesando campo "${fieldName}": "${currentValue}" -> "${newValue}"`);

            if (fieldName.includes('Fecha') && fieldName.includes('Real')) {
                // Manejo de fechas
                if (newValue && newValue.trim() !== '') {
                    try {
                        const parsedDate = new Date(newValue);
                        if (!isNaN(parsedDate.getTime())) {
                            // Comparar fechas normalizando a la misma zona horaria
                            const currentTime = (currentValue instanceof Date) ? currentValue.getTime() : 0;
                            const newTime = parsedDate.getTime();

                            hasChanged = Math.abs(currentTime - newTime) > 60000; // Diferencia > 1 minuto
                            newValue = parsedDate;

                            Logger.log(`📅 Comparación fecha ${fieldName}: ${currentTime} vs ${newTime} (cambió: ${hasChanged})`);
                        } else {
                            Logger.log(`⚠️ Fecha inválida para ${fieldName}: ${newValue}`);
                            continue;
                        }
                    } catch (e) {
                        Logger.log(`❌ Error parseando fecha ${fieldName}: ${e.message}`);
                        continue;
                    }
                } else {
                    // Campo de fecha vacío
                    newValue = '';
                    hasChanged = (currentValue !== '' && currentValue != null);
                    Logger.log(`📅 Limpiando fecha ${fieldName} (cambió: ${hasChanged})`);
                }

            } else if (['Mal Escalado', 'Seguimiento Especial', 'Resolucion Primer Contacto'].includes(fieldName)) {
                const normalizeBoolean = (val) => {
                    if (typeof val === 'boolean') return val;
                    if (typeof val === 'string') {
                        const lower = val.toLowerCase().trim();
                        return lower === 'true' || lower === '1' || lower === 'yes';
                    }
                    return Boolean(val);
                };

                const currentBool = normalizeBoolean(currentValue);
                const newBool = normalizeBoolean(newValue);

                hasChanged = currentBool !== newBool;
                newValue = newBool;

                Logger.log(`☑️ Campo booleano ${fieldName}: ${currentValue}(${currentBool}) -> ${newValue}(${newBool}) (cambió: ${hasChanged})`);

            } else {
                // Campos de texto normales
                const currentStr = String(currentValue || '').trim();
                const newStr = String(newValue || '').trim();
                hasChanged = currentStr !== newStr;
                newValue = newStr;

                Logger.log(`📝 Campo texto ${fieldName}: "${currentStr}" -> "${newStr}" (cambió: ${hasChanged})`);
            }

            if (hasChanged) {
                try {
                    sheet.getRange(sheetRowIndex, colIdx + 1).setValue(newValue);
                    changes.push(`${fieldName}: "${currentValue}" → "${newValue}"`);
                    Logger.log(`✅ ${fieldName} actualizado exitosamente`);
                } catch (e) {
                    Logger.log(`❌ Error actualizando ${fieldName}: ${e.message}`);
                }
            } else {
                Logger.log(`➡️ ${fieldName}: Sin cambios`);
            }
        }

        if (changes.length > 0) {
            const changeDescription = `Auditoría actualizada: ${changes.join('; ')}`;
            logTicketChange(ticketId, changeDescription);

            Logger.log(`✅ Ticket ${ticketId} actualizado con ${changes.length} cambios`);
            CacheManager.invalidate("TICKETS_MASTER_RAW_DATA");
            return {
                status: "success",
                message: `Ticket actualizado correctamente. ${changes.length} campo(s) modificado(s).`,
                changes: changes,
                changesCount: changes.length
            };
        }

        Logger.log(`ℹ️ Ticket ${ticketId}: No se detectaron cambios reales`);
        return {
            status: "info",
            message: "No se detectaron cambios en los datos. Los valores actuales ya coinciden con los enviados.",
            changesCount: 0
        };

    } catch (error) {
        Logger.log(`❌ ERROR en updateTicketDetailsLogic: ${error.message}`);
        Logger.log(`❌ Stack trace: ${error.stack}`);
        return { status: "error", message: `Error en el servidor: ${error.message}` };
    }
}

// =============================================================================================================================================================================================

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
        CacheManager.invalidate("TICKETS_MASTER_RAW_DATA");
        return { status: "success", message: `Ticket ${ticketId} tomado por ti.` };
    } catch (error) {
        Logger.log(`ERROR al tomar ticket: ${error.message}. Stack: ${error.stack}`);
        return { status: "error", message: `Error en el servidor al tomar el ticket: ${error.message}` };
    }
}

// =============================================================================================================================================================================================

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

// =============================================================================================================================================================================================


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