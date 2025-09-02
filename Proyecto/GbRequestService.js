// ===============================================================
// SERVICIO DE SOLICITUD DE GBS
// ===============================================================

/**
 * Registra una solicitud de GBs en la hoja correspondiente.
 * @param {Object} requestData - Datos de la solicitud.
 */
function submitGbRequestLogic(requestData) {
    try {
        const sheet = getOrCreateGbRequestSheet();
        const timestamp = new Date();
        const row = [
            timestamp,
            requestData.usuario.trim(),
            requestData.idGestion.trim(),
            requestData.numeroContrato.trim(),
            requestData.numeroLinea.trim(),
            requestData.smsBackup.trim(),
            requestData.motivo.trim(),
            false
        ];
        sheet.appendRow(row);
        return { status: "success", message: "Solicitud de GBs registrada." };
    } catch (error) {
        Logger.log(`ERROR al registrar solicitud de GBs: ${error.message}.`);
        return { status: "error", message: `Error al registrar la solicitud: ${error.message}` };
    }
}

/**
 * Obtiene o crea la hoja donde se almacenan las solicitudes de GBs.
 */
function getOrCreateGbRequestSheet() {
    const SHEET_NAME = "Solicitudes_GBs";
    const HEADERS = [
        'Timestamp',
        'Usuario',
        'ID Gestion FAN',
        'Numero de Contrato',
        'Numero de Linea',
        'Recibio SMS WiFi BackUp',
        'Motivo de Escalamiento',
        'Revisado'
    ];
    try {
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        let sheet = ss.getSheetByName(SHEET_NAME);
        if (!sheet) {
            sheet = ss.insertSheet(SHEET_NAME);
            sheet.appendRow(HEADERS);
        } else if (sheet.getLastRow() === 0) {
            sheet.appendRow(HEADERS);
        }
        return sheet;
    } catch (error) {
        Logger.log(`ERROR al obtener hoja ${SHEET_NAME}: ${error.message}`);
        throw error;
    }
}

/**
 * Obtiene todas las solicitudes de GBs registradas.
 * Devuelve un array de objetos con la información y el número de fila.
 */
function getGbRequestsHistoryLogic() {
    try {
        const sheet = getOrCreateGbRequestSheet();
        const data = sheet.getDataRange().getValues();
        if (data.length <= 1) {
            return { status: "success", data: [] };
        }
        const rows = [];
        for (let i = 1; i < data.length; i++) {
            const row = data[i];
            rows.push({
                rowNumber: i + 1,
                timestamp: row[0],
                usuario: row[1],
                idGestion: row[2],
                numeroContrato: row[3],
                numeroLinea: row[4],
                smsBackup: row[5],
                motivo: row[6],
                revisado: row[7] === true
            });
        }
        return { status: "success", data: rows };
    } catch (error) {
        Logger.log(`ERROR al obtener historial de GBs: ${error.message}`);
        return { status: "error", message: error.message };
    }
}

/**
 * Actualiza el estado de revisión de una solicitud de GBs.
 * @param {Object} params - Parámetros con número de fila y nuevo estado.
 */
function updateGbRequestAuditStatusLogic(params) {
    try {
        const sheet = getOrCreateGbRequestSheet();
        sheet.getRange(params.rowNumber, 8).setValue(params.revisado);
        return { status: "success", message: "Estado de revisión actualizado." };
    } catch (error) {
        Logger.log(`ERROR al actualizar revisión de GBs: ${error.message}`);
        return { status: "error", message: error.message };
    }
}
