// ===============================================================
// SERVICIO DE NOTIFICACIONES
// ===============================================================

/**
 * Añade una notificación a la hoja de notificaciones.
 */
function addNotificationToSheet(recipient, message, ticketId = null) {
  try {
    const sheet = repairNotificationsStructure(); // Ensures sheet is valid
    const timestamp = new Date();
    let recipientsList = [];

    if (recipient === 'SUPERUSER' || recipient === 'AUDITOR') {
      recipientsList = Object.keys(USER_ROLES).filter(email => USER_ROLES[email] === recipient);
    } else {
      recipientsList.push(recipient);
    }

    recipientsList.forEach(email => {
      sheet.appendRow([timestamp, email, message, false, ticketId || '']);
    });
    Logger.log(`Notificación añadida para ${recipient}: ${message}`);
  } catch (e) {
    Logger.log(`❌ Error al añadir notificación: ${e.message}`);
  }
}


/**
 * 🔧 REPARA ESTRUCTURA DE NOTIFICACIONES
 */
function repairNotificationsStructure() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  let sheet = ss.getSheetByName("Notifications");
  const HEADERS = ['Timestamp', 'Recipient', 'Message', 'Read', 'TicketId'];
  if (!sheet) {
    sheet = ss.insertSheet("Notifications");
    sheet.getRange(1, 1, 1, 5).setValues([HEADERS]).setFontWeight("bold");
  } else {
    if (sheet.getLastRow() === 0) {
      sheet.clear();
      sheet.getRange(1, 1, 1, 5).setValues([HEADERS]).setFontWeight("bold");
    } else {
      const existingHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const missingHeaders = HEADERS.filter(h => !existingHeaders.includes(h));
      if (missingHeaders.length > 0) {
        let lastCol = sheet.getLastColumn();
        missingHeaders.forEach(header => {
          lastCol++;
          sheet.getRange(1, lastCol).setValue(header).setFontWeight("bold");
        });
      }
    }
  }
  return sheet;
}

/**
 * Obtiene las notificaciones no leídas para el usuario actual.
 */
function getUnreadNotificationsLogic() {
  try {
    const sheet = repairNotificationsStructure(); // Ensures sheet exists and is valid
    if (sheet.getLastRow() <= 1) return { notifications: [], status: "ok" };

    const currentUserEmail = Session.getActiveUser().getEmail();
    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const recipientColIdx = headers.indexOf('Recipient');
    const readColIdx = headers.indexOf('Read');
    if (recipientColIdx === -1 || readColIdx === -1) return { notifications: [], status: "ok" };

    const unreadNotifications = data
      .map((row, index) => ({ row, index }))
      .filter(({ row }) => row[recipientColIdx] === currentUserEmail && row[readColIdx] === false)
      .map(({ row, index }) => ({
        row: index + 2, // Sheet row number is index + 2 (1 for header, 1 for 0-based)
        message: row[headers.indexOf('Message')],
        timestamp: row[headers.indexOf('Timestamp')].toISOString(),
        ticketId: row[headers.indexOf('TicketId')]
      }))
      .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));

    return { notifications: unreadNotifications, status: "ok" };
  } catch (e) {
    Logger.log(`❌ Error al obtener notificaciones: ${e.message}`);
    return { status: "error", message: `Error al obtener notificaciones: ${e.message}` };
  }
}

/**
 * Marca una notificación como leída.
 */
function markNotificationAsReadLogic(row) {
  try {
    const sheet = repairNotificationsStructure();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const readColIdx = headers.indexOf('Read');
    const recipientColIdx = headers.indexOf('Recipient');
    if (readColIdx === -1 || recipientColIdx === -1) throw new Error('Columnas de notificación no encontradas.');

    const currentUserEmail = Session.getActiveUser().getEmail();
    const recipientEmail = sheet.getRange(row, recipientColIdx + 1).getValue();

    if (recipientEmail === currentUserEmail) {
      sheet.getRange(row, readColIdx + 1).setValue(true);
      return { status: "success", message: "Notificación marcada como leída." };
    } else {
      throw new Error('Acceso denegado a esta notificación.');
    }
  } catch (e) {
    Logger.log(`Error al marcar notificación como leída: ${e.message}`);
    return { status: "error", message: `Error: ${e.message}` };
  }
}
