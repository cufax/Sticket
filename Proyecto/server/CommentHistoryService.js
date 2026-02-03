// ==============================================================================================================================
// SERVICIO DE COMENTARIOS E HISTORIAL - PROPIETARIOS DEL CÓDIGO: FACUNDO ALVAREZ  Y ALEJANDRO DAVID GARCIA PARDO
// ==============================================================================================================================

// =============================================================================================================================================================================================
/**
 * Añade un comentario a un ticket específico.
 */
function addTicketCommentLogic(commentData) {
    try {
        const { ticketId, commentText } = commentData;
        if (!ticketId || !commentText || commentText.trim() === '') {
            throw new Error('El ID del ticket y el texto del comentario son obligatorios.');
        }
        const commentsSheet = getOrCreateCommentsSheet();
        const currentUser = Session.getActiveUser().getEmail();
        const timestamp = new Date();
        commentsSheet.appendRow([ticketId, timestamp, currentUser, commentText.trim()]);
        Logger.log(`Comentario añadido al ticket ${ticketId} por ${currentUser}.`);
        return { status: "success", message: "Comentario añadido correctamente." };
    } catch (error) {
        Logger.log(`ERROR al añadir comentario: ${error.message}.`);
        return { status: "error", message: `Error al añadir comentario: ${error.message}` };
    }
}

// =============================================================================================================================================================================================

/**
 * Obtiene todos los comentarios para un ticket específico.
 */
function getTicketCommentsLogic(ticketId) {
    try {
        const commentsSheet = getOrCreateCommentsSheet();
        const data = commentsSheet.getDataRange().getValues();
        if (data.length <= 1) return { comments: [], status: "ok" };

        const headers = data.shift();
        const ticketIdColIdx = headers.indexOf('Ticket Interno');

        const comments = data
            .filter(row => row[ticketIdColIdx] === ticketId)
            .map(row => ({
                ticketInterno: row[headers.indexOf('Ticket Interno')],
                timestamp: row[headers.indexOf('Timestamp')] instanceof Date ? row[headers.indexOf('Timestamp')].toISOString() : row[headers.indexOf('Timestamp')],
                autor: row[headers.indexOf('Autor')],
                comentario: row[headers.indexOf('Comentario')]
            }))
            .sort((a, b) => new Date(a.timestamp) - new Date(b.timestamp));

        return { comments: comments, status: "ok" };
    } catch (error) {
        Logger.log(`ERROR al obtener comentarios: ${error.message}.`);
        return { status: "error", message: `Error al obtener comentarios: ${error.message}` };
    }
}

// =============================================================================================================================================================================================

/**
 * Registra un cambio en el historial de un ticket.
 */
function logTicketChange(ticketId, changeDescription) {
    try {
        const historySheet = getOrCreateHistorySheet();
        const timestamp = new Date();
        const userEmail = Session.getActiveUser().getEmail();
        historySheet.appendRow([ticketId, timestamp, changeDescription, userEmail]);
    } catch (e) {
        Logger.log(`Error al registrar cambio para ticket ${ticketId}: ${e.message}`);
    }
}

// =============================================================================================================================================================================================

/**
 * Obtiene el historial de cambios de un ticket específico.
 */
function getTicketHistoryLogic(ticketId) {
    try {
        if (!ticketId || ticketId.trim() === '') throw new Error('ID de ticket inválido.');
        const historySheet = getOrCreateHistorySheet();
        const data = historySheet.getDataRange().getValues();
        if (data.length <= 1) return { status: "success", history: [] };

        const headers = data.shift();
        const historyData = data
            .filter(row => row[0] === ticketId)
            .map(row => ({
                ticketInterno: row[0],
                timestamp: row[1] instanceof Date ? row[1].toISOString() : row[1],
                cambio: row[2],
                usuario: row[3]
            }))
            .sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));

        return { status: "success", history: historyData };
    } catch (error) {
        Logger.log(`ERROR al obtener historial del ticket ${ticketId}: ${error.message}.`);
        return { status: "error", message: `Error al obtener historial: ${error.message}`, history: [] };
    }
}