// ===============================================================
// SERVICIO DE BOT CON IA (GEMINI)
// ===============================================================

/**
 * Obtiene propiedades del script (ID de proyecto, etc.).
 */
function getScriptProperties() {
  const scriptProperties = PropertiesService.getScriptProperties();
  const gcpProjectId = scriptProperties.getProperty('GCP_PROJECT_ID');
  const knowledgeBaseDocId = scriptProperties.getProperty('KNOWLEDGE_BASE_DOC_ID');
  if (!gcpProjectId || !knowledgeBaseDocId) {
    throw new Error('Propiedades del script "GCP_PROJECT_ID" y "KNOWLEDGE_BASE_DOC_ID" deben estar configuradas.');
  }
  return { gcpProjectId, knowledgeBaseDocId };
}

/**
 * Busca tickets relevantes para una consulta.
 */
function findRelevantTickets(userQuery) {
  try {
    if (!userQuery || typeof userQuery !== 'string') return "No se proporcionó una consulta válida.";
    const sheet = getOrCreateSheet();
    if (sheet.getLastRow() <= 1) return "No hay tickets en la base de datos.";

    const data = sheet.getDataRange().getValues();
    const headers = data.shift();
    const asuntoColIdx = headers.indexOf('Pedido de escalamiento (Asunto)');
    const planillaColIdx = headers.indexOf('Planilla (Cliente)');
    if (asuntoColIdx === -1 || planillaColIdx === -1) return `Error de configuración en la hoja "Tickets_Master".`;

    const estadoColIdx = headers.indexOf('Estado');
    const ticketIdColIdx = headers.indexOf('Ticket Interno');
    const queryTokens = userQuery.toLowerCase().split(/\s+/).filter(token => token.length > 2);
    if (queryTokens.length === 0) return "No se encontraron tickets relevantes.";

    const relevantTickets = data
      .filter(row => queryTokens.some(token =>
      ((row[asuntoColIdx] || '').toString().toLowerCase().includes(token) ||
        (row[planillaColIdx] || '').toString().toLowerCase().includes(token))
      ))
      .slice(-5) // Tomar los 5 más recientes
      .map(row => `  - Ticket: ${row[ticketIdColIdx]}, Asunto: "${row[asuntoColIdx]}", Planilla: "${row[planillaColIdx]}", Estado: ${row[estadoColIdx]}`)
      .join('\n');

    return relevantTickets.length > 0 ? relevantTickets : "No se encontraron tickets relevantes para la consulta.";
  } catch (e) {
    Logger.log(`Error buscando tickets: ${e.message}`);
    return "Error al acceder a la base de datos de tickets.";
  }
}

/**
 * Obtiene contenido de un Google Doc como base de conocimiento.
 */
function getKnowledgeBaseContent(docId) {
  try {
    return DocumentApp.openById(docId).getBody().getText();
  } catch (e) {
    Logger.log(`Error leyendo el documento de conocimiento: ${e.message}`);
    return "No se pudo acceder a la base de conocimiento.";
  }
}

/**
 * Obtiene una respuesta del bot de IA.
 */
function getBotResponseLogic(userQuery) {
  try {
    const { gcpProjectId, knowledgeBaseDocId } = getScriptProperties();
    const ticketContext = findRelevantTickets(userQuery);
    const knowledgeContext = getKnowledgeBaseContent(knowledgeBaseDocId);
    const prompt = `
      Eres "Sticket Pro Assistant", un asistente experto en gestión de tickets.
      Responde a la pregunta del usuario basándote ESTRICTAMENTE en el contexto proporcionado.
      Si la respuesta no está en el contexto, indica que no tienes esa información. Sé conciso.
      --- CONTEXTO DE TICKETS RELEVANTES ---
      ${ticketContext}
      --- BASE DE CONOCIMIENTO (PROCEDIMIENTOS) ---
      ${knowledgeContext}
      ---
      Pregunta del Usuario: "${userQuery}"
      Respuesta:
    `;

    const geminiApiUrl = `https://us-central1-aiplatform.googleapis.com/v1/projects/${gcpProjectId}/locations/us-central1/publishers/google/models/gemini-1.5-flash-001:generateContent`;
    const token = ScriptApp.getOAuthToken();
    const payload = { contents: [{ parts: [{ text: prompt }] }] };
    const options = {
      method: 'post',
      contentType: 'application/json',
      headers: { 'Authorization': 'Bearer ' + token },
      payload: JSON.stringify(payload)
    };
    const response = UrlFetchApp.fetch(geminiApiUrl, options);
    const responseData = JSON.parse(response.getContentText());
    return responseData.candidates[0].content.parts[0].text.trim();
  } catch (e) {
    Logger.log(`Error en getBotResponse: ${e.message} - Stack: ${e.stack}`);
        return "Lo siento, tuve un problema para procesar tu solicitud. Intenta de nuevo.";
    }
}

/**
 * Obtiene el historial de solicitudes hechas al bot de Gemini.
 */
function getGbRequestsHistoryLogic() {
    try {
        const SHEET_NAME = "GB_Request_History";
        const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
        const sheet = ss.getSheetByName(SHEET_NAME);
        if (!sheet) {
            return { status: "success", data: [] };
        }
        const rows = sheet.getDataRange().getValues();
        rows.shift(); // quitar encabezados
        return { status: "success", data: rows };
    } catch (e) {
        Logger.log(`Error en getGbRequestsHistoryLogic: ${e.message}`);
        return { status: "error", message: e.message };
    }
}
