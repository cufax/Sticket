// ==============================================================================================================================
// SERVICIO DE BOT CON IA (CEREBRO EXPERTO TELECOM) - VERSIÓN FINAL US-CENTRAL1
// ==============================================================================================================================

const AI_CONFIG = {
    // 👇 NUEVA URL OPTIMIZADA (US-CENTRAL1)
    API_URL: 'https://gateway-ai-cerebro-7tupfafqyq-uc.a.run.app',
    API_KEY: 'Er5WsAIwF4MCE3MckIVihlBXzJuYfOgfT7F6zSWPAMh_sRS3kF4hbQ'
};

/**
 * Obtiene el ID de la KB de forma segura.
 */
function getKbIdSafe() {
    try {
        const props = PropertiesService.getScriptProperties();
        const id = props.getProperty('KNOWLEDGE_BASE_DOC_ID');
        // Validación básica: Si es null o parece una URL completa, limpiarlo
        if (!id) return null;
        if (id.includes('/d/')) {
            const match = id.match(/\/d\/([a-zA-Z0-9-_]+)/);
            return match ? match[1] : null;
        }
        return id.trim();
    } catch (e) {
        console.warn("No se pudo leer la propiedad del script:", e);
        return null;
    }
}

/**
 * Búsqueda Omnisciente (Encuentra tickets por cualquier campo)
 */
function findRelevantTickets(userQuery) {
    try {
        if (!userQuery || typeof userQuery !== 'string') return "Consulta vacía.";
        const sheet = getOrCreateSheet();
        const lastRow = sheet.getLastRow();
        if (lastRow <= 1) return "Base de datos vacía.";

        const data = sheet.getDataRange().getValues();
        const headers = data.shift();

        const cols = {
            id: headers.indexOf('Ticket Interno'),
            asunto: headers.indexOf('Pedido de escalamiento (Asunto)'),
            planilla: headers.indexOf('Planilla (Cliente)'),
            estado: headers.indexOf('Estado'),
            creadoPor: headers.indexOf('Creado Por'),
            prioridad: headers.indexOf('Prioridad'),
            fecha: headers.indexOf('Timestamp'),
            herramienta: headers.indexOf('Herramienta')
        };

        if ([cols.id, cols.asunto].includes(-1)) return "Error estructura BD.";

        const queryTokens = userQuery.toLowerCase().split(/\s+/).filter(t => t.length > 2);
        if (queryTokens.length === 0) return "Consulta corta.";

        const relevantTickets = data.filter(row => {
            const searchableText = [
                row[cols.id], row[cols.asunto], row[cols.planilla],
                row[cols.creadoPor], row[cols.estado],
                (cols.herramienta !== -1 ? row[cols.herramienta] : '')
            ].join(' ').toLowerCase();
            return queryTokens.some(token => searchableText.includes(token));
        });

        if (relevantTickets.length === 0) return "Sin coincidencias en tickets.";

        return relevantTickets.slice(0, 10).map(row => {
            const fecha = row[cols.fecha] instanceof Date ? row[cols.fecha].toLocaleDateString() : row[cols.fecha];
            const tool = cols.herramienta !== -1 ? ` | Herramienta: ${row[cols.herramienta]}` : '';
            return `[DB] ID: ${row[cols.id]} | Estado: ${row[cols.estado]} | Prioridad: ${row[cols.prioridad]}${tool} | Fecha: ${fecha}\n   Asunto: "${row[cols.asunto]}"\n   Detalle: "${(row[cols.planilla] || '').toString().substring(0, 150)}..."`;
        }).join('\n\n');

    } catch (e) {
        console.error(`Error búsqueda: ${e.message}`);
        return "Error lectura BD.";
    }
}

/**
 * Obtiene contenido KB de forma segura (Anti-Crash)
 */
function getKnowledgeBaseContentSafe() {
    const docId = getKbIdSafe();
    if (!docId) {
        console.warn("⚠️ AVISO: No hay ID de Base de Conocimiento configurado.");
        return "Nota: No hay base de conocimiento conectada por el momento.";
    }

    try {
        const text = DocumentApp.openById(docId).getBody().getText();
        // Límite seguro para evitar Timeouts
        return text.length > 8000 ? text.substring(0, 8000) + "..." : text;
    } catch (e) {
        console.error("❌ Error leyendo Doc KB (ID inválido o sin permisos):", e);
        return "Nota: Error al leer la base de conocimiento técnica.";
    }
}

/**
 * CEREBRO PRINCIPAL
 */
function getBotResponseLogic(userQuery) {
    try {
        // 1. Contexto Seguro
        const ticketContext = findRelevantTickets(userQuery);
        const knowledgeContext = getKnowledgeBaseContentSafe();

        // 2. Prompt Experto Telecom
        const systemPrompt = `
        ROL:
        Eres "Sticket Assistant", un Ingeniero Experto en Telecomunicaciones y Soporte Nivel 2.
        
        MISIÓN:
        Responder consultas con precisión técnica y una presentación IMPECABLE.

        REGLAS DE FORMATO (ESTRICTAS):
        1. ESTRUCTURA: Usa párrafos cortos. Nunca escribas bloques de texto de más de 3 líneas sin un salto.
        2. LISTAS: Siempre que enumeres pasos, características o tickets, usa viñetas (•) o números (1.).
        3. ESPACIADO: Deja siempre una línea en blanco entre el saludo, el cuerpo de la respuesta y la conclusión.
        4. ÉNFASIS: Usa **negritas** para resaltar:
           - IDs de Tickets (ej: **TK-12345**)
           - Estados (ej: **En Progreso**)
           - Conceptos Clave (ej: **Latencia**, **FTTH**)
        5. CLARIDAD: Si das una explicación técnica, sepárala del estado del ticket con una línea divisoria o un título claro.

        FUENTES DE INFORMACIÓN:
        1. [REGISTROS DB]: Verdad absoluta sobre tickets.
        2. [BASE DE CONOCIMIENTO]: Procedimientos.
        3. [EXPERTO]: Conceptos técnicos generales.

        --- DATOS DEL SISTEMA ---
        
        >> TICKETS ENCONTRADOS:
        ${ticketContext}

        >> KB:
        ${knowledgeContext.substring(0, 6000)} 
        
        --- FIN DATOS ---
        `;

        // 3. Payload
        const payload = {
            query: userQuery,
            system_instruction: systemPrompt,
            model: 'gemini-2.5-flash'
        };

        // 4. Llamada
        const options = {
            method: 'post',
            contentType: 'application/json',
            headers: { 'Authorization': 'Bearer ' + AI_CONFIG.API_KEY },
            payload: JSON.stringify(payload),
            muteHttpExceptions: true
        };

        const response = UrlFetchApp.fetch(AI_CONFIG.API_URL, options);
        const code = response.getResponseCode();
        const text = response.getContentText();

        if (code !== 200) {
            console.error(`Error API IA (${code}): ${text}`);
            return `⚠️ Error de conexión (${code}). El servidor está despertando, intenta de nuevo.`;
        }

        try {
            const json = JSON.parse(text);
            return json.status === 'success' ? json.data.trim() : `Error cerebro: ${json.message}`;
        } catch (e) {
            return "Error técnico: Respuesta ilegible del servidor.";
        }

    } catch (e) {
        console.error(`Error Fatal Bot: ${e.message}`);
        return "Error crítico del sistema.";
    }
}