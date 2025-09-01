// ===============================================================
// SERVICIO DE CACHÉ PARA OPTIMIZACIÓN DE RENDIMIENTO
// ===============================================================

/**
 * Sistema de caché en memoria para optimizar consultas a Google Sheets
 * Reduce significativamente los tiempos de carga
 */

const CACHE_DURATION = 5 * 60 * 1000; // 5 minutos en milisegundos
const dataCache = {
    tickets: null,
    ticketsTimestamp: 0,
    userMetrics: null,
    userMetricsTimestamp: 0,
    planillas: null,
    planillasTimestamp: 0,
    notifications: null,
    notificationsTimestamp: 0
};

/**
 * Obtiene datos del caché si están disponibles y no han expirado
 * @param {string} cacheKey - Clave del caché a consultar
 * @returns {any|null} Datos en caché o null si no están disponibles/expirados
 */
function getCachedData(cacheKey) {
    try {
        const now = Date.now();
        const timestampKey = `${cacheKey}Timestamp`;

        if (dataCache[cacheKey] &&
            dataCache[timestampKey] &&
            (now - dataCache[timestampKey]) < CACHE_DURATION) {

            Logger.log(`✅ Cache HIT para: ${cacheKey} (edad: ${Math.round((now - dataCache[timestampKey]) / 1000)}s)`);
            return dataCache[cacheKey];
        }

        Logger.log(`❌ Cache MISS para: ${cacheKey}`);
        return null;
    } catch (error) {
        Logger.log(`Error en getCachedData: ${error.message}`);
        return null;
    }
}

/**
 * Guarda datos en el caché con timestamp
 * @param {string} cacheKey - Clave del caché
 * @param {any} data - Datos a guardar
 */
function setCachedData(cacheKey, data) {
    try {
        const timestampKey = `${cacheKey}Timestamp`;
        dataCache[cacheKey] = data;
        dataCache[timestampKey] = Date.now();

        Logger.log(`💾 Datos guardados en caché: ${cacheKey} (${Array.isArray(data) ? data.length : 'N/A'} elementos)`);
    } catch (error) {
        Logger.log(`Error en setCachedData: ${error.message}`);
    }
}

/**
 * Invalida el caché para una clave específica o todo el caché
 * @param {string|null} cacheKey - Clave específica a invalidar, o null para limpiar todo
 */
function invalidateCache(cacheKey = null) {
    try {
        if (cacheKey) {
            const timestampKey = `${cacheKey}Timestamp`;
            dataCache[cacheKey] = null;
            dataCache[timestampKey] = 0;
            Logger.log(`🗑️ Caché invalidado para: ${cacheKey}`);
        } else {
            // Limpiar todo el caché
            Object.keys(dataCache).forEach(key => {
                if (key.endsWith('Timestamp')) {
                    dataCache[key] = 0;
                } else {
                    dataCache[key] = null;
                }
            });
            Logger.log(`🗑️ Todo el caché ha sido invalidado`);
        }
    } catch (error) {
        Logger.log(`Error en invalidateCache: ${error.message}`);
    }
}

/**
 * Obtiene estadísticas del caché para debugging
 * @returns {Object} Estadísticas del caché
 */
function getCacheStats() {
    const stats = {};
    const now = Date.now();

    Object.keys(dataCache).forEach(key => {
        if (!key.endsWith('Timestamp')) {
            const timestampKey = `${key}Timestamp`;
            const timestamp = dataCache[timestampKey] || 0;
            const age = timestamp > 0 ? Math.round((now - timestamp) / 1000) : -1;
            const isValid = timestamp > 0 && (now - timestamp) < CACHE_DURATION;

            stats[key] = {
                hasData: dataCache[key] !== null,
                ageSeconds: age,
                isValid: isValid,
                size: Array.isArray(dataCache[key]) ? dataCache[key].length : (dataCache[key] ? 1 : 0)
            };
        }
    });

    return stats;
}
