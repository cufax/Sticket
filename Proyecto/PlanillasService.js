// ===============================================================
// SERVICIO DE PLANILLAS MASTER
// ===============================================================

/**
 * Busca una planilla operativa dado un motivo.
 */
function getPlanillaOperativaByMotivoLogic(motivoOperativo) {
  try {
    const sheet = getOrCreatePlanillasMasterSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return null;
    const headers = data.shift();
    const motivoColIdx = headers.indexOf('Motivo operativo');
    const planillaColIdx = headers.indexOf('Planilla operativa');
    if (motivoColIdx === -1 || planillaColIdx === -1) throw new Error('Columnas clave no encontradas en "Planillas_Master".');

    const row = data.find(r => (r[motivoColIdx] || '').toString().trim().toLowerCase() === motivoOperativo.toLowerCase().trim());
    return row ? (row[planillaColIdx] || '').toString().trim() : null;
  } catch (error) {
    Logger.log(`❌ Error en getPlanillaOperativaByMotivo: ${error.message}`);
    throw new Error(`Error al buscar planilla operativa: ${error.message}`);
  }
}

/**
 * Busca la ruta operativa dado un motivo.
 */
function getRutaOperativaByMotivoLogic(motivoOperativo) {
  try {
    const sheet = getOrCreatePlanillasMasterSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return null;
    const headers = data.shift();
    const motivoColIdx = headers.indexOf('Motivo operativo');
    const rutaColIdx = headers.indexOf('Ruta operativa');
    if (motivoColIdx === -1 || rutaColIdx === -1) throw new Error('Columnas clave no encontradas en "Planillas_Master".');

    const row = data.find(r => (r[motivoColIdx] || '').toString().trim().toLowerCase() === motivoOperativo.toLowerCase().trim());
    return row ? (row[rutaColIdx] || '').toString().trim() : null;
  } catch (error) {
    Logger.log(`❌ Error en getRutaOperativaByMotivo: ${error.message}`);
    throw new Error(`Error al buscar ruta operativa: ${error.message}`);
  }
}

/**
 * Obtiene todos los motivos operativos únicos.
 */
function getAllMotivosOperativosLogic() {
  try {
    const sheet = getOrCreatePlanillasMasterSheet();
    const data = sheet.getDataRange().getValues();
    if (data.length <= 1) return [];
    const headers = data.shift();
    const motivoColIdx = headers.indexOf('Motivo operativo');
    if (motivoColIdx === -1) throw new Error('Columna "Motivo operativo" no encontrada.');

    return [...new Set(data.map(row => (row[motivoColIdx] || '').toString().trim()).filter(Boolean))].sort();
  } catch (error) {
    Logger.log(`❌ Error en getAllMotivosOperativos: ${error.message}`);
    throw new Error(`Error al obtener motivos operativos: ${error.message}`);
  }
}
