// ===============================================================
// UTILIDADES GENERALES
// ===============================================================

/**
 * Obtiene el perfil del usuario activo (email, rol y foto de perfil).
 * Limita el acceso solo a los usuarios definidos en USER_ROLES.
 */
function getUserProfileLogic() {
  try {
    const email = Session.getActiveUser().getEmail();
    if (!email) {
      throw new Error('No se pudo obtener el email del usuario activo.');
    }

    if (!USER_ROLES.hasOwnProperty(email)) {
      Logger.log(`🚫 Acceso denegado para usuario no autorizado: ${email}`);
      throw new Error('Acceso denegado. Tu cuenta no está autorizada para usar esta aplicación.');
    }

    const role = USER_ROLES[email];
    const photoUrl = generateProfileImageWithFallbacks(email);

    Logger.log(`👤 Perfil generado para: ${email}, Rol: ${role}, Foto: ${photoUrl}`);

    return {
      email: email,
      role: role,
      photoUrl: photoUrl,
      status: "success"
    };
  } catch (e) {
    Logger.log(`Error crítico al obtener perfil de usuario: ${e.message}`);
    throw new Error(`No se pudo obtener el perfil de usuario: ${e.message}`);
  }
}


/**
 * Genera una foto de perfil con múltiples fallbacks.
 */
function generateProfileImageWithFallbacks(email) {
  try {
    const cleanEmail = email.toLowerCase().trim();
    const emailHash = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, cleanEmail)
      .map(byte => {
        const unsignedByte = byte < 0 ? byte + 256 : byte;
        return unsignedByte.toString(16).padStart(2, '0');
      })
      .join('');

    const fallbackUrl = `https://anonymous-animals.azurewebsites.net/avatar/${encodeURIComponent(cleanEmail)}`;
    const gravatarUrl = `https://www.gravatar.com/avatar/${emailHash}?s=80&d=${encodeURIComponent(fallbackUrl)}&r=pg`;

    Logger.log(`🔗 Gravatar URL con fallback: ${gravatarUrl}`);
    return gravatarUrl;
  } catch (e) {
    Logger.log(`Error generando imagen con fallbacks: ${e.message}`);
    return `https://api.dicebear.com/8.x/micah/png?seed=${encodeURIComponent(email)}&radius=50`;
  }
}


/**
 * 🔧 FUNCIÓN DE SANITIZACIÓN ROBUSTA PARA EL BACKEND - VERSIÓN DEFINITIVA
 * Convierte un array de filas crudas (array de arrays) en un array de objetos de ticket limpios.
 * Es el único punto donde los datos de la hoja de cálculo se transforman en objetos.
 * @param {Array<Array<any>>} rawData - Array de filas directamente desde la hoja.
 * @param {Array<string>} headers - Array con los nombres de las columnas.
 * @returns {Array<Object>} - Array de objetos de ticket sanitizados.
 */
function sanitizeTicketData(rawData, headers) {
  const headerMap = {};
  headers.forEach((header, index) => {
    headerMap[header] = index;
  });

  const dateColumns = ['Timestamp', 'Fecha Resolucion Real']; // Solo procesamos las fechas que existen

  return rawData.map(row => {
    const ticket = {};
    for (const header in headerMap) {
      const index = headerMap[header];
      let value = row[index];

      if (dateColumns.includes(header)) {
        // Procesa las columnas de fecha, convirtiéndolas a formato ISO String para el transporte
        if (value instanceof Date && !isNaN(value)) {
          ticket[header] = value.toISOString();
        } else {
          ticket[header] = null; // Si no es una fecha válida, se envía como nulo
        }
      } else if (typeof value === 'boolean') {
        ticket[header] = value;
      } else {
        // Para el resto de los campos, se convierten a String o a un string vacío
        ticket[header] = (value === null || value === undefined) ? '' : String(value).trim();
      }
    }
    return ticket;
  });
}


function ejemploUsoCommentsSheet() {
  try {
    const commentsSheet = getOrCreateCommentsSheet();
    const validation = validateCommentsSheet(commentsSheet);
    if (!validation.valid) throw new Error(`Hoja inválida: ${validation.error}`);
    Logger.log('✅ Hoja de comentarios lista para usar');
    addCommentSafely('TK-TEST', 'Usuario Test', 'Este es un comentario de prueba');
  } catch (error) {
    Logger.log(`❌ Error en ejemploUsoCommentsSheet: ${error.message}`);
  }
}

function addCommentSafely(ticketInterno, autor, comentario) {
  try {
    const commentsSheet = getOrCreateCommentsSheet();
    const validation = validateCommentsSheet(commentsSheet);
    if (!validation.valid) throw new Error(`Hoja de comentarios inválida: ${validation.error}`);
    commentsSheet.appendRow([ticketInterno, new Date(), autor, comentario]);
    Logger.log(`✅ Comentario añadido para ticket ${ticketInterno}`);
    return true;
  } catch (error) {
    Logger.log(`❌ Error al añadir comentario: ${error.message}`);
    return false;
  }
}

function ejemploUsoHistorySheet() {
  try {
    const historySheet = getOrCreateHistorySheet();
    const validation = validateHistorySheet(historySheet);
    if (!validation.valid) throw new Error(`Hoja inválida: ${validation.error}`);
    Logger.log('✅ Hoja de historial lista para usar');
  } catch (error) {
    Logger.log(`❌ Error en ejemploUsoHistorySheet: ${error.message}`);
  }
}
