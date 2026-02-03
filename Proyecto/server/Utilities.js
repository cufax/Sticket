// ==============================================================================================================================
// UTILIDADES GENERALES - PROPIETARIOS DEL CÓDIGO: FACUNDO ALVAREZ  Y ALEJANDRO DAVID GARCIA PARDO
// ==============================================================================================================================

/**
 * Obtiene el perfil del usuario activo (email, rol y foto de perfil).
 * Limita el acceso solo a los usuarios definidos en USER_ROLES.
 */
function getUserProfileLogic() { //MODIFICACIÓN LOGICA DE USUARIOS 
    try {
        const email = Session.getActiveUser().getEmail();
        if (!email) {
            throw new Error('No se pudo obtener el email del usuario activo.');
        }
        // Verifica si el email del usuario está en la lista de USER_ROLES
        if (!USER_ROLES.hasOwnProperty(email)) {
            // Si el usuario no está en la lista, denegar el acceso
            Logger.log(`🚫 Acceso denegado para usuario no autorizado: ${email}`);
            throw new Error('Acceso denegado. Tu cuenta no está autorizada para usar esta aplicación.');
        }

        const role = USER_ROLES[email]; // Si llegamos aquí, el usuario está en USER_ROLES

        // Generar foto de perfil usando Gravatar + fallback
        const photoUrl = generateProfileImageWithFallbacks(email);

        Logger.log(`👤 Perfil generado para: ${email}, Rol: ${role}, Foto: ${photoUrl}`); // Debug

        return {
            email: email,
            role: role,
            photoUrl: photoUrl,
            status: "success"
        };
    } catch (e) {
        Logger.log(`Error crítico al obtener perfil de usuario: ${e.message}`);
        // Propagar el error para que el frontend lo maneje
        throw new Error(`No se pudo obtener el perfil de usuario: ${e.message}`);
    }
}

// =============================================================================================================================================================================================


/**
 * Genera una foto de perfil con múltiples fallbacks.
 */
function generateProfileImageWithFallbacks(email) {
    try {
        // Limpiar y normalizar el email
        const cleanEmail = email.toLowerCase().trim();

        // Generar hash MD5 para Gravatar
        const emailHash = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, cleanEmail)
            .map(byte => {
                const unsignedByte = byte < 0 ? byte + 256 : byte;
                return unsignedByte.toString(16).padStart(2, '0');
            })
            .join('');

        const fallbackUrl = `https://anonymous-animals.azurewebsites.net/avatar/${encodeURIComponent(cleanEmail)}`;

        // Esta línea USA la fallbackUrl
        const gravatarUrl = `https://www.gravatar.com/avatar/${emailHash}?s=80&d=${encodeURIComponent(fallbackUrl)}&r=pg`;

        Logger.log(`🔗 Gravatar URL con fallback a DiceBear: ${gravatarUrl}`);

        return gravatarUrl;

    } catch (e) {
        Logger.log(`Error generando imagen con fallbacks: ${e.message}`);

        return `https://api.dicebear.com/8.x/micah/png?seed=${encodeURIComponent(email)}&radius=50`;
    }
}

// =============================================================================================================================================================================================

function sanitizeTicketData(tickets) {
    return tickets.map(ticket => {
        const sanitized = {};
        for (const field in ticket) {
            if (Object.prototype.hasOwnProperty.call(ticket, field)) {
                let value = ticket[field];

                // Preservar los booleanos
                if (typeof value === 'boolean') {
                    sanitized[field] = value;
                }
                // Convertir las fechas a formato ISO para consistencia
                else if (value instanceof Date) {
                    sanitized[field] = value.toISOString();
                }
                // Manejar nulos/indefinidos y convertir el resto a string
                else {
                    sanitized[field] = (value === null || value === undefined) ? '' : String(value);
                }
            }
        }
        return sanitized;
    });
}