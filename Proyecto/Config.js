// ===============================================================
// CONFIGURACIÓN CENTRALIZADA
// ===============================================================

/**
 * @const {string} El ID único de la Hoja de Cálculo de Google que actuará como base de datos.
 */
const SPREADSHEET_ID = "1-ZuQDVkzFCVXchRKWh8K2Yq4HSIhq04VX9tw8cQ9HVs";

/**
 * @const {string} El ID de la carpeta de Google Drive donde se guardarán los archivos adjuntos.
 */
const SUPERUSER_DRIVE_FOLDER_ID = "16U50VOK8jSqY9qwQQDuXvXBgYSTiD2mB";

/**
 * @const {Object<string, string>} Un objeto que define roles de usuario para control de acceso.
 */
const USER_ROLES = {
  'alejandro.garciap@konecta.com': 'SUPERUSER',
  'facundo.alvarez@konecta.com': 'SUPERUSER',
  'lautaro.escalante@konecta.com': 'SUPERUSER',
  'lucas.maldonado@konecta.com': 'AUDITOR',
  'mauricio.corbalan@konecta.com': 'AUDITOR',
  'franco.alegranza@konecta.com': 'AUDITOR',
  'pablo.rosa@konecta.com': 'AUDITOR',
  'anna.sena@konecta.com': 'AUDITOR',
  'lucas.quevedo@konecta.com': 'AUDITOR',
  'santiago.malbran@konecta.com': 'AUDITOR',
  'martin.stegemann@konecta.com': 'AUDITOR',
  'gabriel.nesteruk@konecta.com': 'AUDITOR',
  'sergio.gragera@konecta.com': 'AUDITOR',
  'federico.silva@konecta.com': 'AUDITOR',
  'diegon.perez@konecta.com': 'AUDITOR',
  'gonzalo.silva@konecta.com': 'AUDITOR',
  'mauricio.iviris@konecta.com': 'AUDITOR',
  'gustavo.astudillo@konecta.com': 'AUDITOR',
  'sebastian.sosa@konecta.com': 'AUDITOR',
  'elias.stessens@konecta.com': 'AUDITOR'
};
