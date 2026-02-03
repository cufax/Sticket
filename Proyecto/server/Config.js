// ==============================================================================================================================
// CONFIGURACIÓN CENTRALIZADA - PROPIETARIOS DEL CÓDIGO: FACUNDO ALVAREZ  Y ALEJANDRO DAVID GARCIA PARDO
// ==============================================================================================================================

/**
 * @const {string} El ID único de la Hoja de Cálculo de Google que actuará como base de datos.
 */
const SPREADSHEET_ID = "1gHM6mIH8n5LS6_qLQv42ytRuU8a3R1c0mN0pXxXzZ3M";
// =============================================================================================================================================================================================
/**
 * @const {string} El ID de la carpeta de Google Drive donde se guardarán los archivos adjuntos.
 */
const SUPERUSER_DRIVE_FOLDER_ID = "16U50VOK8jSqY9qwQQDuXvXBgYSTiD2mB";
// =============================================================================================================================================================================================
/**
 * @const {Object<string, string>} Un objeto que define roles de usuario para control de acceso.
 */
const USER_ROLES = {
    'alejandro.garciap@konecta.com': 'SUPERUSER',
    'facundo.alvarez@konecta.com': 'SUPERUSER',
    'lautaro.escalante@konecta.com': 'AUDITOR',
    'lucas.maldonado@konecta.com': 'AUDITOR',
    'mauricio.corbalan@konecta.com': 'AUDITOR',
    'franco.alegranza@konecta.com': 'AUDITOR',
    'pablo.rosa@konecta.com': 'AUDITOR',
    'anna.sena@konecta.com': 'AUDITOR',
    'lucas.quevedo@konecta.com': 'USER',
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
    'ivan.mora@konecta.com': 'AUDITOR',
    'iara.garcia@konecta.com': 'USER',
    'maria.olivares@konecta.com': 'USER',
    'leandro.olmos@konecta.com': 'AUDITOR',
    'ivana.piutri@konecta.com': 'USER',
    'luigi.ligo@konecta.com': 'USER',
    'rodrigom.rodriguez@konecta.com': 'AUDITOR',
    'mateo.dominguez@konecta.com': 'USER',
    'tomas.nievas@konecta.com': 'USER',
    'micaela.cariddi@konecta.com': 'AUDITOR',
    'miguel.rojasm@konecta.com': 'USER',
    'lautaron.sosa@konecta.com': 'USER',
    'pauloj.passini@konecta.com': 'USER',
    'joaquin.rey@konecta.com': 'AUDITOR',
    'ezequiel.delavega@konecta.com': 'USER',
    'micaela.gallegos@konecta.com': 'USER',
    'marcelo.lencina@konecta.com': 'USER',
    'ailen.vosahlo@konecta.com': 'AUDITOR',
    'sebastian.ramosa@konecta.com': 'USER',
    'franco.acuna@konecta.com': 'AUDITOR',
    'lucas.avila@konecta.com': 'USER',
    'lautaro.padin@konecta.com': 'USER',
    'ivanap.rodriguez@konecta.com': 'USER',
    'elias.stessens@konecta.com': 'AUDITOR',
    'milagros.spaggiari@konecta.com': 'USER',
    'marcelo.colazo@konecta.com': 'USER',
    'angel.romero@konecta.com': 'AUDITOR',
    'juan.ojeda@konecta.com': 'AUDITOR',
    'liliana.biani@konecta.com': 'AUDITOR',
    'hipolito.blas@konecta.com': 'AUDITOR',
    'daniel.bravo@konecta.com': 'AUDITOR',
    'silvana.bustos@konecta.com': 'AUDITOR',
    'vanessa.cuello@konecta.com': 'AUDITOR',
    'franco.delamora@konecta.com': 'AUDITOR',
    'micaias.gutierrez@konecta.com': 'AUDITOR',
    'enrique.marquez@konecta.com': 'AUDITOR',
    'hectorm.martinez@konecta.com': 'AUDITOR',
    'veronicas.molina@konecta.com': 'AUDITOR',
    'esteladv.piedrabuena@konecta.com': 'AUDITOR',
    'martina.tapia@konecta.com': 'USER',
    'davidn.amaya@konecta.com': 'AUDITOR',
    'agustina.belbruno@konecta.com': 'AUDITOR',
    'anny.cardenas@konecta.com': 'AUDITOR',
    'marcelo.colazo@konecta.com': 'AUDITOR',
    'sheila.diaz@konecta.com': 'AUDITOR',
    'mateo.dominguez@konecta.com': 'AUDITOR',
    'martin.eschoyez@konecta.com': 'AUDITOR',
    'maria.ibanez@konecta.com': 'AUDITOR',
    'karim.iramain@konecta.com': 'AUDITOR',
    'leila.licera@konecta.com': 'AUDITOR',
    'camila.magnaterra@konecta.com': 'AUDITOR',
    'ornella.meolans@konecta.com': 'AUDITOR',
    'tomas.nievas@konecta.com': 'AUDITOR',
    'rodrigom.rodriguez@konecta.com': 'AUDITOR',
    'pablo.rosa@konecta.com': 'AUDITOR',
    'anna.sena@konecta.com': 'AUDITOR',
    'erika.villafane@konecta.com': 'AUDITOR',
    'gonzalo.almada@konecta.com': 'AUDITOR',
    'tomas.arroyo@konecta.com': 'AUDITOR',
    'santiago.malbran@konecta.com': 'AUDITOR',
    'karen.morales@konecta.com': 'AUDITOR',
    'lautaro.padin@konecta.com': 'AUDITOR',
    'jairo.panoso@konecta.com': 'AUDITOR',
    'ivanap.rodriguez@konecta.com': 'AUDITOR',
    'ezequiel.benavidez@konecta.com': 'AUDITOR',
    'nicolas.brumbertolo@konecta.com': 'AUDITOR',
    'luciana.cornacchione@konecta.com': 'AUDITOR',
    'franco.gallardo@konecta.com': 'AUDITOR',
    'iara.garcia@konecta.com': 'AUDITOR',
    'maria.garro@konecta.com': 'AUDITOR',
    'ayelenm.martinez@konecta.com': 'AUDITOR',
    'rocio.medina@konecta.com': 'AUDITOR',
    'mauricio.muller@konecta.com': 'AUDITOR',
    'erika.navas@konecta.com': 'AUDITOR',
    'juan.pereyra@konecta.com': 'AUDITOR',
    'ivana.piutri@konecta.com': 'AUDITOR',
    'maria.fara@konecta.com': 'USER'

};









