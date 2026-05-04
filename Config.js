// ====================================================================
//  CONFIG.JS — Configuración específica del proyecto FEMSA COCA
// ====================================================================

// --------------------------------------------------------------------
//  IDENTIDAD DEL PROYECTO
// --------------------------------------------------------------------
const PROJECT_NAME = "FEMSA COCA — Reporte Avance Físico";

// --------------------------------------------------------------------
//  FUENTE DE DATOS: Excel en Drive
//  ID de la CARPETA de Drive que contiene el archivo Excel origen.
// --------------------------------------------------------------------
const SOURCE_EXCEL_FOLDER_ID = "REEMPLAZAR_CON_ID_CARPETA_DRIVE";

// --------------------------------------------------------------------
//  URL PÚBLICA DEL WEB APP (se actualiza al re-desplegar)
// --------------------------------------------------------------------
var REPORTE_FISICO_URL = "REEMPLAZAR_CON_URL_WEBAPP";

// --------------------------------------------------------------------
//  HOJA CACHE DE DATOS
//  ID del Google Sheets que actúa como caché para el dashboard.
//  DASHBOARD_KEY: clave única del proyecto (sin espacios ni tildes).
// --------------------------------------------------------------------
var CACHE_SHEET_ID = "REEMPLAZAR_CON_ID_GOOGLE_SHEETS_CACHE";
var DASHBOARD_KEY  = "femsa_coca";

// --------------------------------------------------------------------
//  HOJAS DE GOOGLE SHEETS — destino del ETL
// --------------------------------------------------------------------
const DEST_PLAN_SHEET_NAME      = "Planificación Inicial";
const DEST_REAL_CANT_SHEET_NAME = "AVANCE_REAL_CANT (2)";
const DEST_REAL_HH_SHEET_NAME   = "AVANCE_HH_REAL (2)";
const LOG_SHEET_NAME            = "ETL_LOG";

// --------------------------------------------------------------------
//  HOJAS DEL EXCEL ORIGEN — nombres de pestañas
// --------------------------------------------------------------------
const SRC_SHEET_SEMANAS   = "SEMANAS";
const SRC_SHEET_PLAN_LB   = "PLANIFICACIÓN LB";
const SRC_SHEET_REAL_CANT = "AVANCE_REAL_CANT";
const SRC_SHEET_REAL_HH   = "AVANCE_HH_REAL";

// --------------------------------------------------------------------
//  UNIDAD FÍSICA PRINCIPAL
// --------------------------------------------------------------------
const MAIN_UNIT = {
  unit      : "m",                        // valor en columna "Unidad" del plan (minúsculas)
  kpiTitle  : "AVANCE UNIDAD PRINCIPAL",  // título de la barra KPI
  kpiIcon   : "〰",                       // emoji/símbolo del ícono
  kpiColor  : "#0369a1",                  // color del texto/barra
  kpiBg     : "#e0f2fe",                  // fondo del ícono
  unitLabel : "m",                        // sufijo de unidad en textos
};

// --------------------------------------------------------------------
//  TAREAS CLAVE — tabla "Avance por Sistema × Tarea"
// --------------------------------------------------------------------
const KEY_TASKS = [
  { key: "tarea1", label: "Tarea Clave 1", re: /tarea.?1/i },
  { key: "tarea2", label: "Tarea Clave 2", re: /tarea.?2/i },
  { key: "tarea3", label: "Tarea Clave 3", re: /tarea.?3/i },
  { key: "tarea4", label: "Tarea Clave 4", re: /tarea.?4/i },
];
