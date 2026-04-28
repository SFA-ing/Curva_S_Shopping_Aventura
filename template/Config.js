// ====================================================================
//  CONFIG.JS — Configuración específica del proyecto
//
//  Para adaptar este dashboard a un nuevo proyecto:
//  1. Ajusta las secciones marcadas con  ← EDITAR
//  2. No es necesario modificar Código.js, ETL.js ni Index.html.html
// ====================================================================

// --------------------------------------------------------------------
//  IDENTIDAD DEL PROYECTO                                  ← EDITAR
// --------------------------------------------------------------------
const PROJECT_NAME = "Mi Proyecto — Reporte Avance Físico";

// --------------------------------------------------------------------
//  FUENTE DE DATOS: Excel en Drive                         ← EDITAR
//  ID de la CARPETA de Drive que contiene el archivo Excel origen.
// --------------------------------------------------------------------
const SOURCE_EXCEL_FOLDER_ID = "REEMPLAZAR_CON_ID_CARPETA_DRIVE";

// --------------------------------------------------------------------
//  URL PÚBLICA DEL WEB APP (se actualiza al re-desplegar)  ← EDITAR
// --------------------------------------------------------------------
var REPORTE_FISICO_URL = "REEMPLAZAR_CON_URL_WEBAPP";

// --------------------------------------------------------------------
//  HOJA CACHE DE DATOS                                     ← EDITAR
//  ID del Google Sheets que actúa como caché para el dashboard.
//  DASHBOARD_KEY: clave única del proyecto (sin espacios ni tildes).
// --------------------------------------------------------------------
var CACHE_SHEET_ID = "REEMPLAZAR_CON_ID_GOOGLE_SHEETS_CACHE";
var DASHBOARD_KEY  = "mi_proyecto";

// --------------------------------------------------------------------
//  HOJAS DE GOOGLE SHEETS — destino del ETL               ← EDITAR si cambias nombres
// --------------------------------------------------------------------
const DEST_PLAN_SHEET_NAME      = "Planificación Inicial";
const DEST_REAL_CANT_SHEET_NAME = "AVANCE_REAL_CANT (2)";
const DEST_REAL_HH_SHEET_NAME   = "AVANCE_HH_REAL (2)";
const LOG_SHEET_NAME            = "ETL_LOG";

// --------------------------------------------------------------------
//  HOJAS DEL EXCEL ORIGEN — nombres de pestañas           ← EDITAR si el Excel cambia
// --------------------------------------------------------------------
const SRC_SHEET_SEMANAS   = "SEMANAS";
const SRC_SHEET_PLAN_LB   = "PLANIFICACIÓN LB";
const SRC_SHEET_REAL_CANT = "AVANCE_REAL_CANT";
const SRC_SHEET_REAL_HH   = "AVANCE_HH_REAL";

// --------------------------------------------------------------------
//  UNIDAD FÍSICA PRINCIPAL                                 ← EDITAR
//  Ejemplos: "m" (metros lineales), "m2" (m²), "un" (unidades)
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
//  TAREAS CLAVE — tabla "Avance por Sistema × Tarea"       ← EDITAR
//  Definir entre 1 y 6 tareas. El campo "re" es la regex de búsqueda
//  contra la columna TAREA del plan.
// --------------------------------------------------------------------
const KEY_TASKS = [
  { key: "tarea1", label: "Tarea Clave 1", re: /tarea.?1/i },
  { key: "tarea2", label: "Tarea Clave 2", re: /tarea.?2/i },
  { key: "tarea3", label: "Tarea Clave 3", re: /tarea.?3/i },
  { key: "tarea4", label: "Tarea Clave 4", re: /tarea.?4/i },
];
