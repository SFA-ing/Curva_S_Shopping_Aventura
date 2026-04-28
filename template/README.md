# Curva S — Dashboard de Avance Físico (Template)

Template reutilizable de Google Apps Script para reportes de avance físico de proyectos de construcción.

## Archivos

| Archivo | Descripción |
|---|---|
| `Config.js` | **Único archivo que debes editar** para adaptar el dashboard a un nuevo proyecto |
| `Código.js` | Lógica del servidor (GAS): cálculos, curva S, proyecciones, KPIs |
| `ETL.js` | Extracción y transformación del Excel origen hacia Google Sheets |
| `Index.html.html` | Frontend HTML/CSS/JS del dashboard |

## Cómo crear un nuevo proyecto desde este template

### 1. Crear el Google Apps Script
- Abre el Google Sheets destino → Extensiones → Apps Script
- Crea 4 archivos de script: `Config`, `Código`, `ETL`, `Index` (tipo HTML para el último)
- Pega el contenido de cada archivo de este repo

### 2. Editar `Config.js` (única edición necesaria)

Reemplaza todos los valores marcados con `← EDITAR`:

```javascript
const PROJECT_NAME = "Nombre del Proyecto";
const SOURCE_EXCEL_FOLDER_ID = "1abc...xyz"; // ID carpeta Drive
var REPORTE_FISICO_URL = "https://script.google.com/..."; // URL Web App
var CACHE_SHEET_ID = "1abc...xyz"; // ID Google Sheets caché
var DASHBOARD_KEY  = "nombre_unico_sin_tildes";

const MAIN_UNIT = {
  unit      : "m",   // "m", "m2", "un", etc.
  kpiTitle  : "AVANCE METROS",
  kpiIcon   : "〰",
  kpiColor  : "#0369a1",
  kpiBg     : "#e0f2fe",
  unitLabel : "m",
};

const KEY_TASKS = [
  { key: "tarea1", label: "Nombre Tarea 1", re: /regex_tarea1/i },
  // ... hasta 6 tareas
];
```

### 3. Hojas requeridas en Google Sheets destino
- `Planificación Inicial`
- `AVANCE_REAL_CANT (2)`
- `AVANCE_HH_REAL (2)`
- `ETL_LOG`

### 4. Estructura del Excel origen

| Hoja | Columnas obligatorias |
|---|---|
| `SEMANAS` | `WEEK_START`, `WEEK_KEY` |
| `PLANIFICACIÓN LB` | `Personal`, `ETAPA`, `ACTIVIDAD`, `TAREA`, `Cantidad Teórica`, `Unidad`, `Rendimiento Teorico`, `HH TOTALES` + cols semanales |
| `AVANCE_REAL_CANT` | `PERSONAL`, `ETAPA`, `ACTIVIDAD`, `TAREA`, `Cantidad Real` + cols semanales |
| `AVANCE_HH_REAL` | `Personal`, `ETAPA`, `ACTIVIDAD`, `TAREA` + cols semanales |

### 5. Pasos finales
1. Habilitar **Drive API** en Apps Script → Servicios
2. Menú **Dashboard → ETL: correr ahora**
3. **Desplegar → Nueva implementación → Web App**
4. Pegar la URL en `REPORTE_FISICO_URL` de `Config.js` y re-desplegar
