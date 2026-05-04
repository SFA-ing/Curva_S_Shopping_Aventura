# FEMSA COCA — Curva S / Dashboard Avance Físico

Proyecto basado en el template `Curva S` del repo `SFA-ing/Curva_S_Shopping_Aventura` rama `template`.

## Archivos clave

| Archivo | Descripción |
|---|---|
| `Config.js` | Archivo de configuración específico del proyecto FEMSA COCA |
| `Código.js` | Lógica del servidor (GAS): cálculos, curva S, proyecciones, KPIs |
| `ETL.js` | Extracción y transformación del Excel origen hacia Google Sheets |
| `Index.html.html` | Frontend HTML/CSS/JS del dashboard |

## Pasos inmediatos

1. Actualizar `Config.js` con los valores reales de FEMSA COCA:
   - `SOURCE_EXCEL_FOLDER_ID`
   - `REPORTE_FISICO_URL`
   - `CACHE_SHEET_ID`
   - `DASHBOARD_KEY`
2. Desplegar la Web App en Google Apps Script.
3. Ejecutar el ETL y verificar las hojas destino.
