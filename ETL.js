/***************
 * Las constantes de configuración viven en Config.js.
 * Este archivo solo contiene la lógica de transformación.
 ***************/

/**
 * ETL end-to-end:
 * 1) Convierte el Excel a Google Sheets (temporal)
 * 2) Transforma y carga las 3 tablas destino en el spreadsheet activo
 * 3) Borra el temporal
 */
function ETL_actualizarTodo() {
  const destSS = SpreadsheetApp.getActive();
  let tempGsheetId = null;
  let sourceFile = null;

  try {
    sourceFile = resolveSourceExcelFromFolder_();
    tempGsheetId = convertExcelToGoogleSheet_(sourceFile.id);
    const srcSS = SpreadsheetApp.openById(tempGsheetId);

    const weekMap = buildWeekMap_(srcSS.getSheetByName(SRC_SHEET_SEMANAS));

    const planRows = transformPlanificacionLB_(srcSS.getSheetByName(SRC_SHEET_PLAN_LB), weekMap);

    // FIX 1: Se pasa el flag `hasCantTotalCol` según la hoja.
    // AVANCE_REAL_CANT tiene una columna extra "Cantidad Real" (total acumulado)
    // entre TAREA y las semanas. AVANCE_HH_REAL NO la tiene.
    const realCantRows = transformAvanceRealMatriz_(srcSS.getSheetByName(SRC_SHEET_REAL_CANT), weekMap, 1, true);
    const realHHRows   = transformAvanceRealMatriz_(srcSS.getSheetByName(SRC_SHEET_REAL_HH),   weekMap, 1, false);

    writeTable_(destSS.getSheetByName(DEST_PLAN_SHEET_NAME),      planRows);
    writeTable_(destSS.getSheetByName(DEST_REAL_CANT_SHEET_NAME), realCantRows);
    writeTable_(destSS.getSheetByName(DEST_REAL_HH_SHEET_NAME),   realHHRows);

    logETL_({
      status: "OK",
      fileId: sourceFile.id,
      detail: "Carga completa | " + sourceFile.name,
      planRows: planRows.length - 1,
      realRows: `HH:${realHHRows.length - 1} | Cant:${realCantRows.length - 1}`
    });

  } catch (e) {
    logETL_({
      status: "ERROR",
      fileId: sourceFile && sourceFile.id ? sourceFile.id : SOURCE_EXCEL_FOLDER_ID,
      detail: (sourceFile && sourceFile.name ? "[" + sourceFile.name + "] " : "") + ((e && e.message) ? e.message : String(e)),
      planRows: "",
      realRows: ""
    });
    throw e;

  } finally {
    if (tempGsheetId) {
      try { DriveApp.getFileById(tempGsheetId).setTrashed(true); } catch(e) {}
    }
  }
}

/***************
 * Convert Excel -> Google Sheets (TEMP)
 * Requiere: Advanced Drive Service (Drive API)
 ***************/
function convertExcelToGoogleSheet_(fileId) {
  const realId = resolveShortcutId_(fileId);

  // Sin 'fields' para compatibilidad con Drive API v2 (title) y v3 (name)
  const meta = Drive.Files.get(realId, { supportsAllDrives: true });
  const fileName = (meta.name || meta.title || "ORIGEN").replace(/\.(xlsx|xls)$/i, "") + " (ETL_TEMP)";

  // Drive API v3: usa 'name'; v2 legacy usa 'title'. Soportamos ambos.
  const resource = { name: fileName, title: fileName, mimeType: MimeType.GOOGLE_SHEETS };

  const newFile = Drive.Files.copy(resource, realId, {
    convert: true,
    supportsAllDrives: true
  });

  return newFile.id;
}

/**
 * Si el ID es un Shortcut (acceso directo), devuelve el targetId real.
 */
function resolveShortcutId_(fileId) {
  // Drive API v2 (GAS) usa sintaxis de fields sin paréntesis anidados
  const f = Drive.Files.get(fileId, {
    supportsAllDrives: true,
    fields: 'id,mimeType,shortcutDetails'
  });
  const mime = (f && f.mimeType) || "";

  if (mime === "application/vnd.google-apps.shortcut") {
    const targetId = f.shortcutDetails && f.shortcutDetails.targetId;
    if (!targetId) throw new Error("El archivo es un shortcut pero no pude leer targetId.");
    return targetId;
  }

  return fileId;
}

function resolveSourceExcelFromFolder_() {
  const folderId = resolveShortcutId_(SOURCE_EXCEL_FOLDER_ID);
  const resp = Drive.Files.list({
    q: "'" + folderId + "' in parents and trashed = false",
    supportsAllDrives: true,
    includeItemsFromAllDrives: true,
    maxResults: 200
  });

  const files = resp.files || resp.items || [];
  const candidates = [];

  for (let i = 0; i < files.length; i++) {
    const candidate = buildExcelCandidate_(files[i]);
    if (candidate) candidates.push(candidate);
  }

  if (!candidates.length) {
    throw new Error("No encontre ningun archivo Excel (.xlsx/.xls) dentro de la carpeta origen.");
  }

  candidates.sort(function(a, b) {
    return new Date(b.modifiedDate).getTime() - new Date(a.modifiedDate).getTime();
  });

  return candidates[0];
}

function buildExcelCandidate_(file) {
  if (!file || !file.id) return null;

  const mime = String(file.mimeType || "");
  const isShortcut = mime === "application/vnd.google-apps.shortcut";

  if (isShortcut) {
    const targetId = file.shortcutDetails && file.shortcutDetails.targetId;
    if (!targetId) return null;

    const target = Drive.Files.get(targetId, { supportsAllDrives: true });

    if (!isExcelMimeType_(target.mimeType)) return null;

    return {
      id: target.id,
      name: target.title || "ORIGEN",
      mimeType: target.mimeType,
      modifiedDate: target.modifiedDate || file.modifiedDate || new Date(0).toISOString(),
      viaShortcutId: file.id
    };
  }

  if (!isExcelMimeType_(mime)) return null;

  return {
    id: file.id,
    name: file.title || "ORIGEN",
    mimeType: mime,
    modifiedDate: file.modifiedDate || new Date(0).toISOString(),
    viaShortcutId: null
  };
}

function isExcelMimeType_(mime) {
  return mime === "application/vnd.ms-excel" ||
    mime === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
}

/***************
 * SEMANAS -> Map FechaInicioSemana -> WEEK_KEY
 ***************/
function buildWeekMap_(shSemanas) {
  const v = shSemanas.getDataRange().getValues();
  const h = v[0].map(String);

  const iStart = h.indexOf("WEEK_START");
  const iKey   = h.indexOf("WEEK_KEY");
  if (iStart === -1 || iKey === -1) {
    throw new Error('En "SEMANAS" faltan columnas WEEK_START y/o WEEK_KEY.');
  }

  const map = {}; // "yyyy-MM-dd" -> "W202552"
  for (let r = 1; r < v.length; r++) {
    const d  = v[r][iStart];
    const wk = v[r][iKey];
    if (!d || !wk) continue;
    map[dateKey_(d)] = String(wk).trim();
  }
  return map;
}

/***************
 * PLANIFICACIÓN LB -> base (unpivot)
 * columnas fechas contienen % semanal incremental
 ***************/
function transformPlanificacionLB_(shPlanLB, weekMap) {
  const v = shPlanLB.getDataRange().getValues();
  const h = v[0].map(String);

  const iPersonal = h.indexOf("Personal");
  const iEtapa    = h.indexOf("ETAPA");
  const iAct      = h.indexOf("ACTIVIDAD");
  const iTar      = h.indexOf("TAREA");
  const iQty      = h.indexOf("Cantidad Teórica");
  const iUni      = h.indexOf("Unidad");
  const iRend     = h.indexOf("Rendimiento Teorico");
  const iHHTot    = h.indexOf("HH TOTALES");

  if ([iPersonal, iEtapa, iAct, iTar, iQty, iUni, iRend, iHHTot].some(x => x === -1)) {
    throw new Error('En "PLANIFICACIÓN LB" faltan columnas base (Personal/ETAPA/ACTIVIDAD/TAREA/Cantidad Teórica/Unidad/Rendimiento Teorico/HH TOTALES).');
  }

  // Las columnas semanales empiezan después de "Check"
  let startWeekCol = h.indexOf("Check");
  if (startWeekCol === -1) startWeekCol = iHHTot;
  startWeekCol = startWeekCol + 1;

  const out = [];
  out.push([
    "SEMANAS.WEEK_KEY",
    "Personal", "ETAPA", "ACTIVIDAD", "TAREA",
    "Cantidad Teórica", "Unidad", "Rendimiento Teorico", "HH TOTALES",
    "Avance Real", "Avance HH"
  ]);

  for (let r = 1; r < v.length; r++) {
    const personal = v[r][iPersonal];
    const etapa    = v[r][iEtapa];
    const act      = v[r][iAct];
    const tar      = v[r][iTar];
    if (!act || !tar) continue;

    const qtyTot = toNum_(v[r][iQty]);
    const rend   = toNum_(v[r][iRend]);
    const hhTot  = hhTotFix_(v[r][iHHTot]);
    const unidad = v[r][iUni];

    let wroteAny = false;
    for (let c = startWeekCol; c < h.length; c++) {
      const header = v[0][c];
      if (!header) continue;

      const wkKey = weekKeyFromHeader_(header, weekMap);
      if (!wkKey) continue;

      const pct = toNum_(v[r][c]);
      if (!pct) continue;

      const qtyWeek = qtyTot * pct;
      const hhWeek  = hhTot  * pct;

      out.push([wkKey, personal, etapa, act, tar, qtyTot, unidad, rend, hhTot, qtyWeek, hhWeek]);
      wroteAny = true;
    }
    // Fila catálogo: asegura que toda tarea con HH aparezca en la hoja
    // aunque no tenga avance semanal (todas las columnas = 0).
    // La clave "" es ignorada por sumaPorSemana_ y similares.
    if (!wroteAny && hhTot > 0) {
      out.push(["", personal, etapa, act, tar, qtyTot, unidad, rend, hhTot, 0, 0]);
    }
  }

  return out;
}

// Override robusto:
// en este reporte las columnas semanales del LB vienen en cantidades,
// no en porcentajes. Se calcula el pct de cada semana a partir de la
// cantidad planificada para derivar las HH de esa misma semana.
function transformPlanificacionLB_(shPlanLB, weekMap) {
  const v = shPlanLB.getDataRange().getValues();

  // Auto-detectar fila de headers: buscar en las primeras 5 filas
  // la que contenga ETAPA + ACTIVIDAD + TAREA
  let headerRowIdx = 0;
  for (let r = 0; r < Math.min(5, v.length); r++) {
    const row = v[r].map(String);
    if (row.indexOf("ETAPA") !== -1 && row.indexOf("ACTIVIDAD") !== -1 && row.indexOf("TAREA") !== -1) {
      headerRowIdx = r;
      break;
    }
  }

  const h = v[headerRowIdx].map(String);

  const iPersonal = h.indexOf("Personal");
  const iEtapa    = h.indexOf("ETAPA");
  const iAct      = h.indexOf("ACTIVIDAD");
  const iTar      = h.indexOf("TAREA");
  const iQty      = h.indexOf("Cantidad Teórica");
  const iUni      = h.indexOf("Unidad");
  const iRend     = h.indexOf("Rendimiento Teorico");
  const iHHTot    = h.indexOf("HH TOTALES");

  if ([iPersonal, iEtapa, iAct, iTar, iQty, iUni, iRend, iHHTot].some(x => x === -1)) {
    throw new Error('En "PLANIFICACION LB" faltan columnas base (Personal/ETAPA/ACTIVIDAD/TAREA/Cantidad Teorica/Unidad/Rendimiento Teorico/HH TOTALES). Header detectado en fila ' + (headerRowIdx + 1) + '.');
  }

  const out = [[
    "SEMANAS.WEEK_KEY",
    "Personal", "ETAPA", "ACTIVIDAD", "TAREA",
    "Cantidad Teórica", "Unidad", "Rendimiento Teorico", "HH TOTALES",
    "Avance Real", "Avance HH"
  ]];

  const weekCols = [];
  for (let c = 0; c < h.length; c++) {
    const wkKey = weekKeyFromHeader_(v[headerRowIdx][c], weekMap);
    if (wkKey) weekCols.push({ col: c, wkKey: wkKey });
  }

  if (!weekCols.length) {
    throw new Error('En "PLANIFICACION LB" no encontre columnas semanales validas (fecha/WEEK_KEY).');
  }

  for (let r = headerRowIdx + 1; r < v.length; r++) {
    const personal = v[r][iPersonal];
    const etapa    = v[r][iEtapa];
    const act      = v[r][iAct];
    const tar      = v[r][iTar];
    if (!act || !tar) continue;

    const qtyTot = toNum_(v[r][iQty]);
    const rend   = toNum_(v[r][iRend]);
    const hhTot  = hhTotFix_(v[r][iHHTot]);
    const unidad = v[r][iUni];

    let wroteAny2 = false;
    for (const weekCol of weekCols) {
      const qtyWeek = toNum_(v[r][weekCol.col]);
      if (!qtyWeek) continue;

      const pctWeek = qtyTot > 0 ? (qtyWeek / qtyTot) : 0;
      // Priorizar hhTot × pctWeek (fuente de verdad = HH TOTALES del plan).
      // Solo usar rend como fallback si hhTot no está disponible.
      const hhWeek = (hhTot > 0)
        ? (hhTot * pctWeek)
        : (rend > 0 ? qtyWeek * rend : 0);

      out.push([weekCol.wkKey, personal, etapa, act, tar, qtyTot, unidad, rend, hhTot, qtyWeek, hhWeek]);
      wroteAny2 = true;
    }
    // Fila catálogo: garantiza que toda tarea con HH figure en la hoja
    // aunque no tenga avance semanal planificado.
    if (!wroteAny2 && hhTot > 0) {
      out.push(["", personal, etapa, act, tar, qtyTot, unidad, rend, hhTot, 0, 0]);
    }
  }

  return out;
}

/***************
 * AVANCE_REAL_CANT / AVANCE_HH_REAL (matriz) -> base (unpivot)
 *
 * FIX 1: Se agrega el parámetro `hasCantTotalCol` (boolean).
 *
 * En AVANCE_REAL_CANT, la estructura es:
 *   Col A-D : PERSONAL | ETAPA | ACTIVIDAD | TAREA
 *   Col E   : "Cantidad Real" — total acumulado real de la tarea (NO es semana)
 *   Col F.. : semanas (fechas o WEEK_KEYs)
 *
 * En AVANCE_HH_REAL, la estructura es:
 *   Col A-D : Personal | ETAPA | ACTIVIDAD | TAREA
 *   Col E.. : semanas (no hay columna de total intermedia)
 *
 * Cuando hasCantTotalCol=true se detecta y saltea esa columna extra,
 * y además se exporta su valor en una columna "Cantidad Total Real"
 * para que el Code.gs pueda usarlo como denominador alternativo.
 ***************/
function transformAvanceRealMatriz_(sheet, weekMap, headerRowIndex1based, hasCantTotalCol) {
  const values = sheet.getDataRange().getValues();

  const headerRow = values[headerRowIndex1based - 1];
  const h = headerRow.map(String);

  const iPersonal = indexOfAny_(h, ["PERSONAL", "Personal"]);
  const iEtapa    = h.indexOf("ETAPA");
  const iAct      = h.indexOf("ACTIVIDAD");
  const iTar      = h.indexOf("TAREA");

  if ([iPersonal, iEtapa, iAct, iTar].some(x => x === -1)) {
    throw new Error(`En "${sheet.getName()}" no encontré PERSONAL/ETAPA/ACTIVIDAD/TAREA en fila ${headerRowIndex1based}.`);
  }

  // FIX 1A: Detectar la columna "Cantidad Real" por nombre (más robusto que offset fijo).
  // Los posibles nombres que puede tener esa columna en el Excel:
  const CANT_TOTAL_NAMES = ["Cantidad \nReal", "Cantidad Real", "CANTIDAD REAL", "Cant. Real", "CantReal"];

  let iCantTotal = -1;
  if (hasCantTotalCol) {
    for (const name of CANT_TOTAL_NAMES) {
      // Buscar en la fila de headers (fila 2 del Excel)
      const idx = h.findIndex(cell => cell.trim().replace(/\s+/g, " ") === name.trim().replace(/\s+/g, " "));
      if (idx !== -1) {
        iCantTotal = idx;
        break;
      }
    }

    // FIX 1B: Si no se encuentra por nombre, usar fallback posicional (iTar + 1)
    // y validar que ese header NO sea una fecha ni un WEEK_KEY.
    if (iCantTotal === -1) {
      const candidateIdx = iTar + 1;
      const candidateHeader = headerRow[candidateIdx];
      const isDate   = (candidateHeader instanceof Date);
      const isWeekKey = /^W\d{6}$/.test(String(candidateHeader || "").trim());
      const isNumber = (typeof candidateHeader === "number");

      if (!isDate && !isWeekKey && !isNumber) {
        // Es un encabezado de texto no reconocido → asumir que es la col de total
        iCantTotal = candidateIdx;
        Logger.log(`[WARN] "${sheet.getName()}": columna total detectada por fallback en col ${candidateIdx}: "${candidateHeader}"`);
      } else {
        // No encontramos columna de total: hasCantTotalCol puede estar mal configurado.
        // Loguear advertencia y continuar sin columna de total.
        Logger.log(`[WARN] "${sheet.getName()}": hasCantTotalCol=true pero no se encontró columna de total. Se ignorará.`);
      }
    }
  }

  // Las semanas empiezan después de la columna de total (si existe) o después de TAREA
  const startWeekCol = (iCantTotal !== -1) ? iCantTotal + 1 : iTar + 1;

  // FIX 1C: El output incluye "Cantidad Total Real" cuando aplica,
  // para que Code.gs pueda usarla como denominador de % avance real.
  const incluirCantTotal = (iCantTotal !== -1);

  const outHeaders = ["SEMANAS.WEEK_KEY", "Personal", "ETAPA", "ACTIVIDAD", "TAREA", "Valor"];
  if (incluirCantTotal) outHeaders.push("Cantidad Total Real");

  const out = [];
  out.push(outHeaders);

  for (let r = headerRowIndex1based; r < values.length; r++) {
    const row      = values[r];
    const personal = row[iPersonal];
    const etapa    = row[iEtapa];
    const act      = row[iAct];
    const tar      = row[iTar];
    if (!act || !tar) continue;

    // FIX 1D: Leer la cantidad total real de la fila (si la columna existe)
    const cantTotalReal = incluirCantTotal ? toNum_(row[iCantTotal]) : null;

    for (let c = startWeekCol; c < headerRow.length; c++) {
      const header = headerRow[c];
      if (!header) continue;

      const wkKey = weekKeyFromHeader_(header, weekMap);
      if (!wkKey) continue;

      const val = toNum_(row[c]);
      if (!val) continue;

      const outRow = [wkKey, personal, etapa, act, tar, val];
      if (incluirCantTotal) outRow.push(cantTotalReal);
      out.push(outRow);
    }
  }

  return out;
}

// Override robusto:
// toma los valores ejecutados por semana detectando columnas semanales
// por fecha/WEEK_KEY real, sin depender de offsets fijos.
function transformAvanceRealMatriz_(sheet, weekMap, headerRowIndex1based, hasCantTotalCol) {
  const values = sheet.getDataRange().getValues();

  // Auto-detectar fila de headers: buscar en las primeras 5 filas
  // la que contenga ETAPA + ACTIVIDAD + TAREA
  let headerIdx = headerRowIndex1based - 1;
  for (let r = 0; r < Math.min(5, values.length); r++) {
    const row = values[r].map(String);
    if (row.indexOf("ETAPA") !== -1 && row.indexOf("ACTIVIDAD") !== -1 && row.indexOf("TAREA") !== -1) {
      headerIdx = r;
      break;
    }
  }

  const headerRow = values[headerIdx];
  const h = headerRow.map(String);

  const iPersonal = indexOfAny_(h, ["PERSONAL", "Personal"]);
  const iEtapa    = h.indexOf("ETAPA");
  const iAct      = h.indexOf("ACTIVIDAD");
  const iTar      = h.indexOf("TAREA");

  if ([iPersonal, iEtapa, iAct, iTar].some(x => x === -1)) {
    throw new Error(`En "${sheet.getName()}" no encontre PERSONAL/ETAPA/ACTIVIDAD/TAREA en fila ${headerIdx + 1}.`);
  }

  const CANT_TOTAL_NAMES = [
    "Cantidad \nReal",
    "Cantidad Real",
    "CANTIDAD REAL",
    "Cant. Real",
    "CantReal",
    "Cantidad Total",
    "CANTIDAD TOTAL",
    "Cant. Total",
    "CantTotal"
  ];

  let iCantTotal = -1;
  if (hasCantTotalCol) {
    for (const name of CANT_TOTAL_NAMES) {
      const idx = h.findIndex(cell => cell.trim().replace(/\s+/g, " ") === name.trim().replace(/\s+/g, " "));
      if (idx !== -1) {
        iCantTotal = idx;
        break;
      }
    }

    if (iCantTotal === -1) {
      Logger.log(`[WARN] "${sheet.getName()}": hasCantTotalCol=true pero no se encontro columna de total conocida. Se ignorara.`);
    }
  }

  const weekCols = [];
  for (let c = 0; c < headerRow.length; c++) {
    const wkKey = weekKeyFromHeader_(headerRow[c], weekMap);
    if (wkKey) weekCols.push({ col: c, wkKey: wkKey });
  }

  if (!weekCols.length) {
    throw new Error(`En "${sheet.getName()}" no encontre columnas semanales validas (fecha/WEEK_KEY).`);
  }

  const incluirCantTotal = (iCantTotal !== -1);
  const outHeaders = ["SEMANAS.WEEK_KEY", "Personal", "ETAPA", "ACTIVIDAD", "TAREA", "Valor"];
  if (incluirCantTotal) outHeaders.push("Cantidad Total Real");

  const out = [outHeaders];

  for (let r = headerIdx + 1; r < values.length; r++) {
    const row      = values[r];
    const personal = row[iPersonal];
    const etapa    = row[iEtapa];
    const act      = row[iAct];
    const tar      = row[iTar];
    if (!act || !tar) continue;

    const cantTotalReal = incluirCantTotal ? toNum_(row[iCantTotal]) : null;

    for (const weekCol of weekCols) {
      const val = toNum_(row[weekCol.col]);
      if (!val) continue;

      const outRow = [weekCol.wkKey, personal, etapa, act, tar, val];
      if (incluirCantTotal) outRow.push(cantTotalReal);
      out.push(outRow);
    }
  }

  return out;
}

/***************
 * Helpers
 ***************/
function writeTable_(sheet, rows) {
  if (!sheet) throw new Error("No existe hoja destino (revisa nombres).");
  sheet.clearContents();
  if (!rows || rows.length === 0) return;

  const numRows = rows.length;
  const numCols = rows[0].length;
  const range = sheet.getRange(1, 1, numRows, numCols);

  // FIX CRÍTICO: resetear formato a número plano ANTES de escribir.
  // Evita que Sheets reinterprete números como fechas por formato residual.
  range.setNumberFormat("0.##########");
  // Los headers (fila 1) van como texto plano
  sheet.getRange(1, 1, 1, numCols).setNumberFormat("@");

  range.setValues(rows);
}

function dateKey_(d) {
  const tz = Session.getScriptTimeZone();
  return Utilities.formatDate(new Date(d), tz, "yyyy-MM-dd");
}

function weekKeyFromHeader_(header, weekMap) {
  if (header instanceof Date) {
    return weekMap[dateKey_(header)] || null;
  }
  const asDate = new Date(header);
  if (!isNaN(asDate.getTime())) {
    return weekMap[dateKey_(asDate)] || null;
  }
  const s = String(header).trim();
  if (/^W\d{6}$/.test(s)) return s;
  return null;
}

function indexOfAny_(arr, candidates) {
  for (const c of candidates) {
    const idx = arr.indexOf(c);
    if (idx !== -1) return idx;
  }
  return -1;
}

function toNum_(x) {
  if (x === null || x === undefined || x === "") return 0;
  if (typeof x === "number") return isFinite(x) ? x : 0;
  let s = String(x).trim();
  if (!s) return 0;
  s = s.replace(/\s+/g, "");
  if (s.includes(",") && s.includes(".")) {
    s = s.replace(/\./g, "").replace(",", ".");
  } else if (s.includes(",") && !s.includes(".")) {
    s = s.replace(",", ".");
  }
  const n = Number(s);
  return isFinite(n) ? n : 0;
}

function hhTotFix_(x) {
  // FIX CRÍTICO: cuando Excel tiene un número en la celda HH TOTALES y Google
  // Sheets lo convierte al abrir el xlsx, a veces lo interpreta como fecha serial.
  // Ej: 1272 HH → se guarda como serial de fecha → getValues() devuelve Date(1903-06-25)
  // El número de días desde 1899-12-30 (epoch de Excel) ES el valor original de HH.
  if (x instanceof Date) {
    // Usamos UTC para evitar desfases por timezone del script de Apps Script.
    // Epoch de Excel en UTC: 30 de diciembre de 1899
    const EXCEL_EPOCH_MS = Date.UTC(1899, 11, 30); // mes 11 = diciembre
    const diffMs   = x.getTime() - EXCEL_EPOCH_MS;
    const diffDays = Math.round(diffMs / (1000 * 60 * 60 * 24));
    // diffDays es el número serial de Excel → valor original de HH
    if (diffDays > 0 && diffDays < 1e6) return diffDays;
    return 0;
  }
  if (typeof x === "number") {
    if (!isFinite(x) || x < 0 || x > 1e9) return 0;
    return x;
  }
  let s = String(x || "").trim();
  if (!s) return 0;
  s = s.replace(/\s+/g, "");
  if (s.includes(",") && s.includes(".")) s = s.replace(/\./g, "").replace(",", ".");
  else if (s.includes(",") && !s.includes(".")) s = s.replace(",", ".");
  const n = Number(s);
  if (!isFinite(n) || n < 0 || n > 1e9) return 0;
  return n;
}

/***************
 * TRIGGERS
 ***************/
function crearTriggerETLNocturno() {
  borrarTriggerETLNocturno();

  ScriptApp.newTrigger('ETL_actualizarTodo')
    .timeBased()
    .atHour(23)
    .everyDays(1)
    .create();

  return 'Trigger nocturno ETL creado: corre todos los días alrededor de las 23:00.';
}

function borrarTriggerETLNocturno() {
  var triggers = ScriptApp.getProjectTriggers();
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'ETL_actualizarTodo') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  return 'Trigger ETL eliminado.';
}

function verTriggerETLNocturno() {
  var triggers = ScriptApp.getProjectTriggers();
  var out = [];
  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'ETL_actualizarTodo') {
      out.push(triggers[i].getHandlerFunction() + ' | ' + triggers[i].getEventType());
    }
  }
  Logger.log(out.length ? out.join('\n') : 'No hay trigger ETL configurado.');
  return out;
}

/***************
 * LOG
 ***************/
const ETL_LOG_SHEET = "ETL_LOG";

function logETL_(data) {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(ETL_LOG_SHEET);
  if (!sh) {
    sh = ss.insertSheet(ETL_LOG_SHEET);
    sh.getRange(1, 1, 1, 6).setValues([["Timestamp", "Status", "OrigenFileId", "Detalle", "PlanRows", "RealRows(HH/Cant)"]]);
  }
  sh.appendRow([
    new Date(),
    data.status   || "",
    data.fileId   || "",
    data.detail   || "",
    data.planRows || "",
    data.realRows || ""
  ]);
}

/***************
 * DEBUG / TEST
 ***************/
function TEST_driveApiOrigen() {
  const selected = resolveSourceExcelFromFolder_();
  Logger.log("Folder ID: " + SOURCE_EXCEL_FOLDER_ID);
  Logger.log("Archivo seleccionado ID: " + selected.id);
  Logger.log("Archivo seleccionado: " + selected.name);
  Logger.log("mimeType: " + selected.mimeType);
  Logger.log("modifiedDate: " + selected.modifiedDate);
  if (selected.viaShortcutId) {
    Logger.log("Seleccionado a traves de shortcut: " + selected.viaShortcutId);
  }
}

function DEBUG_headers_PLAN_LB() {
  const source = resolveSourceExcelFromFolder_();
  const tempId = convertExcelToGoogleSheet_(source.id);
  try {
    const ss = SpreadsheetApp.openById(tempId);
    const sh = ss.getSheetByName(SRC_SHEET_PLAN_LB);
    const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
    Logger.log(headers.map((x, i) => `${i}: ${x}`).join(" | "));
    Logger.log("Ejemplo fila 2 (valores 0-15): " + sh.getRange(2, 1, 1, Math.min(16, sh.getLastColumn())).getValues()[0].join(" | "));
  } finally {
    try { DriveApp.getFileById(tempId).setTrashed(true); } catch (e) {}
  }
}

/**
 * DEBUG: Muestra los primeros headers y filas de AVANCE_REAL_CANT
 * para verificar que la detección de columna de total es correcta.
 */
function DEBUG_headers_AVANCE_REAL_CANT() {
  const source = resolveSourceExcelFromFolder_();
  const tempId = convertExcelToGoogleSheet_(source.id);
  try {
    const ss = SpreadsheetApp.openById(tempId);
    const sh = ss.getSheetByName(SRC_SHEET_REAL_CANT);
    const maxCol = Math.min(10, sh.getLastColumn());

    Logger.log("=== AVANCE_REAL_CANT ===");
    Logger.log("Fila 1 (row 0): " + sh.getRange(1, 1, 1, maxCol).getValues()[0].join(" | "));
    Logger.log("Fila 2 headers: " + sh.getRange(2, 1, 1, maxCol).getValues()[0].join(" | "));
    Logger.log("Fila 3 datos  : " + sh.getRange(3, 1, 1, maxCol).getValues()[0].join(" | "));
  } finally {
    try { DriveApp.getFileById(tempId).setTrashed(true); } catch (e) {}
  }
}
