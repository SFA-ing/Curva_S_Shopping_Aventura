// ============================================================
//  Code.gs  — Reporte Avance Físico
//  v2 — Fixes aplicados:
//    FIX 2: earnedHHReal sin cap (permite sobreejcución > 100%)
//            pero llevando acumulado por tarea para no sumar
//            el mismo incremento dos veces.
//    FIX 3: rendimiento real = HH acum / Qty acum al corte
//            (no ratio semanal), con fallback al teórico si
//            el acumulado real no es válido.
//    FIX 4: normalización de claves act||tar para evitar
//            mismatches por espacios / tildes / mayúsculas.
//    DENOMINADOR: siempre cantidad teórica del LB (qtyPlan).
// ============================================================

function doGet() {
  return HtmlService.createHtmlOutputFromFile("Index")
    .setTitle("Shopping Aventura - Reporte Avance Fisico")
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// ============================================================
//  EXPORTACIÓN A HOJA CACHE
// ============================================================

var CACHE_SHEET_ID = '1ijIAkVvTYYkgcKOTH8RNpgox6bN4-fvNcROAo84CFw0';
var CACHE_SHEET_NAME = 'data';

// CAMBIAR EN CADA REPORTE
var DASHBOARD_KEY = 'aventura_shopping';
var REPORTE_FISICO_URL = 'https://script.google.com/a/macros/ingener.com/s/AKfycbxK8UUzaeZ4H3ke1_ZzKRXzyzkm88a5-MkZkRmVqkgAOJyJvEkVQhl2Lj-caDdPI3tQ5A/exec';

function actualizarManualmente() {
  ETL_actualizarTodo();
  exportarParaDashboard();
  return 'Actualización manual completada: ' + new Date().toLocaleString();
}

function exportarParaDashboard() {
  var data = getDashboardData();
  var json = JSON.stringify(data);

  var ss = SpreadsheetApp.openById(CACHE_SHEET_ID);
  var sh = ss.getSheetByName(CACHE_SHEET_NAME);

  if (!sh) {
    throw new Error('No existe la hoja cache: ' + CACHE_SHEET_NAME);
  }

  var values = sh.getDataRange().getValues();
  var foundRow = null;

  for (var i = 1; i < values.length; i++) {
    var key = sanitizeDashboardKey_(values[i][0] || '');
    if (key === sanitizeDashboardKey_(DASHBOARD_KEY)) {
      foundRow = i + 1;
      break;
    }
  }

  if (!foundRow) {
    foundRow = Math.max(2, sh.getLastRow() + 1);
  }

  sh.getRange(foundRow, 1).setValue(sanitizeDashboardKey_(DASHBOARD_KEY));
  sh.getRange(foundRow, 2).setValue(json);
  sh.getRange(foundRow, 3).setValue(new Date());

  Logger.log('Exportación OK: ' + DASHBOARD_KEY + ' @ ' + new Date());
  return 'Exportado correctamente con clave: ' + sanitizeDashboardKey_(DASHBOARD_KEY);
}

function crearTriggerExportacionNocturna() {
  borrarTriggersExportacionNocturna();

  ScriptApp.newTrigger('exportarParaDashboard')
    .timeBased()
    .atHour(0)
    .everyDays(1)
    .create();

  return 'Trigger nocturno creado para exportar una vez por día alrededor de las 00:00.';
}

function borrarTriggersExportacionNocturna() {
  var triggers = ScriptApp.getProjectTriggers();

  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'exportarParaDashboard') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }

  return 'Triggers de exportación eliminados.';
}

function verTriggersExportacionNocturna() {
  var triggers = ScriptApp.getProjectTriggers();
  var out = [];

  for (var i = 0; i < triggers.length; i++) {
    out.push(triggers[i].getHandlerFunction() + ' | ' + triggers[i].getEventType());
  }

  Logger.log(out.join('\n'));
  return out;
}

function sanitizeDashboardKey_(key) {
  return String(key || 'fisico_default')
    .trim()
    .toLowerCase()
    .replace(/[^a-z0-9_-]/g, '_')
    .replace(/_+/g, '_')
    .replace(/^_+|_+$/g, '') || 'fisico_default';
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Dashboard')
    .addItem('Exportar ahora', 'exportarParaDashboard')
    .addSeparator()
    .addItem('ETL: correr ahora', 'ETL_actualizarTodo')
    .addItem('ETL: crear trigger nocturno 23:00', 'crearTriggerETLNocturno')
    .addItem('ETL: ver trigger', 'verTriggerETLNocturno')
    .addItem('ETL: borrar trigger', 'borrarTriggerETLNocturno')
    .addSeparator()
    .addItem('Exportación: crear trigger nocturno 00:00', 'crearTriggerExportacionNocturna')
    .addItem('Exportación: ver triggers', 'verTriggersExportacionNocturna')
    .addItem('Exportación: borrar triggers', 'borrarTriggersExportacionNocturna')
    .addToUi();
}


// ------------------------------------------------------------
//  ENTRY POINT PRINCIPAL
// ------------------------------------------------------------
function getDashboardData() {
  const ss = SpreadsheetApp.getActive();

  const shPlan = ss.getSheetByName("Planificación Inicial");
  const shHH   = ss.getSheetByName("AVANCE_HH_REAL (2)");
  const shQty  = ss.getSheetByName("AVANCE_REAL_CANT (2)");

  if (!shPlan || !shHH || !shQty) {
    throw new Error('Faltan hojas. Deben existir: "Planificación Inicial", "AVANCE_HH_REAL (2)", "AVANCE_REAL_CANT (2)".');
  }

  // Corte real = lunes de la semana pasada
  const today        = new Date();
  const cutMonday    = startOfWeekMonday_(addDays_(today, -7));
  const cutWeekKey   = dateToWeekKey_(cutMonday);
  const prevCutMonday = addDays_(cutMonday, -7);

  // Meta tareas (denominadores teóricos del LB)
  const taskMeta = construirTaskMeta_(shPlan);
  const denom    = construirDenominadores_(taskMeta);

  if (denom.globalHH <= 0) {
    const stats = diagnosticoDenominador_(taskMeta);
    throw new Error(
      "Denominador HH global inválido. " +
      `Tareas=${stats.total} | conQty=${stats.conQty} | conHH=${stats.conHH} | conAmbos=${stats.conAmbos} | ` +
      `ejemploSinHH=${stats.ejemploSinHH} | ejemploSinQty=${stats.ejemploSinQty}`
    );
  }

  const hhPlanTotal = denom.globalHH;

  // Series globales HH plan/real por semana
  const hhPlanSem = sumaPorSemana_(shPlan, "SEMANAS.WEEK_KEY", "Avance HH");
  const hhRealSem = sumaPorSemana_(shHH,   "SEMANAS.WEEK_KEY", "Valor");

  // Series qty plan/real por semana y tarea
  const qtyPlanByWeekTask = cantidadPorSemanaPorTarea_(shPlan, "SEMANAS.WEEK_KEY", "ACTIVIDAD", "TAREA", "Avance Real");
  const qtyRealByWeekTask = cantidadPorSemanaPorTarea_(shQty,  "SEMANAS.WEEK_KEY", "ACTIVIDAD", "TAREA", "Valor");

  // Universo de semanas
  const weekSet = new Set([
    ...Object.keys(hhPlanSem),
    ...Object.keys(hhRealSem),
    ...Object.keys(qtyPlanByWeekTask),
    ...Object.keys(qtyRealByWeekTask),
  ]);
  const weeks = Array.from(weekSet).sort((a, b) => weekKeyToDate_(a) - weekKeyToDate_(b));
  if (weeks.length === 0) throw new Error("No encontré semanas para construir el reporte.");

  // ----------------------------------------------------------
  //  Curvas acumuladas
  //  FIX 2: earnedHHReal NO tiene cap individual por tarea,
  //  pero SÍ llevamos acumulado por tarea para que cada unidad
  //  real sólo se cuente una vez (evitar doble conteo semanal).
  //  El exceso > qtyPlan se refleja directamente (sobreejcución).
  // ----------------------------------------------------------
  let hhPlanAc    = 0;
  let hhRealAc    = 0;
  let earnedHHPlan = 0;  // PV ponderado HH
  let earnedHHReal = 0;  // EV ponderado HH

  // Acumulado de cantidad por tarea (para no recutar lo mismo semana a semana)
  const qtyPlanAcumByTask = {}; // key -> qty acumulada plan
  const qtyRealAcumByTask = {}; // key -> qty acumulada real

  const curveMap = {}; // wk -> { hhPlanAc, hhRealAc|null, pctPlan, pctReal|null }
  const rendimientoSemanalMap = {}; // wk -> factor real semanal vs plan (1 = igual al plan)

  for (const wk of weeks) {
    const d      = weekKeyToDate_(wk);
    const inReal = (d <= cutMonday);
    const hhRealWeek = inReal ? Number(hhRealSem[wk] || 0) : null;
    let earnedHHRealWeek = 0;

    // HH acumuladas calendario
    hhPlanAc += Number(hhPlanSem[wk] || 0);
    if (inReal) hhRealAc += Number(hhRealSem[wk] || 0);

    // --- Plan: incremento qty por tarea → earned HH plan ---
    const qpw = qtyPlanByWeekTask[wk] || {};
    for (const [k, v] of Object.entries(qpw)) {
      const m = taskMeta[k];
      if (!m || !(m.qtyPlan > 0) || !(m.hhPlan > 0)) continue;
      const qtyInc = Number(v || 0);
      const pctInc = qtyInc / m.qtyPlan;         // sin cap: plan tampoco debería exceder
      earnedHHPlan += pctInc * m.hhPlan;
      qtyPlanAcumByTask[k] = (qtyPlanAcumByTask[k] || 0) + qtyInc;
    }

    // --- Real: incremento qty por tarea → earned HH real ---
    // FIX 2: Sin cap. Si la cantidad real supera la teórica,
    // el porcentaje superará 1 y eso se refleja en la curva.
    if (inReal) {
      const qrw = qtyRealByWeekTask[wk] || {};
      for (const [k, v] of Object.entries(qrw)) {
        const m = taskMeta[k];
        if (!m || !(m.qtyPlan > 0) || !(m.hhPlan > 0)) continue;
        const qtyInc = Number(v || 0);
        const pctInc = qtyInc / m.qtyPlan;       // FIX 2: sin Math.min(1, ...)
        const earnedInc = pctInc * m.hhPlan;
        earnedHHReal += earnedInc;
        earnedHHRealWeek += earnedInc;
        qtyRealAcumByTask[k] = (qtyRealAcumByTask[k] || 0) + qtyInc;
      }
    }

    const pctPlan = Math.min(1, earnedHHPlan / denom.globalHH);
    const pctReal = earnedHHReal / denom.globalHH;

    curveMap[wk] = {
      hhPlanAc,
      hhRealAc : inReal ? hhRealAc : null,
      pctPlan,
      pctReal  : inReal ? pctReal  : null
    };

    rendimientoSemanalMap[wk] = {
      hhPor1Pct: (inReal && earnedHHRealWeek > 0 && hhRealWeek != null && hhRealWeek > 0)
        ? Math.round(hhRealWeek / (earnedHHRealWeek / hhPlanTotal * 100))
        : null
    };
  }

  const cutVals  = encontrarUltimoHasta_(curveMap, weeks, cutMonday);
  const prevVals = encontrarUltimoHasta_(curveMap, weeks, prevCutMonday);

  const hhRealCut  = cutVals.hhRealAc  ?? 0;
  const hhPlanCut  = cutVals.hhPlanAc  ?? 0;
  const pctRealCut = cutVals.pctReal   ?? 0;
  const pctPlanCut = cutVals.pctPlan   ?? 0;

  const pctRealPrev = prevVals.pctReal ?? 0;
  const pctPlanPrev = prevVals.pctPlan ?? 0;

  const hhPrevPorAvance    = interpolarHHPlanPorPct_(curveMap, weeks, pctRealCut);
  const diffHH_vsPrevAvance = hhRealCut - hhPrevPorAvance;
  const diffPctAcum         = pctRealCut - pctPlanCut;

  const ultAvReal = pctRealCut - pctRealPrev;
  const ultAvPlan = pctPlanCut - pctPlanPrev;
  const diffUltAv = ultAvReal - ultAvPlan;

  // ----------------------------------------------------------
  //  Corte por tarea (real acumulado al corte)
  //  FIX 3: rendimiento = HH acum / Qty acum (totales al corte)
  // ----------------------------------------------------------
  const byTaskCut = construirCortePorTarea_(shHH, shQty, taskMeta, cutMonday);

  // Pendientes
  const hhPendTeorico = Math.max(0, hhPlanTotal - hhRealCut);
  const hhPendRend    = calcularHHRestantesPorRendActual_(taskMeta, byTaskCut);
  const diffPend      = hhPendRend - hhPendTeorico;

  // HH por 1%
  const hhPor1PctTeor = hhPlanTotal / 100;
  const hhPor1PctReal = (pctRealCut > 0) ? (hhRealCut / (pctRealCut * 100)) : null;

  // Series para charts
  const curveSeries = [["WEEK_KEY", "HH Plan Acum", "HH Real Acum", "% Físico Plan", "% Físico Real"]];
  for (const wk of weeks) {
    const v = curveMap[wk];
    curveSeries.push([wk, v.hhPlanAc, v.hhRealAc, v.pctPlan, v.pctReal]);
  }

  const last6       = ultimasNSemanasHasta_(weeks, cutMonday, 6);
  const trendSeries = [["WEEK_KEY", "HH por 1% avance"]];
  for (const wk of last6) {
    const v = rendimientoSemanalMap[wk] || {};
    trendSeries.push([wk, v.hhPor1Pct ?? null]);
  }

  // ----------------------------------------------------------
  //  PROYECCIONES
  //  Se calculan a partir del corte usando el rendimiento real
  //  acumulado (earnedHHReal / hhRealCut = pctReal / HH reales).
  //
  //  Escenario A — Mismo ritmo de HH: ¿en qué semana termino?
  //    → Cada semana futura se proyecta con el mismo avance semanal
  //      promedio real de las últimas N semanas.
  //    → La proyección de HH sigue igual (misma cadencia).
  //
  //  Escenario B — Misma fecha de término: ¿cuántas HH necesito?
  //    → Se distribuye el avance faltante en las semanas restantes
  //      del plan original.
  //    → Las HH se escalan proporcionalmente al avance adicional
  //      requerido por semana usando el rendimiento real.
  // ----------------------------------------------------------
  const proyecciones = calcularProyecciones_(
    curveMap, weeks, cutMonday,
    pctRealCut, hhRealCut,
    denom.globalHH, hhPendRend,
    taskMeta, byTaskCut
  );

  // Tablas
  const stageSummary    = construirResumenPorEtapa_(shPlan, shHH, taskMeta, denom, byTaskCut, cutMonday);
  const activitySummary = construirResumenPorActividad_(shPlan, shHH, taskMeta, denom, byTaskCut, cutMonday);
  const taskTable       = construirTablaTareasV2_(shPlan, cutMonday, taskMeta, byTaskCut);

  const kpis = {
    corte: { weekKey: cutWeekKey, mondayISO: formatISODate_(cutMonday) },
    hhReales          : hhRealCut,
    hhPrevPorAvance,
    diffRealVsPrev    : diffHH_vsPrevAvance,
    pctRealAcum       : pctRealCut,
    pctPlanAcum       : pctPlanCut,
    diffPctAcum,
    ultAvReal,
    ultAvPlan,
    diffUltAv,
    hhPendTeorico,
    hhPendRend,
    diffPend,
    hhPor1PctTeor,
    hhPor1PctReal
  };

  return {
  meta: {
    generatedAtISO : formatISODateTime_(new Date()),
    corteWeekKey   : cutWeekKey,
    corteMondayISO : formatISODate_(cutMonday),
    reportUrl      : REPORTE_FISICO_URL
  },
  kpis,
  curveSeries,
  trendSeries,
  proyecciones,
  stageSummary,
  activitySummary,
  taskTable
  };
}

// ============================================================
//  RESUMEN POR ETAPA (al corte)
// ============================================================
function construirResumenPorEtapa_(shPlan, shHH, taskMeta, denom, byTaskCut, cutMonday) {
  const hhPlanByWeekStage = sumaPorSemanaGrupo_(shPlan, "SEMANAS.WEEK_KEY", "ETAPA", "Avance HH");
  const hhRealByWeekStage = sumaPorSemanaGrupo_(shHH,   "SEMANAS.WEEK_KEY", "ETAPA", "Valor");

  const qtyPlanCutByTask = sumaHastaCortePorTarea_(shPlan, cutMonday, "Avance Real");

  const qtyRealCutByTask = {};
  for (const [k, v] of Object.entries(byTaskCut)) {
    qtyRealCutByTask[k] = Number(v.qtyReal || 0);
  }

  const stages = Object.keys(denom.byEtapa || {}).sort((a, b) => a.localeCompare(b));

  const rows = [];
  rows.push(["ETAPA", "m (corte / total)", "Rend. Plan (HH/m)", "Rend. Real (HH/m)", "HH Plan (corte)", "HH Real (corte)", "Δ HH", "% Fisico Plan", "% Fisico Real", "Δ % (pp)"]);

  for (const etapa of stages) {
    const hhP = hhAcumGrupoHastaCorte_(hhPlanByWeekStage, etapa, cutMonday);
    const hhR = hhAcumGrupoHastaCorte_(hhRealByWeekStage, etapa, cutMonday);

    const denomEtapa   = denom.byEtapa[etapa] || 0;
    const pctPlanEtapa = denomEtapa > 0 ? pctFisicoGrupoDesdeQty_(taskMeta, "etapa", etapa, denomEtapa, qtyPlanCutByTask) : null;
    const pctRealEtapa = denomEtapa > 0 ? pctFisicoGrupoDesdeQty_(taskMeta, "etapa", etapa, denomEtapa, qtyRealCutByTask) : null;

    // Metros (unidad === "m") y rendimientos para esta etapa
    let mTotal  = 0;
    let mReal   = 0;
    let hhPlanM = 0;
    let hhRealM = 0;
    for (const [k, m] of Object.entries(taskMeta)) {
      if (m.etapa !== etapa || String(m.unidad || "").trim().toLowerCase() !== "m") continue;
      mTotal  += Number(m.qtyPlan || 0);
      mReal   += Number(qtyRealCutByTask[k] || 0);
      hhPlanM += Number(m.hhPlan || 0);
      hhRealM += Number((byTaskCut[k] && byTaskCut[k].hhReal) || 0);
    }
    const metrosCell = mTotal  > 0 ? { r: mReal, t: mTotal } : null;
    const rendPlan   = mTotal  > 0 ? hhPlanM / mTotal        : null;
    const rendReal   = mReal   > 0 ? hhRealM / mReal         : null;

    rows.push([
      etapa,
      metrosCell,
      rendPlan,
      rendReal,
      hhP,
      hhR,
      hhR - hhP,
      pctPlanEtapa,
      pctRealEtapa,
      (pctPlanEtapa != null && pctRealEtapa != null) ? (pctRealEtapa - pctPlanEtapa) : null
    ]);
  }

  return rows;
}

// ============================================================
//  RESUMEN POR ACTIVIDAD (al corte)
// ============================================================
function construirResumenPorActividad_(shPlan, shHH, taskMeta, denom, byTaskCut, cutMonday) {
  const hhPlanByWeekAct = sumaPorSemanaActividad_(shPlan, "SEMANAS.WEEK_KEY", "ACTIVIDAD", "Avance HH");
  const hhRealByWeekAct = sumaPorSemanaActividad_(shHH,   "SEMANAS.WEEK_KEY", "ACTIVIDAD", "Valor");

  const qtyPlanCutByTask = sumaHastaCortePorTarea_(shPlan, cutMonday, "Avance Real");

  // FIX 4: construir qtyRealCutByTask desde byTaskCut (que ya usa claves normalizadas)
  const qtyRealCutByTask = {};
  for (const [k, v] of Object.entries(byTaskCut)) {
    qtyRealCutByTask[k] = Number(v.qtyReal || 0);
  }

  const acts = Object.keys(denom.byAct).sort((a, b) => a.localeCompare(b));

  const rows = [];
  rows.push(["ACTIVIDAD", "HH Plan (corte)", "HH Real (corte)", "Δ HH", "% Físico Plan", "% Físico Real", "Δ % (pp)"]);

  for (const act of acts) {
    const hhP = hhAcumActHastaCorte_(hhPlanByWeekAct, act, cutMonday);
    const hhR = hhAcumActHastaCorte_(hhRealByWeekAct, act, cutMonday);

    const denomAct   = denom.byAct[act] || 0;
    const pctPlanAct = denomAct > 0 ? pctFisicoActDesdeQty_(taskMeta, act, denomAct, qtyPlanCutByTask) : null;
    const pctRealAct = denomAct > 0 ? pctFisicoActDesdeQty_(taskMeta, act, denomAct, qtyRealCutByTask) : null;

    rows.push([
      act,
      hhP,
      hhR,
      hhR - hhP,
      pctPlanAct,
      pctRealAct,
      (pctPlanAct != null && pctRealAct != null) ? (pctRealAct - pctPlanAct) : null
    ]);
  }
  return rows;
}

// ============================================================
//  TABLA POR TAREA (V2)
// ============================================================
function construirTablaTareasV2_(shPlan, cutMonday, taskMeta, byTaskCut) {
  const hhPlanCutByTask  = sumaHastaCortePorTarea_(shPlan, cutMonday, "Avance HH");
  const qtyPlanCutByTask = sumaHastaCortePorTarea_(shPlan, cutMonday, "Avance Real");

  const rows = [];
  rows.push([
    "ACTIVIDAD", "TAREA",
    "Cant. Total", "Unidad",
    "Cant. Real (corte)", "% Real (corte)",
    "% Plan (corte)",
    "HH Real (corte)", "HH Plan (corte)",
    "Rend Plan (HH/u)", "Rend Real (HH/u)",
    "HH Rest. (rend act.)"
  ]);

  const keys = Object.keys(taskMeta).sort((a, b) => {
    const A = taskMeta[a], B = taskMeta[b];
    if (A.actividad !== B.actividad) return A.actividad.localeCompare(B.actividad);
    return A.tarea.localeCompare(B.tarea);
  });

  for (const k of keys) {
    const m = taskMeta[k];

    const hhPlanCut  = Number(hhPlanCutByTask[k]  || 0);
    const real       = byTaskCut[k] || { hhReal: 0, qtyReal: 0, rendReal: null };
    const hhRealCut  = Number(real.hhReal  || 0);

    const qtyPlanTotal = Number(m.qtyPlan   || 0);
    const qtyPlanCut   = Number(qtyPlanCutByTask[k] || 0);
    const qtyRealCut   = Number(real.qtyReal || 0);

    // DENOMINADOR: siempre qtyPlan (teórico LB)
    // FIX 2: pctReal puede superar 1 si hay sobreejcución — sin cap
    const pctPlan = (qtyPlanTotal > 0) ? Math.max(0, qtyPlanCut  / qtyPlanTotal) : null;
    const pctReal = (qtyPlanTotal > 0) ? Math.max(0, qtyRealCut  / qtyPlanTotal) : null;
    const diffPct = (pctPlan != null && pctReal != null) ? (pctReal - pctPlan) : null;

    const diffHH  = hhRealCut - hhPlanCut;

    // FIX 3: rendimiento real = HH acum / Qty acum al corte (totales, no semanal)
    // rendReal ya viene calculado así desde construirCortePorTarea_ (ver abajo)
    let rend = real.rendReal;
    if (rend == null || !isFinite(rend) || rend <= 0) rend = m.rendPlan; // fallback al teórico

    const qtyRem = Math.max(0, qtyPlanTotal - qtyRealCut);
    const hhRest = (rend != null && isFinite(rend) && rend > 0) ? qtyRem * rend : null;

    rows.push([
      m.actividad,
      m.tarea,
      qtyPlanTotal,             // Cant. Total
      m.unidad || "",           // Unidad
      qtyRealCut,               // Cant. Real (corte)
      pctReal,                  // % Real (corte)
      pctPlan,                  // % Plan (corte)
      hhRealCut,                // HH Real (corte)
      hhPlanCut,                // HH Plan (corte)
      m.rendPlan,               // Rend Plan
      real.rendReal,            // Rend Real
      hhRest                    // HH Rest.
    ]);
  }
  return rows;
}

// ============================================================
//  META / PONDERACIONES
// ============================================================
function construirTaskMeta_(shPlan) {
  const v = shPlan.getDataRange().getValues();
  const h = v[0].map(String);

  const iEta  = h.indexOf("ETAPA");
  const iAct  = h.indexOf("ACTIVIDAD");
  const iTar  = h.indexOf("TAREA");
  const iQty  = h.indexOf("Cantidad Teórica");
  const iHH   = h.indexOf("HH TOTALES");
  const iRend = h.indexOf("Rendimiento Teorico");
  const iUni  = h.indexOf("Unidad");   // nueva: unidad de medida

  if (iEta === -1 || iAct === -1 || iTar === -1 || iQty === -1 || iHH === -1) {
    throw new Error('En "Planificación Inicial" faltan: ACTIVIDAD, TAREA, Cantidad Teórica, HH TOTALES.');
  }

  const map = {};
  for (let r = 1; r < v.length; r++) {
    // FIX 4: normalizar clave para evitar mismatches por espacios/tildes/mayúsculas
    const eta = normalizeKey_(v[r][iEta]);
    const act = normalizeKey_(v[r][iAct]);
    const tar = normalizeKey_(v[r][iTar]);
    if (!eta || !act || !tar) continue;

    const key    = act + "||" + tar;
    const qtyPlan = toNumBack_(v[r][iQty]);
    let   hhPlan  = toNumBack_(v[r][iHH]);

    if (!(hhPlan > 0) || hhPlan > 1e9) {
      const rend = (iRend !== -1) ? toNumBack_(v[r][iRend]) : 0;
      if (qtyPlan > 0 && rend > 0) hhPlan = qtyPlan * rend;
    }

    const unidad = (iUni !== -1) ? String(v[r][iUni] || "").trim() : "";

    if (!map[key]) {
      map[key] = {
        etapa    : eta,
        actividad: act,
        tarea    : tar,
        unidad,
        qtyPlan,
        hhPlan,
        rendPlan : (qtyPlan > 0) ? (hhPlan / qtyPlan) : null
      };
    } else {
      if (!map[key].etapa && eta) map[key].etapa = eta;
      map[key].qtyPlan = Math.max(map[key].qtyPlan || 0, qtyPlan || 0);
      map[key].hhPlan  = Math.max(map[key].hhPlan  || 0, hhPlan  || 0);
      map[key].rendPlan = (map[key].qtyPlan > 0) ? (map[key].hhPlan / map[key].qtyPlan) : null;
      if (!map[key].unidad && unidad) map[key].unidad = unidad;
    }
  }
  return map;
}

function construirDenominadores_(taskMeta) {
  let globalHH = 0;
  const byEtapa = {};
  const byAct  = {};
  for (const m of Object.values(taskMeta)) {
    if ((m.qtyPlan || 0) > 0 && (m.hhPlan || 0) > 0) {
      globalHH += m.hhPlan;
      byEtapa[m.etapa] = (byEtapa[m.etapa] || 0) + m.hhPlan;
      byAct[m.actividad] = (byAct[m.actividad] || 0) + m.hhPlan;
    }
  }
  return { globalHH, byEtapa, byAct };
}

function pctFisicoGrupoDesdeQty_(taskMeta, propName, groupValue, denomGroupHH, cumQtyMap) {
  let num = 0;
  for (const [k, m] of Object.entries(taskMeta)) {
    if (m[propName] !== groupValue) continue;
    const qtyPlan = Number(m.qtyPlan || 0);
    const hh      = Number(m.hhPlan  || 0);
    if (qtyPlan <= 0 || hh <= 0) continue;
    const cum = Number(cumQtyMap[k] || 0);
    // FIX 2: sin cap → permite reflejar sobreejcución por actividad también
    const p = Math.max(0, cum / qtyPlan);
    num += p * hh;
  }
  return denomGroupHH > 0 ? (num / denomGroupHH) : null;
}

function pctFisicoActDesdeQty_(taskMeta, act, denomActHH, cumQtyMap) {
  return pctFisicoGrupoDesdeQty_(taskMeta, "actividad", act, denomActHH, cumQtyMap);
}

// ============================================================
//  CORTE POR TAREA (real acumulado al corte)
//  FIX 3: rendReal = HH acum total / Qty acum total al corte
//          → ratio estable, no distorsionado por semanas vacías
// ============================================================
function construirCortePorTarea_(shHH, shQty, taskMeta, cutMonday) {
  const hhByTask  = sumaHastaCortePorTarea_(shHH,  cutMonday, "Valor");
  const qtyByTask = sumaHastaCortePorTarea_(shQty, cutMonday, "Valor");

  const out = {};
  for (const key of Object.keys(taskMeta)) {
    const hh  = Number(hhByTask[key]  || 0);
    const qty = Number(qtyByTask[key] || 0);

    // FIX 3: rendReal como totales acumulados al corte
    // Solo se calcula si ambos valores son positivos y coherentes
    let rendReal = null;
    if (qty > 0 && hh > 0) {
      rendReal = hh / qty;
      // Sanity check: si el rendimiento calculado es absurdo (> 10x el teórico), descartarlo
      const m = taskMeta[key];
      if (m && m.rendPlan && isFinite(m.rendPlan) && m.rendPlan > 0) {
        if (rendReal > m.rendPlan * 10) {
          Logger.log(`[WARN] rendReal absurdo para "${key}": ${rendReal.toFixed(2)} vs plan ${m.rendPlan.toFixed(2)}. Se descarta.`);
          rendReal = null;
        }
      }
    }

    out[key] = { hhReal: hh, qtyReal: qty, rendReal };
  }
  return out;
}

function calcularHHRestantesPorRendActual_(taskMeta, byTaskCut) {
  let total = 0;
  for (const [k, m] of Object.entries(taskMeta)) {
    const qtyPlan = Number(m.qtyPlan || 0);
    if (qtyPlan <= 0) continue;

    const real    = byTaskCut[k] || { hhReal: 0, qtyReal: 0, rendReal: null };
    const qtyReal = Number(real.qtyReal || 0);
    const qtyRem  = Math.max(0, qtyPlan - qtyReal);

    // FIX 3: usar rendReal (acumulado) con fallback al teórico
    let rend = real.rendReal;
    if (rend == null || !isFinite(rend) || rend <= 0) rend = m.rendPlan;

    if (rend != null && isFinite(rend) && rend > 0) total += qtyRem * rend;
  }
  return total;
}

// ============================================================
//  LECTURAS AGRUPADAS
// ============================================================
function sumaPorSemana_(sheet, weekKeyColName, valorColName) {
  const values  = sheet.getDataRange().getValues();
  const headers = values[0].map(String);
  const iWk     = headers.indexOf(weekKeyColName);
  const iVal    = headers.indexOf(valorColName);
  if (iWk === -1 || iVal === -1) throw new Error(`No encuentro columnas en: ${sheet.getName()}`);

  const map = {};
  for (let r = 1; r < values.length; r++) {
    const wk = values[r][iWk];
    if (!wk) continue;
    map[wk] = (map[wk] || 0) + Number(values[r][iVal] || 0);
  }
  return map;
}

function sumaPorSemanaGrupo_(sheet, weekKeyColName, groupColName, valorColName) {
  const values  = sheet.getDataRange().getValues();
  const headers = values[0].map(String);

  const iWk  = headers.indexOf(weekKeyColName);
  const iGroup = headers.indexOf(groupColName);
  const iVal = headers.indexOf(valorColName);

  if (iWk === -1 || iGroup === -1 || iVal === -1) {
    throw new Error(`En "${sheet.getName()}" faltan columnas: ${weekKeyColName}, ${groupColName}, ${valorColName}.`);
  }

  const out = {};
  for (let r = 1; r < values.length; r++) {
    const wk  = String(values[r][iWk]  || "").trim();
    const group = normalizeKey_(values[r][iGroup]);
    const v   = Number(values[r][iVal] || 0);
    if (!wk || !group) continue;

    if (!out[wk]) out[wk] = {};
    out[wk][group] = (out[wk][group] || 0) + v;
  }
  return out;
}

function sumaPorSemanaActividad_(sheet, weekKeyColName, actColName, valorColName) {
  return sumaPorSemanaGrupo_(sheet, weekKeyColName, actColName, valorColName);
}

function cantidadPorSemanaPorTarea_(sheet, weekKeyColName, actColName, tareaColName, valorColName) {
  const values  = sheet.getDataRange().getValues();
  const headers = values[0].map(String);

  const iWk  = headers.indexOf(weekKeyColName);
  const iAct = headers.indexOf(actColName);
  const iTar = headers.indexOf(tareaColName);
  const iVal = headers.indexOf(valorColName);

  if (iWk === -1 || iAct === -1 || iTar === -1 || iVal === -1) {
    throw new Error(`En "${sheet.getName()}" faltan columnas para agrupar por tarea.`);
  }

  const out = {};
  for (let r = 1; r < values.length; r++) {
    const wk  = String(values[r][iWk] || "").trim();
    // FIX 4: normalizar act y tar al construir la clave
    const act = normalizeKey_(values[r][iAct]);
    const tar = normalizeKey_(values[r][iTar]);
    const v   = Number(values[r][iVal] || 0);
    if (!wk || !act || !tar) continue;

    const key = act + "||" + tar;
    if (!out[wk]) out[wk] = {};
    out[wk][key] = (out[wk][key] || 0) + v;
  }
  return out;
}

function sumaHastaCortePorTarea_(sheet, cutMonday, valorColName) {
  const v = sheet.getDataRange().getValues();
  const h = v[0].map(String);

  const iWk  = h.indexOf("SEMANAS.WEEK_KEY");
  const iAct = h.indexOf("ACTIVIDAD");
  const iTar = h.indexOf("TAREA");
  const iVal = h.indexOf(valorColName);

  if (iWk === -1 || iAct === -1 || iTar === -1 || iVal === -1) {
    throw new Error(`En "${sheet.getName()}" faltan columnas: SEMANAS.WEEK_KEY, ACTIVIDAD, TAREA, ${valorColName}.`);
  }

  const map = {};
  for (let r = 1; r < v.length; r++) {
    const wk = String(v[r][iWk] || "").trim();
    if (!wk) continue;
    if (weekKeyToDate_(wk) > cutMonday) continue;

    // FIX 4: normalizar act y tar al construir la clave
    const act = normalizeKey_(v[r][iAct]);
    const tar = normalizeKey_(v[r][iTar]);
    if (!act || !tar) continue;

    const key = act + "||" + tar;
    map[key] = (map[key] || 0) + Number(v[r][iVal] || 0);
  }
  return map;
}

// ============================================================
//  UTILIDADES SEMANAS / FECHAS
// ============================================================
function addDays_(date, days) {
  const d = new Date(date);
  d.setDate(d.getDate() + days);
  return d;
}

function startOfWeekMonday_(date) {
  const d = new Date(date);
  d.setHours(0, 0, 0, 0);
  const day  = d.getDay();
  const diff = (day === 0 ? -6 : 1 - day);
  d.setDate(d.getDate() + diff);
  return d;
}

function weekKeyToDate_(wk) {
  const m = /^W(\d{4})(\d{2})$/.exec(String(wk).trim());
  if (!m) return new Date("1970-01-01");

  const year = Number(m[1]);
  const week = Number(m[2]);

  const jan4       = new Date(year, 0, 4);
  const jan4Day    = jan4.getDay() || 7;
  const mondayWk1  = new Date(jan4);
  mondayWk1.setDate(jan4.getDate() - (jan4Day - 1));
  mondayWk1.setHours(0, 0, 0, 0);

  const monday = new Date(mondayWk1);
  monday.setDate(mondayWk1.getDate() + (week - 1) * 7);
  monday.setHours(0, 0, 0, 0);
  return monday;
}

function dateToWeekKey_(date) {
  const d = new Date(date);
  d.setHours(0, 0, 0, 0);

  const day      = d.getDay() || 7;
  const thursday = new Date(d);
  thursday.setDate(d.getDate() + (4 - day));
  const isoYear = thursday.getFullYear();

  const jan4      = new Date(isoYear, 0, 4);
  const jan4Day   = jan4.getDay() || 7;
  const mondayWk1 = new Date(jan4);
  mondayWk1.setDate(jan4.getDate() - (jan4Day - 1));
  mondayWk1.setHours(0, 0, 0, 0);

  const weekNo = Math.floor((startOfWeekMonday_(d) - mondayWk1) / (7 * 24 * 3600 * 1000)) + 1;
  return `W${isoYear}${String(weekNo).padStart(2, "0")}`;
}

function encontrarUltimoHasta_(curveMap, weeks, cutMonday) {
  let last = null;
  for (const wk of weeks) {
    if (weekKeyToDate_(wk) <= cutMonday) last = curveMap[wk];
  }
  return last || { hhPlanAc: 0, hhRealAc: 0, pctPlan: 0, pctReal: 0 };
}

function interpolarHHPlanPorPct_(curveMap, weeks, targetPct) {
  if (targetPct <= 0) return 0;

  const pts = [];
  for (const wk of weeks) {
    const v = curveMap[wk];
    if (v && v.pctPlan != null && v.hhPlanAc != null) pts.push({ x: v.pctPlan, y: v.hhPlanAc });
  }
  if (pts.length === 0) return 0;

  const max = pts[pts.length - 1];
  if (targetPct >= max.x) return max.y;

  for (let i = 1; i < pts.length; i++) {
    const p0 = pts[i - 1], p1 = pts[i];
    if (targetPct <= p1.x) {
      const dx = p1.x - p0.x;
      if (dx <= 0) return p1.y;
      const t = (targetPct - p0.x) / dx;
      return p0.y + t * (p1.y - p0.y);
    }
  }
  return max.y;
}

function ultimasNSemanasHasta_(weeksSorted, cutMonday, n) {
  const eligible = weeksSorted.filter(w => weekKeyToDate_(w) <= cutMonday);
  return eligible.slice(Math.max(0, eligible.length - n));
}

function hhAcumGrupoHastaCorte_(hhByWeekGroup, groupValue, cutMonday) {
  let sum = 0;
  const wks = Object.keys(hhByWeekGroup).sort((a, b) => weekKeyToDate_(a) - weekKeyToDate_(b));
  for (const wk of wks) {
    if (weekKeyToDate_(wk) > cutMonday) break;
    sum += Number((hhByWeekGroup[wk] || {})[groupValue] || 0);
  }
  return sum;
}

function hhAcumActHastaCorte_(hhByWeekAct, act, cutMonday) {
  return hhAcumGrupoHastaCorte_(hhByWeekAct, act, cutMonday);
}

function formatISODate_(d) {
  const tz = Session.getScriptTimeZone();
  return Utilities.formatDate(d, tz, "yyyy-MM-dd");
}

function formatISODateTime_(d) {
  const tz = Session.getScriptTimeZone();
  return Utilities.formatDate(d, tz, "yyyy-MM-dd HH:mm");
}

// ============================================================
//  PROYECCIONES (Escenario A y B)
// ============================================================
/**
 * Calcula dos series de proyección a partir del corte:
 *
 * Escenario A — "Mismo consumo de HH, ¿cuándo termino?"
 *   Usa el ritmo de avance físico semanal promedio real
 *   (últimas 4 semanas) y lo extrapola hacia adelante.
 *   Las HH proyectadas siguen el mismo ritmo real.
 *
 * Escenario B — "Termino en el mismo plazo, ¿cuántas HH necesito?"
 *   Toma la semana de fin del plan original y distribuye el
 *   avance faltante linealmente hasta esa semana.
 *   Las HH se calculan con el rendimiento real (HH por 1% de avance).
 *
 * Retorna:
 *   { escA: [[wk, hhAcum, pctFisico], ...],
 *     escB: [[wk, hhAcum, pctFisico], ...],
 *     finEscA: weekKey,   // semana proyectada de término Esc A
 *     finPlan: weekKey,   // semana de fin del plan original
 *     hhExtraEscB: number // HH adicionales requeridas Esc B
 *   }
 */
function calcularProyecciones_(curveMap, weeks, cutMonday, pctRealCut, hhRealCut, globalHH, hhPendRend, taskMeta, byTaskCut) {

  // Semana de fin del plan = última semana con pctPlan > 0
  let finPlanWk = null;
  for (let i = weeks.length - 1; i >= 0; i--) {
    if ((curveMap[weeks[i]].pctPlan || 0) > 0) { finPlanWk = weeks[i]; break; }
  }

  // Semanas futuras (posteriores al corte)
  const futureWeeks = weeks.filter(wk => weekKeyToDate_(wk) > cutMonday);

  // Ritmo real: promedio de avance físico semanal en las últimas 4 semanas reales
  const realWeeks = weeks.filter(wk => weekKeyToDate_(wk) <= cutMonday && curveMap[wk].pctReal != null);
  const last4 = realWeeks.slice(Math.max(0, realWeeks.length - 4));
  let avgAvanceSem = 0;
  if (last4.length >= 2) {
    const pctStart = curveMap[last4[0]].pctReal || 0;
    const pctEnd   = curveMap[last4[last4.length-1]].pctReal || 0;
    avgAvanceSem   = (pctEnd - pctStart) / (last4.length - 1);
  } else {
    // Fallback: ritmo global dividido en semanas transcurridas
    avgAvanceSem = realWeeks.length > 0 ? pctRealCut / realWeeks.length : 0.01;
  }
  // Nunca negativo ni cero
  if (avgAvanceSem <= 0) avgAvanceSem = 0.005;

  // HH por punto de avance real (rendimiento global al corte)
  const hhPor1pct = pctRealCut > 0 ? hhRealCut / (pctRealCut * 100) : (globalHH / 100);

  // ── Escenario A: mismo ritmo ──
  const escA = [];
  let pctA   = pctRealCut;
  let hhA    = hhRealCut;
  let finEscA = finPlanWk; // por defecto fin de plan si ya terminó

  for (const wk of futureWeeks) {
    if (pctA >= 1.0) break; // ya llegó al 100%
    pctA = Math.min(1.0, pctA + avgAvanceSem);
    hhA  = hhA + avgAvanceSem * 100 * hhPor1pct;
    escA.push([wk, hhA, pctA]);
    if (pctA >= 1.0) { finEscA = wk; break; }
  }
  // Si no llegó al 100% en las semanas del plan, extender
  if (pctA < 1.0 && avgAvanceSem > 0) {
    const weeksNeeded = Math.ceil((1.0 - pctA) / avgAvanceSem);
    let lastDate = weekKeyToDate_(weeks[weeks.length - 1]);
    for (let i = 1; i <= weeksNeeded && i <= 52; i++) {
      lastDate = new Date(lastDate.getTime() + 7 * 24 * 3600 * 1000);
      const wk = dateToWeekKey_(lastDate);
      pctA = Math.min(1.0, pctA + avgAvanceSem);
      hhA  = hhA + avgAvanceSem * 100 * hhPor1pct;
      escA.push([wk, hhA, pctA]);
      if (pctA >= 1.0) { finEscA = wk; break; }
    }
  }

  // ── Escenario B: misma fecha de fin ──
  const escB = [];
  let hhExtraEscB = 0;

  if (finPlanWk) {
    const weeksToFin = futureWeeks.filter(wk => weekKeyToDate_(wk) <= weekKeyToDate_(finPlanWk));
    const nSem = weeksToFin.length;

    if (nSem > 0) {
      const pctFaltante  = Math.max(0, 1.0 - pctRealCut);
      const avancePorSem = pctFaltante / nSem;

      // HH necesarias para ese avance adicional por semana con rendimiento real
      // (si el rendimiento real es peor, necesito más HH por % avanzado)
      const hhPorPctReal = pctRealCut > 0 ? (hhRealCut / (pctRealCut)) : globalHH;
      const hhPorSemB    = avancePorSem * hhPorPctReal;

      let pctB = pctRealCut;
      let hhB  = hhRealCut;

      for (const wk of weeksToFin) {
        pctB = Math.min(1.0, pctB + avancePorSem);
        hhB  = hhB + hhPorSemB;
        escB.push([wk, hhB, pctB]);
      }
      hhExtraEscB = Math.max(0, hhB - hhRealCut - hhPendRend);
    }
  }

  return {
    escA,
    escB,
    finEscA,
    finPlan  : finPlanWk,
    hhExtraEscB,
    avgAvanceSemReal: avgAvanceSem,
    hhPor1pct
  };
}

// ============================================================
//  HELPERS
// ============================================================

/**
 * FIX 4: Normaliza una clave de actividad/tarea para comparación.
 * - Trim de espacios al inicio/fin
 * - Colapsa espacios internos múltiples en uno
 * - Elimina el carácter de non-breaking space (\xa0)
 * - Convierte a minúsculas para comparación case-insensitive
 *
 * IMPORTANTE: La normalización se aplica a TODAS las fuentes
 * (plan, HH real, qty real) de forma consistente, por lo que
 * las claves resultantes siempre matchearán entre sí.
 */
function normalizeKey_(x) {
  if (x === null || x === undefined) return "";
  return String(x)
    .replace(/\xa0/g, " ")   // non-breaking space → space
    .replace(/\s+/g, " ")    // múltiples espacios → uno
    .trim()
    .toLowerCase();
}

function toNumBack_(x) {
  if (x === null || x === undefined || x === "") return 0;
  // FIX: mismo problema que HH TOTALES — valores numéricos de Excel
  // interpretados como Date por Google Sheets al abrir el xlsx.
  // El número de días desde el epoch de Excel (1899-12-30) es el valor original.
  if (x instanceof Date) {
    // UTC para evitar desfases por timezone del script
    const EXCEL_EPOCH_MS = Date.UTC(1899, 11, 30);
    const diffDays = Math.round((x.getTime() - EXCEL_EPOCH_MS) / (1000 * 60 * 60 * 24));
    return (diffDays > 0 && diffDays < 1e6) ? diffDays : 0;
  }
  if (typeof x === "number") return isFinite(x) ? x : 0;
  let s = String(x).trim();
  if (!s) return 0;
  s = s.replace(/\s+/g, "");
  if (s.includes(",") && s.includes(".")) s = s.replace(/\./g, "").replace(",", ".");
  else if (s.includes(",") && !s.includes("."))  s = s.replace(",", ".");
  const n = Number(s);
  return isFinite(n) ? n : 0;
}

function diagnosticoDenominador_(taskMeta) {
  let total = 0, conQty = 0, conHH = 0, conAmbos = 0;
  let ejemploSinHH = "-", ejemploSinQty = "-";
  for (const [k, m] of Object.entries(taskMeta)) {
    total++;
    const q = Number(m.qtyPlan || 0);
    const h = Number(m.hhPlan  || 0);
    if (q > 0) conQty++;
    if (h > 0) conHH++;
    if (q > 0 && h > 0) conAmbos++;
    if (ejemploSinHH  === "-" && q > 0 && !(h > 0)) ejemploSinHH  = `${m.actividad}||${m.tarea} qty=${q} hh=${h}`;
    if (ejemploSinQty === "-" && h > 0 && !(q > 0)) ejemploSinQty = `${m.actividad}||${m.tarea} qty=${q} hh=${h}`;
  }
  return { total, conQty, conHH, conAmbos, ejemploSinHH, ejemploSinQty };
}
