/**
 * COCO Tecnologías — Capa de escritura del Dashboard Ejecutivo
 * ============================================================
 * Este script vive DENTRO del Google Sheet "BD_MAESTRA_COCO"
 * (Extensiones → Apps Script) y se despliega como Web App.
 *
 * Endpoints:
 *   doGet  → entrega todas las hojas del modelo como arrays (AOA),
 *            el mismo formato que parseWorkbook() del dashboard consume.
 *   doPost → acciones de escritura (JSON en el body):
 *     - registro_individual : alta de una fila en BD_Indicadores
 *     - carga_masiva        : plantilla de área completa, con modo
 *                             "previsualizar" (no escribe) o "confirmar"
 *     - pipeline_masivo     : filas de Pipeline_Comercial, con modo
 *                             "reemplazar" (snapshot) o "agregar"
 *
 * Seguridad (ver README):
 *   - Lectura: pública (quien tenga la URL del Web App).
 *   - Escritura: requiere token compartido (Script Properties, clave
 *     TOKEN_ESCRITURA) + correo del dominio permitido en el payload.
 *     Session.getActiveUser() se registra cuando está disponible
 *     (llega vacío en llamadas fetch() desde HTML externo — limitación
 *     de CORS de Apps Script, documentada).
 *   - Todo cambio queda en la pestaña LOG.
 *
 * Despliegue: Implementar → Nueva implementación → Aplicación web
 *   Ejecutar como: yo · Acceso: cualquier usuario
 */

/* ================= CONFIGURACIÓN ================= */

// Solo estos dominios pueden ESCRIBIR. cocotech.ai = COCO; cfocus.co = CFO externo.
var DOMINIOS_ESCRITURA = ['cocotech.ai', 'cfocus.co'];
var HOJA_BD = 'BD_Indicadores';
var HOJA_PIPELINE = 'Pipeline_Comercial';
var HOJA_LOG = 'LOG';
var HOJA_CATALOGOS = 'CATALOGOS';

/* Hojas que doGet entrega al dashboard (mismo set que parseWorkbook lee) */
// TRM_Peru es indispensable: sin ella el bloque internacional pierde la conversión de Perú.
var HOJAS_LECTURA = ['Diccionario', 'BD_Indicadores', 'TRM', 'TRM_Peru',
  'Analisis_Reconciliacion', 'Analisis_Churn', 'Calc_LTV_Fin',
  'Pipeline_Comercial', 'CATALOGOS', 'BD_GASTOS_RESUMEN_MENSUAL'];

var COLS_BD = ['periodo', 'pais', 'compania', 'moneda', 'escenario',
  'codigo_indicador', 'segmento', 'detalle', 'valor', 'unidad', 'comentario'];

var COLS_PIPELINE = ['cliente', 'comercial', 'tipo', 'prob', 'mes',
  'mrr_cop', 'impl_cop', 'citas', 'valor_cop', 'ponderado_cop'];

/* ================= PUNTO DE ENTRADA =================
   Una sola dirección sirve para dos cosas:
     .../exec            -> entrega el TABLERO (página web)
     .../exec?api=1      -> entrega los DATOS en JSON (para un HTML externo con token)
   Al abrir el tablero desde aquí, Google ya sabe quién entró: la identidad es
   verificada, no declarada, y no hace falta token. */

function doGet(e) {
  if (e && e.parameter && e.parameter.api === '1') return datosJson_();
  return HtmlService.createHtmlOutputFromFile('Dashboard')
    .setTitle('COCO Tecnologías · Cockpit Financiero')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function datosJson_() {
  try {
    return json_({ ok: true, generado: new Date().toISOString(), hojas: leerHojas_() });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  }
}

function leerHojas_() {
  var libro = SpreadsheetApp.getActiveSpreadsheet();
  var hojas = {};
  HOJAS_LECTURA.forEach(function (nombre) {
    var h = libro.getSheetByName(nombre);
    hojas[nombre] = h ? h.getDataRange().getValues() : [];
  });
  return hojas;
}

/* ========== FUNCIONES QUE LLAMA EL TABLERO SERVIDO DESDE AQUÍ ==========
   Se invocan con google.script.run. No reciben token: la identidad la pone Google.
   Requisito de despliegue: "Ejecutar como: yo" + "Acceso: usuarios de cocotech.ai".
   Así el usuario NO necesita permisos sobre el Sheet (no puede saltarse la validación),
   y aun así getActiveUser() devuelve su correo real por ser del mismo dominio. */

function appDatos() {
  return { ok: true, generado: new Date().toISOString(), hojas: leerHojas_() };
}

function appUsuario() {
  return { correo: usuarioVerificado_() };
}

function usuarioVerificado_() {
  try { return (Session.getActiveUser().getEmail() || '').toLowerCase(); }
  catch (e) { return ''; }
}

/** Escritura desde el tablero servido por Apps Script (sin token). */
function appEscribir(datos) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    var correo = usuarioVerificado_();
    if (!correo) {
      return { ok: false, error: 'No pude verificar tu identidad. Entra con tu cuenta de COCO.' };
    }
    var dominio = correo.split('@')[1];
    if (DOMINIOS_ESCRITURA.indexOf(dominio) < 0) {
      return { ok: false, error: 'El dominio ' + dominio + ' no tiene permiso de escritura.' };
    }
    var res;
    switch (datos.accion) {
      case 'registro_individual': res = registroIndividual_(datos, correo); break;
      case 'carga_masiva':        res = cargaMasiva_(datos, correo); break;
      case 'pipeline_masivo':     res = pipelineMasivo_(datos, correo); break;
      default: return { ok: false, error: 'Acción no reconocida: ' + datos.accion };
    }
    return JSON.parse(res.getContent());   // las funciones internas devuelven ContentService
  } catch (err) {
    return { ok: false, error: String(err) };
  } finally {
    lock.releaseLock();
  }
}

/* ================= ENDPOINT DE ESCRITURA ================= */

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000); // concurrencia: 6-15 personas
  try {
    var datos = JSON.parse(e.postData.contents);
    var auth = validarAcceso_(datos);
    if (!auth.ok) return json_({ ok: false, error: auth.error });

    switch (datos.accion) {
      case 'registro_individual': return registroIndividual_(datos, auth.usuario);
      case 'carga_masiva':        return cargaMasiva_(datos, auth.usuario);
      case 'pipeline_masivo':     return pipelineMasivo_(datos, auth.usuario);
      default:
        return json_({ ok: false, error: 'Acción no reconocida: ' + datos.accion });
    }
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    lock.releaseLock();
  }
}

/* ================= AUTENTICACIÓN ================= */

function validarAcceso_(datos) {
  var tokenEsperado = PropertiesService.getScriptProperties().getProperty('TOKEN_ESCRITURA');
  if (!tokenEsperado) {
    return { ok: false, error: 'El administrador no ha configurado TOKEN_ESCRITURA en Script Properties.' };
  }
  if (datos.token !== tokenEsperado) {
    return { ok: false, error: 'Token de escritura inválido.' };
  }
  // Identidad: la sesión de Google si existe; si no, el correo declarado en el payload
  var sesion = '';
  try { sesion = Session.getActiveUser().getEmail() || ''; } catch (ignorado) {}
  var usuario = (sesion || String(datos.usuario || '')).trim().toLowerCase();
  if (!usuario || usuario.indexOf('@') < 0) {
    return { ok: false, error: 'Falta el correo del usuario (campo "usuario").' };
  }
  var dominio = usuario.split('@')[1];
  if (DOMINIOS_ESCRITURA.indexOf(dominio) < 0) {
    return { ok: false, error: 'El dominio ' + dominio + ' no tiene permiso de escritura.' };
  }
  return { ok: true, usuario: usuario + (sesion ? '' : ' (declarado)') };
}

/* ================= CATÁLOGOS Y VALIDACIÓN ================= */

function leerCatalogos_() {
  var h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_CATALOGOS);
  if (!h) throw new Error('No existe la pestaña ' + HOJA_CATALOGOS);
  var vals = h.getDataRange().getValues();
  var porArea = {};      // area → Set de códigos
  var listas = { paises: [], escenarios: [], monedas: [], companias: [] };
  for (var r = 1; r < vals.length; r++) {
    var area = String(vals[r][0] || '').trim();
    var cod = String(vals[r][1] || '').trim();
    if (area && cod) {
      if (!porArea[area]) porArea[area] = {};
      porArea[area][cod] = true;
    }
    if (vals[r][3]) listas.paises.push(String(vals[r][3]).trim());
    if (vals[r][4]) listas.escenarios.push(String(vals[r][4]).trim());
    if (vals[r][5]) listas.monedas.push(String(vals[r][5]).trim());
    if (vals[r][6]) listas.companias.push(String(vals[r][6]).trim());
  }
  return { porArea: porArea, listas: listas };
}

/** Valida una fila (objeto con claves COLS_BD) contra el catálogo del área.
 *  Devuelve lista de motivos de error; vacía = fila válida. */
function validarFila_(f, area, cat) {
  var errores = [];
  var periodo = String(f.periodo || '').trim();
  if (!/^\d{4}-(0[1-9]|1[0-2])$/.test(periodo)) {
    errores.push('periodo debe ser AAAA-MM (llegó: "' + periodo + '")');
  }
  var codigo = String(f.codigo_indicador || '').trim();
  if (!codigo) {
    errores.push('falta codigo_indicador');
  } else if (!cat.porArea[area]) {
    errores.push('área desconocida: ' + area);
  } else if (!cat.porArea[area][codigo]) {
    errores.push('el código "' + codigo + '" no pertenece al área ' + area);
  }
  var v = f.valor;
  if (v === null || v === undefined || String(v).trim() === '') {
    errores.push('falta valor');
  } else if (typeof v !== 'number') {
    var n = Number(String(v).replace(/[ ,]/g, ''));
    if (isNaN(n)) errores.push('valor no numérico: "' + v + '"');
  }
  var esc = String(f.escenario || '').trim();
  if (esc && cat.listas.escenarios.indexOf(esc) < 0) {
    errores.push('escenario inválido: "' + esc + '" (usa ' + cat.listas.escenarios.join('/') + ')');
  }
  return errores;
}

/* ================= ESCRITURA EN BD_Indicadores ================= */

/** Encuentra la fila de encabezado real de BD_Indicadores (mismo criterio
 *  que parseWorkbook: primera fila que contiene periodo y codigo_indicador). */
function encabezadoBD_(hoja) {
  var top = hoja.getRange(1, 1, Math.min(12, hoja.getLastRow()), hoja.getLastColumn()).getValues();
  for (var r = 0; r < top.length; r++) {
    var fila = top[r].map(function (c) { return String(c || '').toLowerCase(); });
    if (fila.indexOf('periodo') >= 0 && fila.indexOf('codigo_indicador') >= 0) {
      return { fila: r + 1, cols: fila };
    }
  }
  throw new Error('No se encontró el encabezado de ' + HOJA_BD);
}

/** Convierte un objeto fila a array en el orden físico de la hoja,
 *  rellenando area/area_responsable si esas columnas existen. */
function filaFisica_(f, area, cols) {
  return cols.map(function (nombre) {
    if (nombre === 'area' || nombre === 'area_responsable') return f[nombre] || area;
    if (nombre === 'valor') {
      return (typeof f.valor === 'number') ? f.valor : Number(String(f.valor).replace(/[ ,]/g, ''));
    }
    return (f[nombre] !== undefined && f[nombre] !== null) ? f[nombre] : '';
  });
}

function escribirFilasBD_(filasObj, area) {
  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_BD);
  var enc = encabezadoBD_(hoja);
  var fisicas = filasObj.map(function (f) { return filaFisica_(f, area, enc.cols); });
  hoja.getRange(hoja.getLastRow() + 1, 1, fisicas.length, enc.cols.length).setValues(fisicas);
  return fisicas.length;
}

/* ================= ACCIONES ================= */

/** payload: {accion, token, usuario, area, fila:{periodo,...,valor,...}} */
function registroIndividual_(datos, usuario) {
  if (!datos.area || !datos.fila) return json_({ ok: false, error: 'Faltan area o fila.' });
  var cat = leerCatalogos_();
  var errores = validarFila_(datos.fila, datos.area, cat);
  if (errores.length) return json_({ ok: false, error: errores.join(' · ') });
  escribirFilasBD_([datos.fila], datos.area);
  log_(usuario, 'registro_individual', datos.area, HOJA_BD, 1,
    datos.fila.codigo_indicador + ' ' + datos.fila.periodo + ' = ' + datos.fila.valor);
  return json_({ ok: true, filas: 1 });
}

/** payload: {accion, token, usuario, area, modo:'previsualizar'|'confirmar',
 *            filas:[{periodo,...}, ...]}
 *  previsualizar → valida todo y devuelve el detalle, sin escribir.
 *  confirmar     → revalida y escribe SOLO las filas válidas. */
function cargaMasiva_(datos, usuario) {
  if (!datos.area || !Array.isArray(datos.filas) || !datos.filas.length) {
    return json_({ ok: false, error: 'Faltan area o filas.' });
  }
  var cat = leerCatalogos_();
  var validas = [], detalle = [];
  datos.filas.forEach(function (f, i) {
    var errs = validarFila_(f, datos.area, cat);
    if (errs.length) detalle.push({ fila: i + 1, errores: errs });
    else validas.push(f);
  });
  var resumen = { ok: true, modo: datos.modo, total: datos.filas.length,
    validas: validas.length, conError: detalle.length, errores: detalle };

  if (datos.modo === 'confirmar') {
    if (!validas.length) return json_({ ok: false, error: 'Ninguna fila pasó la validación.', errores: detalle });
    resumen.escritas = escribirFilasBD_(validas, datos.area);
    log_(usuario, 'carga_masiva', datos.area, HOJA_BD, validas.length,
      'total=' + datos.filas.length + ' rechazadas=' + detalle.length);
  }
  return json_(resumen);
}

/** payload: {accion, token, usuario, modo:'reemplazar'|'agregar',
 *            filas:[{cliente,comercial,tipo,prob,mes,mrr_cop,impl_cop,citas}, ...]}
 *  valor_cop y ponderado_cop se calculan aquí (no se confía en el cliente). */
function pipelineMasivo_(datos, usuario) {
  if (!Array.isArray(datos.filas) || !datos.filas.length) {
    return json_({ ok: false, error: 'Faltan filas de pipeline.' });
  }
  var detalle = [], fisicas = [];
  datos.filas.forEach(function (f, i) {
    var errs = [];
    if (!String(f.cliente || '').trim()) errs.push('falta cliente');
    var prob = Number(f.prob);
    if (isNaN(prob) || prob < 0 || prob > 1) {
      if (!isNaN(prob) && prob > 1 && prob <= 100) prob = prob / 100; // acepta 60 → 0.60
      else errs.push('prob debe estar entre 0 y 1 (o 0 y 100)');
    }
    var mrr = Number(f.mrr_cop || 0), impl = Number(f.impl_cop || 0);
    if (isNaN(mrr) || isNaN(impl)) errs.push('mrr_cop/impl_cop no numéricos');
    if (errs.length) { detalle.push({ fila: i + 1, errores: errs }); return; }
    var valor = mrr + impl;
    fisicas.push([String(f.cliente).trim(), f.comercial || '', f.tipo || '', prob,
      f.mes || '', mrr, impl, Number(f.citas || 0), valor, valor * prob]);
  });
  if (!fisicas.length) return json_({ ok: false, error: 'Ninguna fila válida.', errores: detalle });

  var hoja = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_PIPELINE);
  if (!hoja) throw new Error('No existe la pestaña ' + HOJA_PIPELINE);
  if (datos.modo === 'reemplazar' && hoja.getLastRow() > 1) {
    hoja.getRange(2, 1, hoja.getLastRow() - 1, COLS_PIPELINE.length).clearContent();
  }
  var inicio = (datos.modo === 'reemplazar') ? 2 : hoja.getLastRow() + 1;
  hoja.getRange(inicio, 1, fisicas.length, COLS_PIPELINE.length).setValues(fisicas);
  log_(usuario, 'pipeline_masivo(' + (datos.modo || 'agregar') + ')', 'Comercial_Ventas',
    HOJA_PIPELINE, fisicas.length, 'rechazadas=' + detalle.length);
  return json_({ ok: true, escritas: fisicas.length, conError: detalle.length, errores: detalle });
}

/* ================= LOG Y UTILIDADES ================= */

function log_(usuario, accion, area, hoja, filas, detalle) {
  var h = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(HOJA_LOG);
  if (!h) return; // el log nunca debe tumbar una escritura
  h.appendRow([new Date(), usuario, accion, area, hoja, filas, detalle || '']);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ================= PRUEBAS MANUALES (Ejecutar en el editor) ================= */

/** 1) Ejecutar UNA VEZ para fijar el token de escritura (cámbialo). */
function configurarToken() {
  PropertiesService.getScriptProperties().setProperty('TOKEN_ESCRITURA', 'CAMBIAME-token-largo-y-unico');
  Logger.log('Token configurado.');
}

/** 2) Prueba de lectura: debe listar las hojas con su número de filas. */
function probarLectura() {
  var r = JSON.parse(doGet({}).getContent());
  Object.keys(r.hojas).forEach(function (k) { Logger.log(k + ': ' + r.hojas[k].length + ' filas'); });
}

/** 3) Prueba de escritura individual (usa un código real de Finanzas). */
function probarRegistroIndividual() {
  var token = PropertiesService.getScriptProperties().getProperty('TOKEN_ESCRITURA');
  var r = doPost({ postData: { contents: JSON.stringify({
    accion: 'registro_individual', token: token, usuario: 'prueba@' + DOMINIOS_ESCRITURA[0],
    area: 'Finanzas',
    fila: { periodo: '2026-07', pais: 'Colombia', compania: 'Coco Colombia',
      moneda: 'COP', escenario: 'Real', codigo_indicador: 'ingresos_totales',
      segmento: '', detalle: '', valor: 123456789, unidad: 'COP',
      comentario: 'FILA DE PRUEBA — borrar' }
  }) } });
  Logger.log(r.getContent());
}

/** 4) Prueba de carga masiva en modo previsualización (no escribe nada). */
function probarCargaMasivaPreview() {
  var token = PropertiesService.getScriptProperties().getProperty('TOKEN_ESCRITURA');
  var r = doPost({ postData: { contents: JSON.stringify({
    accion: 'carga_masiva', token: token, usuario: 'prueba@' + DOMINIOS_ESCRITURA[0],
    area: 'Marketing', modo: 'previsualizar',
    filas: [
      { periodo: '2026-07', codigo_indicador: 'leads_generados', valor: 250, escenario: 'Real' },
      { periodo: '2026-13', codigo_indicador: 'leads_generados', valor: 250 },       // periodo malo
      { periodo: '2026-07', codigo_indicador: 'ingresos_totales', valor: 1 },        // código de otra área
      { periodo: '2026-07', codigo_indicador: 'cac_fully_loaded', valor: 'abc' }     // valor malo
    ]
  }) } });
  Logger.log(r.getContent());
}
