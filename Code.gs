/**
 * ============================================================================
 *  CONTROL SOX DE TARIFAS — Web App (Google Apps Script)
 * ============================================================================
 *  - "Cerebro" = tu Sheet base (CONFIG.PANEL_SS_ID): pestañas Permisos,
 *    "diccionario postas" y "Resumen SOX".
 *  - Sube el dump (varios archivos a la vez o un ZIP). Cada archivo se crea como
 *    Google Sheets NUEVO con su mismo nombre, dentro de una SUBCARPETA por
 *    trimestre en tu carpeta de Drive.
 *  - Separa Columna A respetando comillas y aplica validaciones SOX EN CÓDIGO:
 *      · Consistencia (aprobador ≠ solicitante)
 *      · Aprobador / Solicitante en la matriz de permisos (+ su Rol y Región)
 *      · Zona/País de origen desde "diccionario postas"
 *      · Campos NULL críticos y alcance (Nature of Spot para Spot/Street)
 *  - Consolida las inconsistencias en "Resumen SOX" y reporta TIEMPOS.
 * ============================================================================
 */

/* ========================== 1. CONFIGURACIÓN ============================== */
var CONFIG = {
  APP_VERSION: '2026.08.18.2',
  // Sheet "cerebro" (donde están Permisos y diccionario postas).
  PANEL_SS_ID: '1URq-lB8S0tOVA1tArK66Jy6geb_SUP5GwdrpRqS-lUw',
  // Carpeta de Drive donde se crean las subcarpetas por trimestre.
  CARPETA_DRIVE_ID: '1QU9283mewZDqjj3c6y_KzLUULRzMtiiO',

  EXTENSIONES_PERMITIDAS: ['zip', 'xlsx', 'xls', 'csv'],
  DELIMITADOR: ',',
  NULL_A_VACIO: false,
  REEMPLAZAR_EXISTENTES: true,
  // Copia las pestañas Permisos + diccionario postas DENTRO de cada Excel generado
  // (para tus formulaciones posteriores dentro de cada archivo).
  COPIAR_HOJAS_APOYO: true,

  HOJA_PERMISOS: 'Permisos',
  HOJA_DICCIONARIO: 'diccionario postas',
  HOJA_RESUMEN: 'Resumen SOX',

  // Pestaña Permisos: columnas (1=A). Correo, Rol, Región.
  PERM_COL_EMAIL: 1, PERM_COL_ROL: 2, PERM_COL_REGION: 3,
  // Diccionario postas: POSTA, PAÍS, ZONA, REGIÓN 2.
  DIC_COL_POSTA: 1, DIC_COL_PAIS: 2, DIC_COL_ZONA: 3, DIC_COL_REGION2: 4,

  TIPOS: [
    { clave: 'SPOT',     hoja: 'Spot',     patrones: ['spot'] },
    { clave: 'PROMO',    hoja: 'Promo',    patrones: ['promo'] },
    { clave: 'STRIP',    hoja: 'Street',   patrones: ['strip', 'street'] },
    { clave: 'CONTRACT', hoja: 'Contract', patrones: ['contract', 'contrato'] }
  ],

  COL_SOLICITANTE: ['Uploaded By', 'Created By', 'Requested By', 'Solicitante'],
  COL_APROBADOR:   ['Approved By', 'APPROVER', 'Approver', 'Aprobador'],
  COL_ORIGEN:      ['Origin', 'Origen'],
  COL_DESTINO:     ['Destination', 'Destino'],
  CAMPOS_CRITICOS: ['Currency', 'Effective From Date', 'Origin', 'Destination', 'Rate Id', 'All-in'],

  COL_NATURALEZA: ['Nature of Spot', 'Spot Nature', 'Nature'],
  SPOT_EN_ALCANCE: [
    'BELOW BP', 'EXTEMP ABOVE CONTRACT', 'EXTEMP ABOVE MIN RATE', 'EXTEMP ABOVE STREET',
    'EXTEMP BELOW BP', 'EXTEMP BELOW CONTRACT', 'EXTEMP BELOW STREET', 'EXTEMP NO RATE'
  ],

  COL_RESUMEN: { rateId: ['Rate Id', 'Spot ID', 'Rate ID'], region: ['File Owner', 'Region', 'Región'] },

  // Columnas de fecha para el filtro del trimestre (descartar arrastres).
  COL_CREADO:   ['Created Date', 'Creation Date', 'Fecha creación'],
  COL_MODIFICADO: ['Last Modified Date', 'Modified Date', 'Fecha modificación'],
  FILTRAR_POR_PERIODO: true,

  HOJA_JUSTIFICADOS: 'Justificados',
  HOJA_CORREOS: 'Correos a enviar',
  HOJA_REPORTE: 'Reporte Final',
  HOJA_HISTORIAL: 'Historial de ejecuciones',
  HOJA_HISTORIAL_COMUNICACIONES: 'Historial comunicaciones',
  HOJA_REPORTE_ENVIO_RESUMEN: 'Resumen ejecutivo',
  HOJA_REPORTE_ENVIO_DETALLE: 'Detalle consolidado',
  PREFIJO_REPORTE_ENVIO: 'Reporte SOX para envío - ',

  // Recordatorio mensual de matriz (día 15, último mes cerrado).
  RESPONSABLE_SOX: 'gabrielacastro.aerosan@latam.com',
  FOCALS_EMAILS: [],   // pega aquí los correos de los focals (SAM/NAM/EUR)
  // Lista o grupos de distribución que reciben UN SOLO correo general trimestral.
  DESTINATARIOS_REPORTE: [],

  // Admins que ven la vista de carga/procesamiento (el resto ve su portal).
  ADMINS: ['gabrielacastro.aerosan@latam.com', 'ALEJANDRA_CORREO_AQUI@latam.com'],
  HOJA_JUSTIFICACIONES: 'Justificaciones',

  // Encabezados de columnas SOX que agrega la herramienta.
  H: {
    consist: 'SOX · Consistencia (Aprob≠Solic)',
    apMat:   'SOX · Aprobador en matriz',
    reqMat:  'SOX · Solicitante en matriz',
    rolAp:   'SOX · Rol aprobador',
    rolReq:  'SOX · Rol solicitante',
    regAp:   'SOX · Región aprobador',
    zona:    'SOX · Zona origen (Región 2)',
    pais:    'SOX · País origen',
    result:  'SOX · Resultado',
    estado:  'SOX · Estado',
    periodo: 'SOX · En período (Q)',
    nul:     'SOX · Campos NULL críticos',
    alc:     'SOX · En alcance'
  }
};

/* ===================== 2. WEB APP + MENÚ ================================= */
function doGet(e) {
  var view = (e && e.parameter && e.parameter.view) || '';
  var user = '';
  try { user = String(Session.getActiveUser().getEmail() || '').toLowerCase(); } catch (_) {}
  var admins = CONFIG.ADMINS.map(function (x) { return String(x).toLowerCase(); });
  var esAdmin = admins.indexOf(user) !== -1;
  var archivo = (esAdmin && view !== 'portal') ? 'index' : 'portal';
  return HtmlService.createHtmlOutputFromFile(archivo)
    .setTitle('Control SOX de Tarifas')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/* ===================== PORTAL DE USUARIO =============================== */
// Casos de un correo dado (aprobador o solicitante) desde el Resumen SOX.
function casosParaEmail_(email) {
  var user = String(email || '').trim().toLowerCase();
  var panel = obtenerPanel_();
  var resumen = hojaPorNombre_(panel, CONFIG.HOJA_RESUMEN);
  var casos = [];
  if (user && resumen && resumen.getLastRow() > 1) {
    var d = resumen.getRange(2, 1, resumen.getLastRow() - 1, resumen.getLastColumn()).getValues();
    d.forEach(function (r) {
      var ap = String(r[8]).trim().toLowerCase(), so = String(r[6]).trim().toLowerCase();
      if (ap === user || so === user) {
        casos.push({ tipo: r[0], rateId: String(r[1]), origen: r[3], destino: r[5], zona: r[4],
          solicitante: r[6], aprobador: r[8], resultado: r[10], periodo: r[12] || 'Período actual' });
      }
    });
  }
  var enviados = {};
  var hj = hojaPorNombre_(panel, CONFIG.HOJA_JUSTIFICACIONES);
  if (hj && hj.getLastRow() > 1) {
    var jv = hj.getRange(2, 1, hj.getLastRow() - 1, 7).getValues();
    jv.forEach(function (r) {
      if (String(r[1]).trim().toLowerCase() === user) enviados[String(r[3]).trim()] = String(r[6] || 'Por revisar');
    });
  }
  return { email: user, casos: casos, enviados: enviados };
}

// Casos del usuario que ha iniciado sesión.
function getMisCasos() {
  var user = '';
  try { user = String(Session.getActiveUser().getEmail() || ''); } catch (_) {}
  return casosParaEmail_(user);
}

// Vista previa para ADMIN: ver lo que ve un usuario (solo lectura).
function getCasosDe(email) {
  var user = '';
  try { user = String(Session.getActiveUser().getEmail() || '').toLowerCase(); } catch (_) {}
  var admins = CONFIG.ADMINS.map(function (x) { return String(x).toLowerCase(); });
  if (admins.indexOf(user) === -1) return { email: email, casos: [], enviados: {}, error: 'Solo admins.' };
  return casosParaEmail_(email);
}

/* ===================== COMUNICACIONES (invitaciones) ================== */
function esAdminActual_() {
  var u = '';
  try { u = String(Session.getActiveUser().getEmail() || '').toLowerCase(); } catch (_) {}
  return CONFIG.ADMINS.map(function (x) { return String(x).toLowerCase(); }).indexOf(u) !== -1;
}

// Reinicio controlado para pruebas. Nunca elimina la carpeta raíz ni las hojas
// de referencia (Permisos, diccionario y Justificados). Las carpetas generadas
// se envían a la papelera de Drive, por lo que pueden recuperarse.
function reiniciarDatosPruebaSOX(confirmacion) {
  if (!esAdminActual_()) return { ok: false, mensaje: 'Solo los administradores pueden reiniciar datos.' };
  if (String(confirmacion || '') !== 'REINICIAR') return { ok: false, mensaje: 'Confirmación incorrecta. No se modificó nada.' };
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (_) { return { ok: false, mensaje: 'Hay otro proceso en curso. Intenta nuevamente.' }; }
  try {
    var raiz = obtenerCarpetaDestino_(), carpetas = 0, archivos = 0;
    var itC = raiz.getFolders();
    while (itC.hasNext()) {
      var carpeta = itC.next(), nombre = carpeta.getName();
      if (/^Q[1-4]\s+\d{4}$/i.test(nombre) || nombre === 'Sin periodo') {
        carpeta.setTrashed(true); carpetas++;
      }
    }
    var itF = raiz.getFiles();
    while (itF.hasNext()) {
      var archivo = itF.next();
      if (String(archivo.getName()).indexOf(CONFIG.PREFIJO_REPORTE_ENVIO) === 0) {
        archivo.setTrashed(true); archivos++;
      }
    }

    var panel = obtenerPanel_();
    limpiarResumen_(panel); limpiarCorreos_(panel);
    [CONFIG.HOJA_REPORTE, CONFIG.HOJA_HISTORIAL, CONFIG.HOJA_HISTORIAL_COMUNICACIONES, CONFIG.HOJA_JUSTIFICACIONES]
      .forEach(function (nombreHoja) {
        var sh = hojaPorNombre_(panel, nombreHoja);
        if (sh && panel.getSheets().length > 1) panel.deleteSheet(sh);
      });
    prepararHojaJustificaciones_(panel);
    SpreadsheetApp.flush();
    return { ok: true, carpetasPapelera: carpetas, archivosPapelera: archivos,
      mensaje: 'Datos de prueba reiniciados. La matriz, el diccionario y los justificados se conservaron.' };
  } catch (e) {
    return { ok: false, mensaje: 'No se pudo completar el reinicio: ' + e.message };
  } finally { try { lock.releaseLock(); } catch (_) {} }
}
function urlApp_() {
  try { return ScriptApp.getService().getUrl() || ''; } catch (_) { return ''; }
}

// HTML del correo de invitación al portal (branding LATAM, compatible con email).
function correoInvitacionHTML_(email, nCasos, url) {
  var nombre = String(email).split('@')[0].replace(/[._]/g, ' ');
  return '' +
  '<div style="font-family:Arial,Helvetica,sans-serif;background:#f4f4fa;padding:24px">' +
    '<div style="max-width:560px;margin:0 auto;background:#fff;border-radius:14px;overflow:hidden;border:1px solid #e7e7f0">' +
      '<div style="background:#1B0A5A;padding:22px 26px;color:#fff">' +
        '<span style="font-weight:800;font-size:22px;letter-spacing:2px">LATAM</span>' +
        '<div style="font-size:15px;font-weight:700;margin-top:6px">Control SOX de Tarifas</div>' +
      '</div>' +
      '<div style="padding:26px">' +
        '<p style="font-size:15px;color:#1a1a2e;margin:0 0 12px">Hola <b style="text-transform:capitalize">' + nombre + '</b>,</p>' +
        '<p style="font-size:14px;color:#444;line-height:1.6;margin:0 0 16px">' +
          'En la revisión SOX de tarifas del trimestre se identificaron <b style="color:#ED1650">' + nCasos +
          ' caso(s)</b> donde apareces y requieren tu <b>justificación</b>.</p>' +
        '<p style="font-size:14px;color:#444;line-height:1.6;margin:0 0 22px">' +
          'Ingresa a tu portal para verlos. La justificación formal la das <b>respondiendo el correo original</b> con la evidencia (aprobación del director).</p>' +
        '<a href="' + url + '" style="display:inline-block;background:#ED1650;color:#fff;text-decoration:none;font-weight:700;font-size:15px;padding:13px 26px;border-radius:999px">Abrir mi portal →</a>' +
        '<p style="font-size:12px;color:#8a8aa0;margin:24px 0 0">Si el botón no funciona, copia este enlace:<br>' + url + '</p>' +
      '</div>' +
      '<div style="background:#faf9ff;padding:14px 26px;font-size:11px;color:#8a8aa0;border-top:1px solid #eee">Control SOX de Tarifas · Cargo &amp; Pricing</div>' +
    '</div>' +
  '</div>';
}

function previewCorreoInvitacion(email) {
  if (!esAdminActual_()) return { ok: false, mensaje: 'Solo admins.' };
  var url = urlApp_();
  var casos = casosParaEmail_(email).casos.length;
  return { ok: true, asunto: 'Revisión SOX de tarifas — tienes ' + casos + ' caso(s) por justificar',
    html: correoInvitacionHTML_(email, casos, url), casos: casos, url: url };
}

function enviarInvitacion(email) {
  return { ok: false, mensaje: 'El envío individual está deshabilitado. Usa el correo general trimestral.' };
}

// Lista de usuarios con casos (para enviar invitaciones masivas).
function getUsuariosConCasos() {
  if (!esAdminActual_()) return { ok: false, mensaje: 'Solo admins.', usuarios: [] };
  var panel = obtenerPanel_();
  var resumen = hojaPorNombre_(panel, CONFIG.HOJA_RESUMEN);
  var m = {};
  if (resumen && resumen.getLastRow() > 1) {
    var d = resumen.getRange(2, 1, resumen.getLastRow() - 1, resumen.getLastColumn()).getValues();
    d.forEach(function (r) {
      [String(r[8]).trim(), String(r[6]).trim()].forEach(function (e) {
        if (e && e !== 'NULL' && e.indexOf('@') !== -1) m[e] = (m[e] || 0) + 1;
      });
    });
  }
  var usuarios = Object.keys(m).map(function (e) { return { email: e, casos: m[e] }; })
    .sort(function (a, b) { return b.casos - a.casos; });
  return { ok: true, usuarios: usuarios };
}

function enviarInvitacionesTodos() {
  return enviarCorreoGeneralTrimestral(periodoOperativoActual_());
}

// Calendario operativo: recordatorio mensual el día 15 y cierre por trimestre.
function obtenerCalendarioComunicacionesSOX() {
  if (!esAdminActual_()) return { ok: false, mensaje: 'Solo admins.' };
  var hoy = new Date(), y = hoy.getFullYear(), m = hoy.getMonth(), q = Math.floor(m / 3) + 1;
  var proximo = new Date(y, m, 15, 9, 0, 0);
  if (hoy.getTime() > proximo.getTime()) proximo = new Date(y, m + 1, 15, 9, 0, 0);
  var cierre = new Date(y, q * 3, 0, 23, 59, 59);
  var destinatarios = destinatariosReporte_();
  return {
    ok: true, periodo: 'Q' + q + ' ' + y,
    proximoRecordatorio: fechaDrive_(proximo), cierreTrimestre: fechaDrive_(cierre),
    destinatarios: destinatarios, totalDestinatarios: destinatarios.length,
    recordatorioActivo: ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'enviarRecordatorioMatriz'; })
  };
}

function periodoOperativoActual_() {
  var d = new Date(); return 'Q' + (Math.floor(d.getMonth() / 3) + 1) + ' ' + d.getFullYear();
}

function destinatariosReporte_() {
  var base = (CONFIG.DESTINATARIOS_REPORTE && CONFIG.DESTINATARIOS_REPORTE.length)
    ? CONFIG.DESTINATARIOS_REPORTE : (CONFIG.FOCALS_EMAILS || []);
  var vistos = {}, out = [];
  base.forEach(function (x) {
    var e = String(x || '').trim().toLowerCase();
    if (e && e.indexOf('@') !== -1 && !vistos[e]) { vistos[e] = true; out.push(e); }
  });
  return out;
}

function resumenPeriodoCorreo_(dashboard, periodo) {
  var casos = (dashboard.casos || []).filter(function (c) {
    if (periodo && c.periodo !== periodo && c.periodo !== 'Período actual') return false;
    var ap = String(c.aprobador || '').trim(), sol = String(c.solicitante || '').trim();
    if (!ap || ap.indexOf('@') === -1) return false;
    if (c.categoria === 'Aprobador no autorizado') return true;
    return c.categoria === 'Autoaprobación' && !!sol && sol.indexOf('@') !== -1 && ap.toLowerCase() === sol.toLowerCase();
  });
  var noAut = {}, auto = {};
  casos.forEach(function (c) {
    var mapa = c.categoria === 'Aprobador no autorizado' ? noAut : auto;
    var persona = c.aprobador;
    mapa[persona] = (mapa[persona] || 0) + 1;
  });
  function top(m) { return Object.keys(m).map(function (k) { return { persona: k, registros: m[k] }; })
    .sort(function (a, b) { return b.registros - a.registros; }).slice(0, 10); }
  return {
    total: casos.length,
    noAutorizados: casos.filter(function (c) { return c.categoria === 'Aprobador no autorizado'; }).length,
    autoAprobaciones: casos.filter(function (c) { return c.categoria === 'Autoaprobación'; }).length,
    topNoAutorizados: top(noAut), topAutoAprobaciones: top(auto)
  };
}

function escapeHtml_(v) {
  return String(v == null ? '' : v).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}

function correoGeneralTrimestralHTML_(periodo, resumen, reporteUrl, portalUrl, carpetaUrl) {
  function filas(items) {
    if (!items.length) return '<tr><td colspan="2" style="padding:10px;color:#16864b">Sin registros</td></tr>';
    return items.map(function (x) { return '<tr><td style="padding:8px;border-bottom:1px solid #eee">' + escapeHtml_(x.persona) +
      '</td><td style="padding:8px;border-bottom:1px solid #eee;text-align:right;font-weight:700">' + x.registros + '</td></tr>'; }).join('');
  }
  return '<div style="font-family:Arial,sans-serif;background:#f4f4fa;padding:24px;color:#202033">' +
    '<div style="max-width:720px;margin:auto;background:#fff;border:1px solid #e6e5ef;border-radius:14px;overflow:hidden">' +
    '<div style="background:#0f004f;color:#fff;padding:22px 26px"><b style="font-size:22px;letter-spacing:1px">LATAM</b>' +
    '<div style="margin-top:7px;font-size:15px;font-weight:700">Control SOX de Tarifas - ' + escapeHtml_(periodo) + '</div></div>' +
    '<div style="padding:25px"><p>Hola equipo,</p><p style="line-height:1.6">Finalizamos el procesamiento trimestral. Se identificaron <b>' + resumen.total +
    ' registros críticos</b> en la sábana consolidada. Este es el único correo general del período; cada usuario podrá ingresar al portal y visualizar únicamente los registros asociados a su cuenta.</p>' +
    '<div style="display:flex;gap:10px;flex-wrap:wrap;margin:20px 0"><div style="background:#feecef;padding:12px 16px;border-radius:9px"><b>' + resumen.noAutorizados +
    '</b><br><span style="font-size:12px">Aprobadores no autorizados</span></div><div style="background:#fff3df;padding:12px 16px;border-radius:9px"><b>' + resumen.autoAprobaciones +
    '</b><br><span style="font-size:12px">Autoaprobaciones</span></div></div>' +
    '<h3 style="color:#0f004f;font-size:14px">1. Aprobadores no autorizados</h3><table style="width:100%;border-collapse:collapse;font-size:12px">' + filas(resumen.topNoAutorizados) + '</table>' +
    '<h3 style="color:#0f004f;font-size:14px;margin-top:20px">2. Prueba de consistencia</h3><table style="width:100%;border-collapse:collapse;font-size:12px">' + filas(resumen.topAutoAprobaciones) + '</table>' +
    '<div style="margin-top:24px"><a href="' + escapeHtml_(portalUrl) + '" style="display:inline-block;background:#ed1650;color:#fff;text-decoration:none;font-weight:700;padding:12px 20px;border-radius:8px">Consultar mis registros</a></div>' +
    '<p style="font-size:12px;line-height:1.7;color:#666;margin-top:20px">Reporte trimestral consolidado: ' +
    (reporteUrl ? '<a href="' + escapeHtml_(reporteUrl) + '">abrir Sheet</a>' : '<b>pendiente de generar</b>') +
    '<br>Repositorio de evidencias: <a href="' + escapeHtml_(carpetaUrl) + '">abrir Drive</a></p>' +
    '<p style="font-size:12px;color:#666">Las justificaciones deben incluir la cadena de correos y, cuando corresponda, la aprobación de un jefe o director.</p></div></div></div>';
}

function previewCorreoGeneralTrimestral(periodo) {
  if (!esAdminActual_()) return { ok: false, mensaje: 'Solo admins.' };
  var p = String(periodo || periodoOperativoActual_());
  var d = obtenerDashboardSOX();
  if (!d.ok) return d;
  var resumen = resumenPeriodoCorreo_(d, p), destinatarios = destinatariosReporte_();
  var reporte = obtenerReporteEnvioExistente_(p);
  var portalUrl = urlApp_() + '?view=portal';
  var asunto = 'Control SOX de Tarifas - Resultados ' + p + ' y solicitud de evidencias';
  return { ok: true, periodo: p, asunto: asunto, destinatarios: destinatarios, resumen: resumen,
    reporteGenerado: !!reporte, reporteUrl: reporte ? reporte.getUrl() : '',
    html: correoGeneralTrimestralHTML_(p, resumen, reporte ? reporte.getUrl() : '', portalUrl, d.carpetaUrl) };
}

function prepararHistorialComunicaciones_(panel) {
  var hoja = hojaPorNombre_(panel, CONFIG.HOJA_HISTORIAL_COMUNICACIONES);
  if (!hoja) {
    hoja = panel.insertSheet(CONFIG.HOJA_HISTORIAL_COMUNICACIONES);
    hoja.getRange(1, 1, 1, 6).setValues([['Fecha', 'Tipo', 'Período', 'Destinatarios', 'Asunto', 'Registros']])
      .setFontWeight('bold').setBackground('#1B0A5A').setFontColor('#ffffff');
    hoja.setFrozenRows(1);
  }
  return hoja;
}

function enviarCorreoGeneralTrimestral(periodo) {
  if (!esAdminActual_()) return { ok: false, mensaje: 'Solo admins.' };
  var prev = previewCorreoGeneralTrimestral(periodo);
  if (!prev.ok) return prev;
  if (!prev.destinatarios.length) return { ok: false, mensaje: 'Configura DESTINATARIOS_REPORTE antes de enviar.' };
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (_) { return { ok: false, mensaje: 'Hay otro envío en curso. Intenta nuevamente.' }; }
  try {
    var panel = obtenerPanel_(), historial = prepararHistorialComunicaciones_(panel);
    if (historial.getLastRow() > 1) {
      var anteriores = historial.getRange(2, 2, historial.getLastRow() - 1, 2).getDisplayValues();
      var repetido = anteriores.some(function (r) { return r[0] === 'Correo general trimestral' && r[1] === prev.periodo; });
      if (repetido) return { ok: false, mensaje: 'El correo general de ' + prev.periodo + ' ya fue enviado. No se realizó un segundo envío.' };
    }
    var reporte = generarReporteEnvioTrimestral_(prev.periodo);
    var adjunto = exportarSpreadsheetBlob_(reporte.spreadsheet, 'Reporte_SOX_' + prev.periodo.replace(/\s+/g, '_') + '.xlsx');
    var html = correoGeneralTrimestralHTML_(prev.periodo, prev.resumen, reporte.url, urlApp_() + '?view=portal', obtenerCarpetaDestino_().getUrl());
    var opcionesCorreo = {
      to: prev.destinatarios.join(','), subject: prev.asunto,
      body: 'Resultados del control SOX ' + prev.periodo + '. Consulte el portal para revisar los registros asociados a su cuenta.',
      htmlBody: html
    };
    // Se deja margen frente al límite de adjuntos de Gmail. Si el libro supera
    // 20 MB, el correo conserva el enlace al mismo reporte sin adjuntarlo.
    if (adjunto.getBytes().length <= 20 * 1024 * 1024) opcionesCorreo.attachments = [adjunto];
    MailApp.sendEmail(opcionesCorreo);
    historial.appendRow([
      new Date(), 'Correo general trimestral', prev.periodo, prev.destinatarios.join(', '), prev.asunto, prev.resumen.total
    ]);
    return { ok: true, enviados: 1, destinatarios: prev.destinatarios.length, periodo: prev.periodo,
      reporteUrl: reporte.url, registros: reporte.registros };
  } catch (e) { return { ok: false, mensaje: 'No se pudo enviar el correo general: ' + e.message }; }
  finally { try { lock.releaseLock(); } catch (_) {} }
}

/* ===================== FIN COMUNICACIONES ============================= */
function prepararHojaJustificaciones_(panel) {
  var hoja = hojaPorNombre_(panel, CONFIG.HOJA_JUSTIFICACIONES);
  if (!hoja) {
    hoja = panel.insertSheet(CONFIG.HOJA_JUSTIFICACIONES);
    hoja.getRange(1, 1, 1, 7)
      .setValues([['Fecha', 'Email', 'Tipo', 'Rate Id', 'Justificación', 'Evidencia (link)', 'Estado']])
      .setFontWeight('bold').setBackground('#1B0A5A').setFontColor('#ffffff');
    hoja.setFrozenRows(1);
  }
  return hoja;
}

// El usuario envía su justificación (queda "Por revisar" y te notifica).
function enviarJustificacion(datos) {
  try {
    var user = '';
    try { user = String(Session.getActiveUser().getEmail() || '').toLowerCase(); } catch (_) {}
    if (!user) throw new Error('No pudimos identificar tu correo. Inicia sesión con tu cuenta LATAM.');
    var texto = (datos && datos.texto) ? String(datos.texto).trim() : '';
    if (!texto) texto = '(El usuario confirmó que lo justificará en el correo original)';
    var evidencia = (datos && datos.evidencia) ? String(datos.evidencia).trim() : '';
    if (evidencia && !/^https?:\/\//i.test(evidencia)) throw new Error('El enlace de evidencia debe comenzar por https:// o http://.');
    var panel = obtenerPanel_();
    var hoja = prepararHojaJustificaciones_(panel);
    var fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
    hoja.appendRow([fecha, user, datos.tipo || '', datos.rateId || '', texto, evidencia, 'Por revisar']);
    try {
      MailApp.sendEmail({
        to: CONFIG.ADMINS.join(','),
        subject: 'Justificación SOX recibida — ' + user + (datos.rateId ? ' (Rate ' + datos.rateId + ')' : ''),
        body: 'Usuario: ' + user + '\nTipo: ' + (datos.tipo || '') + '\nRate Id: ' + (datos.rateId || '') +
          '\n\nJustificación:\n' + texto + '\n\nEvidencia: ' + (evidencia || '(sin link)') +
          '\n\nRegistrado en la hoja "' + CONFIG.HOJA_JUSTIFICACIONES + '".'
      });
    } catch (e) { /* si no hay permiso de correo, igual queda registrado */ }
    return { ok: true };
  } catch (error) {
    return { ok: false, mensaje: error.message };
  }
}
function onOpen() {
  try {
    SpreadsheetApp.getUi().createMenu('🛡️ Control SOX')
      .addItem('Abrir app', 'mostrarDialogoCarga')
      .addItem('Configurar recordatorio mensual (día 15)', 'configurarRecordatorioQuincenal')
      .addToUi();
  } catch (e) {}
}
function mostrarDialogoCarga() {
  var html = HtmlService.createHtmlOutputFromFile('index').setWidth(560).setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Control SOX de Tarifas');
}

/* ======================= 3. CEREBRO (tu Sheet) ========================= */
function obtenerPanel_() {
  try { return SpreadsheetApp.openById(CONFIG.PANEL_SS_ID); }
  catch (e) { throw new Error('No pude abrir el Sheet base (PANEL_SS_ID). ¿Tienes acceso? ' + e.message); }
}

function obtenerInfoInicial() {
  var panel = obtenerPanel_();
  var permisos = hojaPorNombre_(panel, CONFIG.HOJA_PERMISOS);
  var n = 0;
  if (permisos && permisos.getLastRow() > 1) {
    var vals = permisos.getRange(2, CONFIG.PERM_COL_EMAIL, permisos.getLastRow() - 1, 1).getValues();
    n = vals.filter(function (x) { return String(x[0]).trim() !== ''; }).length;
  }
  var carpeta = obtenerCarpetaDestino_();
  return {
    version: CONFIG.APP_VERSION,
    panelUrl: panel.getUrl(),
    carpetaUrl: carpeta ? carpeta.getUrl() : '',
    permisos: n,
    tieneDiccionario: !!hojaPorNombre_(panel, CONFIG.HOJA_DICCIONARIO)
  };
}

// Mapa email -> {rol, region}
function cargarPermisos_(panel) {
  var map = {};
  var hoja = hojaPorNombre_(panel, CONFIG.HOJA_PERMISOS);
  if (hoja && hoja.getLastRow() > 1) {
    var ancho = Math.max(CONFIG.PERM_COL_EMAIL, CONFIG.PERM_COL_ROL, CONFIG.PERM_COL_REGION);
    var vals = hoja.getRange(2, 1, hoja.getLastRow() - 1, ancho).getValues();
    vals.forEach(function (r) {
      var e = String(r[CONFIG.PERM_COL_EMAIL - 1]).trim().toLowerCase();
      if (!e) return;
      map[e] = {
        rol: String(r[CONFIG.PERM_COL_ROL - 1] || '').trim(),
        region: String(r[CONFIG.PERM_COL_REGION - 1] || '').trim()
      };
    });
  }
  return map;
}

// Mapa POSTA -> {pais, zona, region2}
function cargarDiccionario_(panel) {
  var map = {};
  var hoja = hojaPorNombre_(panel, CONFIG.HOJA_DICCIONARIO);
  if (hoja && hoja.getLastRow() > 1) {
    var ancho = Math.max(CONFIG.DIC_COL_POSTA, CONFIG.DIC_COL_PAIS, CONFIG.DIC_COL_ZONA, CONFIG.DIC_COL_REGION2);
    var vals = hoja.getRange(2, 1, hoja.getLastRow() - 1, ancho).getValues();
    vals.forEach(function (r) {
      var p = String(r[CONFIG.DIC_COL_POSTA - 1]).trim().toUpperCase();
      if (!p) return;
      map[p] = {
        pais: String(r[CONFIG.DIC_COL_PAIS - 1] || '').trim(),
        zona: String(r[CONFIG.DIC_COL_ZONA - 1] || '').trim(),
        region2: String(r[CONFIG.DIC_COL_REGION2 - 1] || '').trim()
      };
    });
  }
  return map;
}

function prepararHojaPermisos() {
  var panel = obtenerPanel_();
  var hoja = hojaPorNombre_(panel, CONFIG.HOJA_PERMISOS);
  if (!hoja) {
    hoja = panel.insertSheet(CONFIG.HOJA_PERMISOS);
    hoja.getRange(1, 1, 1, 3).setValues([['Correo', 'Rol', 'Región']])
      .setFontWeight('bold').setBackground('#1B0A5A').setFontColor('#ffffff');
    hoja.setFrozenRows(1);
  }
  return panel.getUrl();
}

/* ========================= 4. ORQUESTADOR =============================== */
function iniciarLoteSOX(periodo) {
  var panel = obtenerPanel_();
  var p = String(periodo || '').trim();
  if (!rangoPeriodo_(p)) throw new Error('Selecciona explícitamente un período con formato Qn AAAA.');
  prepararHojaPermisos();
  prepararHojaJustificados_(panel);
  prepararResumenPeriodo_(panel, p);
  limpiarCorreos_(panel);
  var carpeta = obtenerSubcarpeta_(obtenerCarpetaDestino_(), p);
  return { ok: true, periodo: p, carpetaUrl: carpeta.getUrl(), panelUrl: panel.getUrl() };
}

// Consolida al final: genera las 2 tablas y deja un único borrador general.
function finalizarLoteSOX(periodo) {
  var panel = obtenerPanel_();
  construirReporteFinal_(panel);
  var correos = construirCorreos_(panel, periodo);
  var resumen = hojaPorNombre_(panel, CONFIG.HOJA_RESUMEN);
  var pendientes = resumen ? Math.max(0, resumen.getLastRow() - 1) : 0;
  registrarEjecucion_(panel, pendientes, correos);
  return { ok: true, panelUrl: panel.getUrl(), correos: correos, pendientes: pendientes };
}

// Reporte final = sábana consolidada con EXACTAMENTE 2 tablas (feedback jefa):
//  1) Aprobadores no autorizados   2) Consistencia (solicitante = aprobador).
function construirReporteFinal_(panel) {
  var resumen = hojaPorNombre_(panel, CONFIG.HOJA_RESUMEN);
  var rep = hojaPorNombre_(panel, CONFIG.HOJA_REPORTE) || panel.insertSheet(CONFIG.HOJA_REPORTE);
  rep.clear();
  var W = 6;
  function fila(a) { while (a.length < W) a.push(''); return a; }
  function setZona(m, k, z) { (m[k] = m[k] || {}); m[k].z = m[k].z || {}; if (z) m[k].z[z] = (m[k].z[z] || 0) + 1; }
  function zonasStr(z) {
    return Object.keys(z || {}).map(function (k) { return k + ' (' + z[k] + ')'; }).slice(0, 4).join(', ');
  }

  var noAut = {}, cons = {};
  if (resumen && resumen.getLastRow() > 1) {
    var d = resumen.getRange(2, 1, resumen.getLastRow() - 1, resumen.getLastColumn()).getValues();
    // 0 Tipo,1 RateId,2 Región,3 Origen,4 Zona,5 Destino,6 Solic,7 RolSolic,8 Aprob,9 RolAprob,10 Resultado
    d.forEach(function (r) {
      var res = String(r[10]), ap = String(r[8]).trim(), zona = String(r[4]).trim(), tipo = String(r[0]);
      if (res.indexOf('sin autorización') !== -1) {
        var a = noAut[ap] = noAut[ap] || { rol: r[9], count: 0, z: {}, tipos: {} };
        a.count++; if (zona) a.z[zona] = (a.z[zona] || 0) + 1; a.tipos[tipo] = 1;
      } else if (res.indexOf('duplicidad') !== -1) {
        var c = cons[ap] = cons[ap] || { rol: r[9] || r[7], count: 0, z: {}, tipos: {} };
        c.count++; if (zona) c.z[zona] = (c.z[zona] || 0) + 1; c.tipos[tipo] = 1;
      }
    });
  }

  var out = [];
  out.push(fila(['REPORTE FINAL SOX — inconsistencias confirmadas']));
  out.push(fila([]));
  var f1 = out.length + 1;
  out.push(fila(['1) APROBADORES NO AUTORIZADOS', 'Rol conocido', '# casos', 'Zonas afectadas', 'Tipos de tarifa']));
  var arr1 = Object.keys(noAut).map(function (k) { var a = noAut[k]; return [k, a.rol || '(externo)', a.count, zonasStr(a.z), Object.keys(a.tipos).join(', ')]; })
    .sort(function (a, b) { return b[2] - a[2]; });
  if (arr1.length === 0) out.push(fila(['✅ Ninguno'])); else arr1.forEach(function (r) { out.push(fila(r)); });
  out.push(fila([]));
  var f2 = out.length + 1;
  out.push(fila(['2) CONSISTENCIA — solicitante = aprobador', 'Rol conocido', '# casos', 'Zonas afectadas', 'Tipos de tarifa']));
  var arr2 = Object.keys(cons).map(function (k) { var c = cons[k]; return [k, c.rol || '(sin rol)', c.count, zonasStr(c.z), Object.keys(c.tipos).join(', ')]; })
    .sort(function (a, b) { return b[2] - a[2]; });
  if (arr2.length === 0) out.push(fila(['✅ Ninguno'])); else arr2.forEach(function (r) { out.push(fila(r)); });

  rep.getRange(1, 1, out.length, W).setValues(out);
  rep.getRange(1, 1, 1, W).setFontWeight('bold').setFontSize(13).setFontColor('#1B0A5A');
  [f1, f2].forEach(function (f) { rep.getRange(f, 1, 1, W).setFontWeight('bold').setBackground('#990000').setFontColor('#ffffff'); });
  rep.setColumnWidth(1, 300); rep.setColumnWidth(4, 240); rep.setColumnWidth(5, 200);
  rep.setFrozenRows(1);
}

/* Recordatorio general mensual de actualización de matriz (día 15). */
function configurarRecordatorioQuincenal() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'enviarRecordatorioMatriz') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('enviarRecordatorioMatriz').timeBased().onMonthDay(15).atHour(9).create();
  return 'Recordatorio mensual activado (día 15, 9:00).';
}
function enviarRecordatorioMatriz() {
  var panel = obtenerPanel_();
  var para = (CONFIG.FOCALS_EMAILS && CONFIG.FOCALS_EMAILS.length) ? CONFIG.FOCALS_EMAILS.join(',') : CONFIG.RESPONSABLE_SOX;
  var asunto = 'Recordatorio mensual - Actualización de la matriz de aprobadores (SOX)';
  var cuerpo = 'Hola,\n\nRecordatorio para actualizar la matriz de aprobadores SOX (SAM / NAM / EUR) con la información del último mes cerrado.\n\n' +
    'Matriz: ' + panel.getUrl() + '\n\nResponsable del proceso SOX: ' + CONFIG.RESPONSABLE_SOX + '\n\nGracias.';
  MailApp.sendEmail({ to: para, cc: CONFIG.RESPONSABLE_SOX, subject: asunto, body: cuerpo });
  return 'Recordatorio enviado a ' + para + '.';
}

function procesarArchivoIndividualSOX(datos) {
  var t0 = Date.now();
  try {
    if (!datos || !datos.base64 || !datos.nombre) throw new Error('No se recibió ningún archivo.');
    var extension = obtenerExtension_(datos.nombre);
    if (CONFIG.EXTENSIONES_PERMITIDAS.indexOf(extension) === -1) {
      throw new Error('Formato no válido (.' + extension + ').');
    }
    var bytes = Utilities.base64Decode(datos.base64);
    var blobPrincipal = Utilities.newBlob(bytes, datos.mimeType, datos.nombre);
    if (blobPrincipal.getBytes().length === 0) throw new Error('El archivo está vacío (0 bytes).');

    var blobs;
    if (extension === 'zip') {
      blobs = Utilities.unzip(blobPrincipal).filter(function (b) {
        return ['xlsx', 'xls', 'csv'].indexOf(obtenerExtension_(b.getName())) !== -1;
      });
      if (blobs.length === 0) throw new Error('El ZIP no contiene archivos válidos.');
    } else { blobs = [blobPrincipal]; }

    var carpetaBase = obtenerCarpetaDestino_();
    if (!carpetaBase) throw new Error('No se encontró la carpeta de Drive (CARPETA_DRIVE_ID).');
    var panel = obtenerPanel_();
    var permisos = cargarPermisos_(panel);
    var dicc = cargarDiccionario_(panel);
    var justificados = cargarJustificados_(panel);

    var periodoElegido = String(datos.periodo || '').trim();
    if (!rangoPeriodo_(periodoElegido)) throw new Error('No se recibió un período válido para esta carga.');
    var resultados = [], total = 0, reemplazados = 0, subUrl = carpetaBase.getUrl();
    for (var i = 0; i < blobs.length; i++) {
      var periodo = periodoElegido;
      var sub = obtenerSubcarpeta_(carpetaBase, periodo);
      subUrl = sub.getUrl();
      var r = procesarUnArchivo_(blobs[i], sub, permisos, dicc, periodo, panel, justificados);
      resultados.push(r.info);
      agregarResumen_(panel, r.resumen);
      total += r.resumen.length;
      reemplazados += r.info.reemplazados || 0;
    }
    return {
      ok: true, archivos: resultados, inconsistencias: total,
      periodo: periodoElegido, reemplazados: reemplazados,
      carpetaUrl: subUrl, panelUrl: panel.getUrl(),
      segundos: Math.round((Date.now() - t0) / 100) / 10
    };
  } catch (error) {
    return { ok: false, mensaje: (error && error.message) ? error.message : String(error),
      segundos: Math.round((Date.now() - t0) / 100) / 10 };
  }
}

/* ==================== 5. PROCESAR UN ARCHIVO ============================ */
function procesarUnArchivo_(blob, carpeta, permisos, dicc, periodo, panel, justificados) {
  var t0 = Date.now();
  var nombre = blob.getName();
  var base = quitarExtension_(nombre);
  var ext = obtenerExtension_(nombre);
  var reemplazados = CONFIG.REEMPLAZAR_EXISTENTES ? trashExistentes_(base, carpeta) : 0;

  var ss, matriz;
  if (ext === 'csv') {
    matriz = normalizar_(Utilities.parseCsv(blob.getDataAsString('UTF-8'), CONFIG.DELIMITADOR));
    ss = crearSheetEnCarpeta_(base, carpeta);
  } else {
    var conv = convertirAGoogleSheets_(blob, base, carpeta);
    ss = SpreadsheetApp.openById(conv.id);
    matriz = obtenerMatrizSeparada_(ss);
  }

  var hoja = ss.getSheets()[0];
  hoja.setName('Datos');
  escribirMatriz_(hoja, matriz);

  var tipo = detectarTipo_(nombre);
  var stats = aplicarValidacionesSOX_(hoja, matriz, tipo.clave, permisos, dicc, tipo.hoja, periodo, justificados);

  copiarHojasApoyo_(panel, ss);   // copia Permisos + diccionario para formulaciones posteriores

  return {
    info: {
      nombre: base, tipo: tipo.hoja, periodo: periodo, url: ss.getUrl(),
      reemplazados: reemplazados,
      filas: matriz.length - 1, columnas: matriz[0].length,
      noCumple: stats.noCumple, conNull: stats.conNull,
      segundos: Math.round((Date.now() - t0) / 100) / 10
    },
    resumen: stats.resumenRows
  };
}

/* ==================== 6. CONVERSIÓN / CREACIÓN ========================== */
function convertirAGoogleSheets_(blob, nombreBase, carpeta) {
  var recurso = { name: nombreBase, mimeType: MimeType.GOOGLE_SHEETS, parents: [carpeta.getId()] };
  try { return { id: Drive.Files.create(recurso, blob, { supportsAllDrives: true }).id }; }
  catch (e) { throw new Error('No se pudo convertir "' + nombreBase + '". ¿Drive API (v3) habilitada? ' + e.message); }
}
function crearSheetEnCarpeta_(nombreBase, carpeta) {
  var ss = SpreadsheetApp.create(nombreBase);
  try { DriveApp.getFileById(ss.getId()).moveTo(carpeta); } catch (e) {}
  return ss;
}
function trashExistentes_(nombreBase, carpeta) {
  var it = carpeta.getFilesByName(nombreBase), n = 0;
  while (it.hasNext()) { it.next().setTrashed(true); n++; }
  return n;
}
// Copia Permisos + diccionario postas desde el panel a cada Excel generado.
function copiarHojasApoyo_(panel, ss) {
  if (!CONFIG.COPIAR_HOJAS_APOYO || !panel) return;
  [CONFIG.HOJA_PERMISOS, CONFIG.HOJA_DICCIONARIO].forEach(function (nombre) {
    var origen = hojaPorNombre_(panel, nombre);
    if (!origen) return;
    var existente = hojaPorNombre_(ss, nombre);
    if (existente) ss.deleteSheet(existente);
    origen.copyTo(ss).setName(nombre);
  });
}

/* ============= 7. "TEXTO EN COLUMNAS" (respeta comillas) =============== */
function obtenerMatrizSeparada_(ss) {
  var hoja = ss.getSheets()[0];
  var uf = hoja.getLastRow(), uc = hoja.getLastColumn();
  if (uf === 0) throw new Error('El archivo convertido está vacío.');
  var primera = String(hoja.getRange(1, 1).getValue());
  var aglomerado = (uc === 1) || (primera.indexOf(CONFIG.DELIMITADOR) !== -1 && primera.indexOf('"') !== -1);
  if (!aglomerado) return hoja.getRange(1, 1, uf, uc).getValues();

  var colA = hoja.getRange(1, 1, uf, 1).getValues(), lineas = [];
  for (var i = 0; i < colA.length; i++) {
    var v = colA[i][0]; if (v === '' || v === null) continue; lineas.push(String(v));
  }
  if (lineas.length === 0) throw new Error('La Columna A no tiene datos.');
  return normalizar_(Utilities.parseCsv(lineas.join('\n'), CONFIG.DELIMITADOR));
}
function normalizar_(m) {
  var max = 0; m.forEach(function (f) { if (f.length > max) max = f.length; });
  for (var r = 0; r < m.length; r++) for (var c = 0; c < max; c++) {
    if (c >= m[r].length) m[r][c] = '';
    else if (CONFIG.NULL_A_VACIO && m[r][c] === 'NULL') m[r][c] = '';
  }
  return m;
}
function escribirMatriz_(hoja, matriz) {
  hoja.clear();
  var filas = matriz.length, cols = matriz[0].length;
  var rango = hoja.getRange(1, 1, filas, cols);
  rango.setNumberFormat('@');           // texto: preserva IDs largos y fechas
  rango.setValues(matriz);
  hoja.getRange(1, 1, 1, cols).setFontWeight('bold').setBackground('#1B0A5A').setFontColor('#ffffff');
  hoja.setFrozenRows(1);
}

/* ==================== 8. VALIDACIONES SOX (en código) ================== */
function aplicarValidacionesSOX_(hoja, matriz, claveTipo, permisos, dicc, tipoLabel, periodo, justificados) {
  var headers = matriz[0], n = matriz.length, base = headers.length, H = CONFIG.H;

  var colSol = buscarColumna_(headers, CONFIG.COL_SOLICITANTE);
  var colAp = buscarColumna_(headers, CONFIG.COL_APROBADOR);
  var colOri = buscarColumna_(headers, CONFIG.COL_ORIGEN);
  var colDes = buscarColumna_(headers, CONFIG.COL_DESTINO);
  var colNat = buscarColumna_(headers, CONFIG.COL_NATURALEZA);
  var iRate = buscarColumna_(headers, CONFIG.COL_RESUMEN.rateId);
  var iReg = buscarColumna_(headers, CONFIG.COL_RESUMEN.region);
  var idxCriticos = CONFIG.CAMPOS_CRITICOS
    .map(function (c) { return buscarColumna_(headers, [c]); }).filter(function (x) { return x > 0; });

  // Columna de fecha según el tipo (Spot/Street=creación; Promo/Contract=modificación).
  var esSpotStreet = (claveTipo === 'SPOT' || claveTipo === 'STRIP');
  var colFecha = esSpotStreet
    ? buscarColumna_(headers, CONFIG.COL_CREADO)
    : buscarColumna_(headers, CONFIG.COL_MODIFICADO);
  var rango = (CONFIG.FILTRAR_POR_PERIODO ? rangoPeriodo_(periodo) : null);

  var cabecerasSOX = [H.consist, H.apMat, H.reqMat, H.rolAp, H.rolReq, H.regAp, H.zona, H.pais, H.result, H.estado, H.periodo, H.nul, H.alc];
  var cSOX = base + 1;
  hoja.getRange(1, cSOX, 1, cabecerasSOX.length).setValues([cabecerasSOX])
    .setFontWeight('bold').setBackground('#188038').setFontColor('#ffffff');

  var salida = [], resumenRows = [], conNull = 0;
  var totDup = 0, totSinAut = 0, totCumple = 0, totAlc = 0, totFuera = 0, totJustif = 0, totPend = 0;
  var apStats = {}, solStats = {}, consStats = {};

  for (var r = 1; r < n; r++) {
    var fila = matriz[r];
    var ap = colAp ? String(fila[colAp - 1]).trim() : '';
    var sol = colSol ? String(fila[colSol - 1]).trim() : '';
    var apL = ap.toLowerCase(), solL = sol.toLowerCase();
    var apOk = ap && ap !== 'NULL', solOk = sol && sol !== 'NULL';

    var consist = (apOk && solOk && apL === solL) ? 'NO cumple' : 'cumple';
    var apEnMat = (apOk && permisos[apL]) ? 'Está en la matriz' : 'No está en la matriz';
    var reqEnMat = (solOk && permisos[solL]) ? 'Está en la matriz' : 'No está en la matriz';
    var rolAp = (apOk && permisos[apL]) ? permisos[apL].rol : '';
    var rolReq = (solOk && permisos[solL]) ? permisos[solL].rol : '';
    var regAp = (apOk && permisos[apL]) ? permisos[apL].region : '';

    var ori = colOri ? String(fila[colOri - 1]).trim().toUpperCase() : '';
    var d = dicc[ori];
    var zona = d ? d.region2 : '';
    var pais = d ? d.pais : '';

    var resultado;
    if (consist === 'NO cumple') { resultado = 'No cumple: duplicidad de roles'; totDup++; }
    else if (apEnMat === 'No está en la matriz') { resultado = 'No cumple: sin autorización'; totSinAut++; }
    else { resultado = 'Cumple'; totCumple++; }

    // ¿La fila pertenece al trimestre? (descartar arrastres de Q viejos)
    var enPeriodo = 'SÍ';
    if (rango && colFecha) {
      var f = parseFecha_(fila[colFecha - 1]);
      if (f) enPeriodo = (f >= rango.ini && f <= rango.fin) ? 'SÍ' : 'NO';
      else enPeriodo = 'sin fecha';
    }

    // ¿Ya justificado en un Q anterior?
    var rateId = iRate ? String(fila[iRate - 1]).trim() : '';
    var justificado = (rateId && justificados[rateId.toLowerCase()]) ||
      (apOk && justificados[apL]) || false;

    var faltantes = [];
    idxCriticos.forEach(function (ci) {
      var v = fila[ci - 1];
      if (v === 'NULL' || v === '' || v === null || v === undefined) faltantes.push(headers[ci - 1]);
    });
    var strNull = faltantes.join(', ');
    var enAlcNat = enAlcance_(claveTipo, colNat ? fila[colNat - 1] : '');
    var enAlc = enAlcNat && (enPeriodo !== 'NO');   // fuera de período => fuera de alcance

    // Estado (criterios del instructivo).
    var estado;
    if (resultado === 'Cumple') estado = 'Cumple';
    else if (justificado) { estado = 'Atípico justificado'; totJustif++; }
    else if (enAlc) { estado = 'Pendiente'; totPend++; }
    else estado = 'Fuera de alcance';
    if (enPeriodo === 'NO') totFuera++;

    salida.push([consist, apEnMat, reqEnMat, rolAp, rolReq, regAp, zona, pais, resultado, estado, enPeriodo, strNull, enAlc ? 'SÍ' : 'NO']);
    if (strNull) conNull++;
    if (enAlc) totAlc++;

    if (apOk) {
      var a = apStats[ap] || (apStats[ap] = { count: 0, noCumple: 0, enMat: (apEnMat === 'Está en la matriz'), rol: rolAp, zonas: {}, rates: [] });
      a.count++;
      if (resultado.indexOf('No cumple') === 0) a.noCumple++;
      if (zona) a.zonas[zona] = (a.zonas[zona] || 0) + 1;
      if (a.rates.length < 3 && rateId && rateId !== 'NULL' && a.rates.indexOf(rateId) === -1) a.rates.push(rateId);
    }
    if (solOk) {
      var s = solStats[sol] || (solStats[sol] = { count: 0, enMat: (reqEnMat === 'Está en la matriz'), rol: rolReq });
      s.count++;
    }
    // Consistencia: solicitante = aprobador (agrupado por persona).
    if (consist === 'NO cumple' && apOk) {
      var cs = consStats[ap] || (consStats[ap] = { count: 0, rol: rolAp, zonas: {}, rates: [] });
      cs.count++;
      if (zona) cs.zonas[zona] = (cs.zonas[zona] || 0) + 1;
      if (cs.rates.length < 3 && rateId && rateId !== 'NULL' && cs.rates.indexOf(rateId) === -1) cs.rates.push(rateId);
    }

    // Al Resumen solo entran los PENDIENTES (en alcance, no cumple, sin justificar).
    if (enAlc && resultado.indexOf('No cumple') === 0 && !justificado) {
      resumenRows.push([
        tipoLabel, rateId, iReg ? fila[iReg - 1] : '',
        ori, zona, colDes ? fila[colDes - 1] : '',
        sol, rolReq, ap, rolAp, resultado, strNull, periodo
      ]);
    }
  }

  var rng = hoja.getRange(2, cSOX, n - 1, cabecerasSOX.length);
  rng.setNumberFormat('@'); rng.setValues(salida);

  escribirAnalisis_(hoja.getParent(), tipoLabel,
    { total: n - 1, enAlcance: totAlc, cumple: totCumple, dup: totDup, sinAut: totSinAut,
      conNull: conNull, fuera: totFuera, justif: totJustif, pend: totPend },
    apStats, consStats);

  return { noCumple: totPend, conNull: conNull, resumenRows: resumenRows };
}

/* Análisis por archivo: SOLO las 2 tablas críticas (feedback jefa):
   (1) aprobadores no autorizados  (2) consistencia (solicitante = aprobador). */
function escribirAnalisis_(ss, tipoLabel, tot, apStats, consStats) {
  var vieja = hojaPorNombre_(ss, 'Análisis');
  if (vieja) ss.deleteSheet(vieja);
  var sh = ss.insertSheet('Análisis', 1);
  var W = 6;
  function zonasStr(z) {
    return Object.keys(z).map(function (k) { return [k, z[k]]; })
      .sort(function (a, b) { return b[1] - a[1]; }).slice(0, 4)
      .map(function (p) { return p[0] + ' (' + p[1] + ')'; }).join(', ');
  }
  function fila(arr) { while (arr.length < W) arr.push(''); return arr; }

  var out = [];
  out.push(fila(['ANÁLISIS SOX — ' + tipoLabel]));
  out.push(fila([]));
  out.push(fila(['RESUMEN']));
  out.push(fila(['Total filas', tot.total]));
  out.push(fila(['En alcance (período + tipo)', tot.enAlcance]));
  out.push(fila(['Fuera de período (arrastres descartados)', tot.fuera]));
  out.push(fila(['Cumple', tot.cumple]));
  out.push(fila(['Atípico justificado (Q anterior)', tot.justif]));
  out.push(fila(['PENDIENTES a revisar', tot.pend]));
  out.push(fila(['Filas con NULL crítico', tot.conNull]));
  out.push(fila([]));

  // TABLA 1: Aprobadores no autorizados.
  var noAut = Object.keys(apStats).map(function (k) {
    var a = apStats[k];
    return { email: k, rol: a.rol, enMat: a.enMat, count: a.count, zonas: zonasStr(a.zonas), rates: a.rates.join(' · ') };
  }).filter(function (a) { return !a.enMat; }).sort(function (a, b) { return b.count - a.count; });

  var f1 = out.length + 1;
  out.push(fila(['1) APROBADORES NO AUTORIZADOS', 'Rol conocido', '# aprobaciones', 'Zonas afectadas', 'Ejemplos Rate Id']));
  if (noAut.length === 0) out.push(fila(['✅ Ninguno: todos los aprobadores están en la matriz']));
  else noAut.forEach(function (a) { out.push(fila([a.email, a.rol || '(sin rol / externo)', a.count, a.zonas, a.rates])); });
  out.push(fila([]));

  // TABLA 2: Consistencia (solicitante = aprobador).
  var cons = Object.keys(consStats).map(function (k) {
    var c = consStats[k];
    return { email: k, rol: c.rol, count: c.count, zonas: zonasStr(c.zonas), rates: c.rates.join(' · ') };
  }).sort(function (a, b) { return b.count - a.count; });

  var f2 = out.length + 1;
  out.push(fila(['2) CONSISTENCIA — solicitante = aprobador', 'Rol conocido', '# casos', 'Zonas afectadas', 'Ejemplos Rate Id']));
  if (cons.length === 0) out.push(fila(['✅ Ninguno: no hay auto-aprobaciones']));
  else cons.forEach(function (c) { out.push(fila([c.email, c.rol || '(sin rol)', c.count, c.zonas, c.rates])); });

  sh.getRange(1, 1, out.length, W).setValues(out);
  sh.getRange(1, 1, 1, W).setFontWeight('bold').setFontSize(13).setFontColor('#1B0A5A');
  [3, f1, f2].forEach(function (f) {
    sh.getRange(f, 1, 1, W).setFontWeight('bold').setBackground('#1B0A5A').setFontColor('#ffffff');
  });
  if (noAut.length > 0) sh.getRange(f1 + 1, 1, 1, W).setBackground('#f4cccc').setFontWeight('bold');
  if (cons.length > 0) sh.getRange(f2 + 1, 1, 1, W).setBackground('#fce5cd').setFontWeight('bold');
  sh.setColumnWidth(1, 300); sh.setColumnWidth(4, 240); sh.setColumnWidth(5, 220);
  sh.setFrozenRows(1);
}

function enAlcance_(claveTipo, valorNaturaleza) {
  if (claveTipo === 'PROMO' || claveTipo === 'CONTRACT') return true;
  if (!valorNaturaleza) return false;
  return CONFIG.SPOT_EN_ALCANCE.indexOf(String(valorNaturaleza).trim().toUpperCase()) !== -1;
}

/* ==================== 9. RESUMEN CONSOLIDADO =========================== */
function limpiarResumen_(panel) {
  var hoja = hojaPorNombre_(panel, CONFIG.HOJA_RESUMEN);
  if (!hoja) hoja = panel.insertSheet(CONFIG.HOJA_RESUMEN);
  hoja.clear();
  var cab = ['Tipo', 'Rate Id', 'Región archivo', 'Origen', 'Zona origen', 'Destino',
    'Solicitante', 'Rol solic.', 'Aprobador', 'Rol aprob.', 'Resultado SOX', 'Campos NULL', 'Periodo'];
  hoja.getRange(1, 1, 1, cab.length).setValues([cab])
    .setFontWeight('bold').setBackground('#990000').setFontColor('#ffffff');
  hoja.setFrozenRows(1);
  return hoja;
}
// Al reprocesar un trimestre se reemplaza únicamente su consolidado. Los demás
// períodos permanecen disponibles para consulta histórica en la app.
function prepararResumenPeriodo_(panel, periodo) {
  var anterior = hojaPorNombre_(panel, CONFIG.HOJA_RESUMEN), conservar = [];
  if (anterior && anterior.getLastRow() > 1) {
    conservar = anterior.getRange(2, 1, anterior.getLastRow() - 1, Math.max(13, anterior.getLastColumn())).getValues()
      .filter(function (r) { return String(r[12] || '').trim() !== String(periodo); });
  }
  var hoja = limpiarResumen_(panel);
  if (conservar.length) hoja.getRange(2, 1, conservar.length, conservar[0].length).setValues(conservar);
  return hoja;
}
function agregarResumen_(panel, filas) {
  if (!filas || filas.length === 0) return;
  var lock = LockService.getScriptLock();
  try { lock.waitLock(30000); } catch (e) { /* seguimos igual */ }
  try {
    var hoja = hojaPorNombre_(panel, CONFIG.HOJA_RESUMEN) || limpiarResumen_(panel);
    var inicio = hoja.getLastRow() + 1;
    hoja.getRange(inicio, 1, filas.length, filas[0].length).setNumberFormat('@').setValues(filas);
    SpreadsheetApp.flush();
  } finally { try { lock.releaseLock(); } catch (e) {} }
}

/* ==================== 10. SUBCARPETAS POR TRIMESTRE ==================== */
function detectarPeriodo_(nombre) {
  var n = String(nombre);
  var mQ = n.match(/Q\s*([1-4])[\s_\-]*((?:20)?\d{2})/i);
  var anio = (n.match(/20\d{2}/) || [String(new Date().getFullYear())])[0];
  if (mQ) return 'Q' + mQ[1] + ' ' + (mQ[2].length === 2 ? '20' + mQ[2] : mQ[2]);
  var low = n.toLowerCase();
  if (/(jan|ene).*(mar)/.test(low)) return 'Q1 ' + anio;
  if (/(apr|abr).*(jun)/.test(low)) return 'Q2 ' + anio;
  if (/(jul).*(sep)/.test(low)) return 'Q3 ' + anio;
  if (/(oct).*(dec|dic)/.test(low)) return 'Q4 ' + anio;
  return 'Sin periodo';
}
function obtenerSubcarpeta_(carpetaBase, periodo) {
  var it = carpetaBase.getFoldersByName(periodo);
  return it.hasNext() ? it.next() : carpetaBase.createFolder(periodo);
}

// Explica el destino antes de cargar. No crea carpetas durante la consulta.
function obtenerDestinoPeriodoSOX(periodo) {
  if (!esAdminActual_()) return { ok: false, mensaje: 'Solo admins.' };
  var p = String(periodo || '').trim();
  if (!rangoPeriodo_(p)) return { ok: false, mensaje: 'Selecciona un período válido.' };
  var base = obtenerCarpetaDestino_(), it = base.getFoldersByName(p);
  if (it.hasNext()) {
    var existente = it.next();
    return { ok: true, periodo: p, existe: true, carpetaUrl: existente.getUrl(),
      mensaje: 'Se guardará en la carpeta existente "' + p + '".' };
  }
  return { ok: true, periodo: p, existe: false, carpetaUrl: base.getUrl(),
    mensaje: 'Se creará automáticamente la carpeta "' + p + '" dentro del repositorio SOX.' };
}

/* ======================== 11. UTILIDADES ============================== */
function obtenerExtension_(nombre) { var p = String(nombre).toLowerCase().split('.'); return p.length > 1 ? p.pop() : ''; }
function quitarExtension_(nombre) { return String(nombre).replace(/\.[^/.]+$/, ''); }
function detectarTipo_(nombreArchivo) {
  var n = String(nombreArchivo).toLowerCase();
  for (var i = 0; i < CONFIG.TIPOS.length; i++) {
    var t = CONFIG.TIPOS[i];
    for (var j = 0; j < t.patrones.length; j++) if (n.indexOf(t.patrones[j]) !== -1) return t;
  }
  return { clave: 'OTRO', hoja: 'Otro' };
}
// Normaliza nombres: quita espacios raros (dobles, no-break) y baja a minúsculas.
function normNombre_(s) {
  return String(s).replace(/\u00a0/g, ' ').replace(/\s+/g, ' ').trim().toLowerCase();
}
// Busca una pestaña por nombre de forma tolerante:
// 1) coincidencia normalizada exacta  2) coincidencia parcial (contiene).
function hojaPorNombre_(ss, nombre) {
  var obj = normNombre_(nombre);
  var hojas = ss.getSheets();
  for (var i = 0; i < hojas.length; i++) {
    if (normNombre_(hojas[i].getName()) === obj) return hojas[i];
  }
  for (var j = 0; j < hojas.length; j++) {
    var nom = normNombre_(hojas[j].getName());
    if (nom.indexOf(obj) !== -1 || obj.indexOf(nom) !== -1) return hojas[j];
  }
  // Último recurso: primera palabra clave (permiso / diccionario).
  var clave = obj.split(' ')[0];
  for (var k = 0; k < hojas.length; k++) {
    if (normNombre_(hojas[k].getName()).indexOf(clave) !== -1) return hojas[k];
  }
  return null;
}
/* =============== JUSTIFICADOS, PERÍODO, FECHAS, CORREOS =============== */
function prepararHojaJustificados_(panel) {
  var hoja = hojaPorNombre_(panel, CONFIG.HOJA_JUSTIFICADOS);
  if (!hoja) {
    hoja = panel.insertSheet(CONFIG.HOJA_JUSTIFICADOS);
    hoja.getRange(1, 1, 1, 4)
      .setValues([['Clave (Rate Id o email aprobador)', 'Trimestre', 'Justificación', 'Evidencia (link)']])
      .setFontWeight('bold').setBackground('#1B0A5A').setFontColor('#ffffff');
    hoja.setFrozenRows(1);
  }
  return hoja;
}
// Set de claves ya justificadas (Rate Id o email), en minúsculas.
function cargarJustificados_(panel) {
  var set = {};
  var hoja = hojaPorNombre_(panel, CONFIG.HOJA_JUSTIFICADOS);
  if (hoja && hoja.getLastRow() > 1) {
    var vals = hoja.getRange(2, 1, hoja.getLastRow() - 1, 1).getValues();
    vals.forEach(function (r) { var k = String(r[0]).trim().toLowerCase(); if (k) set[k] = true; });
  }
  return set;
}
// Rango de fechas del trimestre a partir de "Qn YYYY".
function rangoPeriodo_(periodo) {
  var m = String(periodo).match(/Q([1-4])\s+(\d{4})/);
  if (!m) return null;
  var q = parseInt(m[1], 10), y = parseInt(m[2], 10);
  var iniMes = (q - 1) * 3;                 // 0,3,6,9
  var ini = new Date(y, iniMes, 1, 0, 0, 0);
  var fin = new Date(y, iniMes + 3, 0, 23, 59, 59);  // último día del trimestre
  return { ini: ini, fin: fin };
}
// Parsea fechas tipo "2026-05-27 11:14:..." o "27/05/2026". Devuelve Date o null.
function parseFecha_(valor) {
  var s = String(valor).trim();
  if (!s || s === 'NULL') return null;
  var iso = s.match(/(\d{4})-(\d{2})-(\d{2})/);
  if (iso) return new Date(+iso[1], +iso[2] - 1, +iso[3]);
  var dmy = s.match(/(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (dmy) return new Date(+dmy[3], +dmy[2] - 1, +dmy[1]);
  var d = new Date(s);
  return isNaN(d.getTime()) ? null : d;
}
// Limpia la hoja de correos (encabezado).
function limpiarCorreos_(panel) {
  var hoja = hojaPorNombre_(panel, CONFIG.HOJA_CORREOS) || panel.insertSheet(CONFIG.HOJA_CORREOS);
  hoja.clear();
  hoja.getRange(1, 1, 1, 4).setValues([['Destinatarios generales', 'Asunto', 'Cuerpo del correo', '# registros']])
    .setFontWeight('bold').setBackground('#990000').setFontColor('#ffffff');
  hoja.setFrozenRows(1);
  return hoja;
}
// Mantiene la hoja histórica de apoyo, pero ahora genera una sola comunicación
// general. La segmentación individual se resuelve dentro del portal.
function construirCorreos_(panel, periodo) {
  var resumen = hojaPorNombre_(panel, CONFIG.HOJA_RESUMEN);
  var hojaC = limpiarCorreos_(panel);
  if (!resumen || resumen.getLastRow() < 2) return 0;
  var total = resumen.getLastRow() - 1, p = String(periodo || periodoOperativoActual_());
  var para = destinatariosReporte_().join(', ');
  hojaC.getRange(2, 1, 1, 4).setValues([[
    para || '(configurar DESTINATARIOS_REPORTE)',
    'Control SOX de Tarifas - Resultados ' + p + ' y solicitud de evidencias',
    'Único correo general del período. Adjunta el reporte consolidado y dirige a cada usuario al portal para consultar sus propios registros.',
    total
  ]]);
  return 1;
}

function buscarColumna_(headers, alias) {
  var lower = headers.map(function (h) { return String(h).trim().toLowerCase(); });
  for (var i = 0; i < alias.length; i++) {
    var idx = lower.indexOf(String(alias[i]).toLowerCase());
    if (idx !== -1) return idx + 1;
  }
  return 0;
}
function obtenerCarpetaDestino_() {
  if (CONFIG.CARPETA_DRIVE_ID) {
    try { return DriveApp.getFolderById(CONFIG.CARPETA_DRIVE_ID); }
    catch (e) { throw new Error('No pude abrir la carpeta (CARPETA_DRIVE_ID). ¿Acceso? ' + e.message); }
  }
  return DriveApp.getRootFolder();
}

/* ==================== 12. CENTRO OPERATIVO / DASHBOARD =============== */
// Registra cada cierre de lote para mantener trazabilidad operativa.
function registrarEjecucion_(panel, pendientes, correos) {
  var hoja = hojaPorNombre_(panel, CONFIG.HOJA_HISTORIAL);
  if (!hoja) {
    hoja = panel.insertSheet(CONFIG.HOJA_HISTORIAL);
    hoja.getRange(1, 1, 1, 5)
      .setValues([['Fecha', 'Ejecutado por', 'Casos pendientes', 'Correos preparados', 'Estado']])
      .setFontWeight('bold').setBackground('#1B0A5A').setFontColor('#ffffff');
    hoja.setFrozenRows(1);
  }
  var user = '';
  try { user = Session.getActiveUser().getEmail() || ''; } catch (_) {}
  hoja.appendRow([new Date(), user, pendientes || 0, correos || 0, 'Completado']);
  hoja.getRange(hoja.getLastRow(), 1).setNumberFormat('yyyy-mm-dd hh:mm');
}

function claveCaso_(email, rateId) {
  return String(email || '').trim().toLowerCase() + '|' + String(rateId || '').trim().toLowerCase();
}

// Fuente única para el dashboard administrativo. Lee las hojas existentes; no duplica datos.
function obtenerDashboardSOX() {
  if (!esAdminActual_()) return { ok: false, mensaje: 'Solo los administradores SOX pueden ver este tablero.' };

  var panel = obtenerPanel_();
  var resumen = hojaPorNombre_(panel, CONFIG.HOJA_RESUMEN);
  var justHoja = hojaPorNombre_(panel, CONFIG.HOJA_JUSTIFICACIONES);
  var estados = {}, evidencias = {};

  if (justHoja && justHoja.getLastRow() > 1) {
    var jv = justHoja.getRange(2, 1, justHoja.getLastRow() - 1, Math.max(7, justHoja.getLastColumn())).getValues();
    jv.forEach(function (r) {
      var key = claveCaso_(r[1], r[3]);
      estados[key] = String(r[6] || 'Por revisar');
      evidencias[key] = String(r[5] || '');
    });
  }

  var casos = [], noAut = {}, consistencia = {}, personas = {}, tipos = {}, zonas = {}, periodos = {};
  var metricas = { total: 0, paraEnvio: 0, datosIncompletos: 0, noAutorizados: 0, autoAprobaciones: 0, conEvidencia: 0, porRevisar: 0, aprobados: 0 };

  if (resumen && resumen.getLastRow() > 1) {
    var d = resumen.getRange(2, 1, resumen.getLastRow() - 1, resumen.getLastColumn()).getDisplayValues();
    d.forEach(function (r, idx) {
      var resultado = String(r[10] || '');
      var ap = String(r[8] || '').trim(), sol = String(r[6] || '').trim();
      if (!ap || ap.toUpperCase() === 'NULL') ap = '';
      if (!sol || sol.toUpperCase() === 'NULL') sol = '';
      var categoria = resultado.indexOf('sin autorización') !== -1 ? (ap.indexOf('@') !== -1 ? 'Aprobador no autorizado' : 'Dato incompleto') :
        (resultado.indexOf('duplicidad') !== -1 && ap && sol && ap.toLowerCase() === sol.toLowerCase() ? 'Autoaprobación' : 'Otro');
      var periodo = String(r[12] || 'Período actual').trim() || 'Período actual';
      var keyAp = claveCaso_(ap, r[1]), keySol = claveCaso_(sol, r[1]);
      var justificador = estados[keyAp] ? ap : (estados[keySol] ? sol : '');
      var estado = estados[keyAp] || estados[keySol] || 'Pendiente de usuario';
      var evidencia = evidencias[keyAp] || evidencias[keySol] || '';
      var caso = {
        id: idx + 2,
        tipo: r[0], rateId: r[1], region: r[2], origen: r[3], zona: r[4], destino: r[5],
        solicitante: sol, rolSolicitante: r[7], aprobador: ap, rolAprobador: r[9],
        resultado: resultado, categoria: categoria, camposNull: r[11], periodo: periodo,
        estado: estado, evidencia: evidencia, justificador: justificador
      };
      casos.push(caso);
      metricas.total++;
      if (categoria === 'Aprobador no autorizado') metricas.noAutorizados++;
      if (categoria === 'Autoaprobación') metricas.autoAprobaciones++;
      if (categoria === 'Dato incompleto') metricas.datosIncompletos++;
      if (categoria === 'Aprobador no autorizado' || categoria === 'Autoaprobación') {
        metricas.paraEnvio++;
        if (evidencia) metricas.conEvidencia++;
        if (estado === 'Por revisar') metricas.porRevisar++;
        if (estado === 'Aprobado') metricas.aprobados++;
      }
      if (ap.indexOf('@') !== -1) personas[ap.toLowerCase()] = true;
      if (sol.indexOf('@') !== -1) personas[sol.toLowerCase()] = true;
      if (r[0]) tipos[r[0]] = true;
      if (r[4]) zonas[r[4]] = true;
      periodos[periodo] = true;

      var mapa = categoria === 'Aprobador no autorizado' ? noAut : (categoria === 'Autoaprobación' ? consistencia : null);
      if (mapa) {
        var clave = ap || '(sin aprobador)';
        var item = mapa[clave] || (mapa[clave] = { email: clave, rol: r[9] || r[7] || '', casos: 0, zonas: {}, tipos: {}, ejemplos: [] });
        item.casos++;
        if (r[4]) item.zonas[r[4]] = true;
        if (r[0]) item.tipos[r[0]] = true;
        if (r[1] && item.ejemplos.length < 3 && item.ejemplos.indexOf(r[1]) === -1) item.ejemplos.push(r[1]);
      }
    });
  }

  function tabla(mapa) {
    return Object.keys(mapa).map(function (k) {
      var x = mapa[k];
      return { email: x.email, rol: x.rol || '(sin rol)', casos: x.casos,
        zonas: Object.keys(x.zonas).join(', '), tipos: Object.keys(x.tipos).join(', '), ejemplos: x.ejemplos.join(', ') };
    }).sort(function (a, b) { return b.casos - a.casos; });
  }

  var historial = hojaPorNombre_(panel, CONFIG.HOJA_HISTORIAL);
  var ultimaEjecucion = '';
  if (historial && historial.getLastRow() > 1) {
    ultimaEjecucion = historial.getRange(historial.getLastRow(), 1).getDisplayValue();
  } else {
    try { ultimaEjecucion = Utilities.formatDate(DriveApp.getFileById(panel.getId()).getLastUpdated(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm'); } catch (_) {}
  }

  return {
    ok: true,
    actualizado: ultimaEjecucion,
    tieneEjecuciones: !!(historial && historial.getLastRow() > 1),
    panelUrl: panel.getUrl(),
    carpetaUrl: obtenerCarpetaDestino_().getUrl(),
    metricas: metricas,
    personas: Object.keys(personas).length,
    filtros: { tipos: Object.keys(tipos).sort(), zonas: Object.keys(zonas).sort(), periodos: Object.keys(periodos).sort().reverse() },
    noAutorizados: tabla(noAut),
    consistencia: tabla(consistencia),
    casos: casos,
    recordatorioActivo: ScriptApp.getProjectTriggers().some(function (t) { return t.getHandlerFunction() === 'enviarRecordatorioMatriz'; })
  };
}

// Permite al administrador cerrar el ciclo de una evidencia sin editar el Sheet manualmente.
function actualizarEstadoJustificacion(email, rateId, estado) {
  if (!esAdminActual_()) return { ok: false, mensaje: 'Solo admins.' };
  var permitidos = ['Por revisar', 'Aprobado', 'Requiere ajuste'];
  if (permitidos.indexOf(estado) === -1) return { ok: false, mensaje: 'Estado no válido.' };
  var panel = obtenerPanel_(), hoja = prepararHojaJustificaciones_(panel);
  var vals = hoja.getLastRow() > 1 ? hoja.getRange(2, 1, hoja.getLastRow() - 1, 7).getValues() : [];
  var objetivo = claveCaso_(email, rateId), fila = 0;
  for (var i = vals.length - 1; i >= 0; i--) {
    if (claveCaso_(vals[i][1], vals[i][3]) === objetivo) { fila = i + 2; break; }
  }
  if (!fila) return { ok: false, mensaje: 'No existe una justificación enviada para este caso.' };
  hoja.getRange(fila, 7).setValue(estado);
  return { ok: true, estado: estado };
}

/* ============== 13. REPORTE TRIMESTRAL PARA DISTRIBUCIÓN ============= */
// Los Sheets creados desde cada archivo conservan su pestaña "Análisis" como
// soporte técnico. Este libro separado es el único entregable para distribución:
// no crea pestañas por persona; toda la segmentación se hace mediante filtros.
function nombreReporteEnvio_(periodo) {
  return CONFIG.PREFIJO_REPORTE_ENVIO + String(periodo || periodoOperativoActual_());
}

function carpetaPeriodoExistente_(periodo) {
  var base = obtenerCarpetaDestino_(), it = base.getFoldersByName(String(periodo));
  return it.hasNext() ? it.next() : null;
}

function obtenerReporteEnvioExistente_(periodo) {
  var carpeta = carpetaPeriodoExistente_(periodo);
  if (!carpeta) return null;
  var it = carpeta.getFilesByName(nombreReporteEnvio_(periodo));
  while (it.hasNext()) {
    var f = it.next();
    if (f.getMimeType() === MimeType.GOOGLE_SHEETS) return SpreadsheetApp.openById(f.getId());
  }
  return null;
}

function limpiarHojaReporte_(sh) {
  if (sh.getFilter()) sh.getFilter().remove();
  sh.getBandings().forEach(function (b) { b.remove(); });
  sh.clear();
  sh.clearConditionalFormatRules();
}

function mapaJustificacionesReporte_(panel) {
  var out = {}, sh = hojaPorNombre_(panel, CONFIG.HOJA_JUSTIFICACIONES);
  if (sh && sh.getLastRow() > 1) {
    sh.getRange(2, 1, sh.getLastRow() - 1, Math.max(7, sh.getLastColumn())).getDisplayValues().forEach(function (r) {
      out[claveCaso_(r[1], r[3])] = { estado: String(r[6] || 'Por revisar'), evidencia: String(r[5] || '') };
    });
  }
  return out;
}

function generarReporteEnvioTrimestral_(periodo) {
  var p = String(periodo || periodoOperativoActual_()).trim();
  if (!rangoPeriodo_(p)) throw new Error('El período debe tener el formato Qn AAAA.');
  var panel = obtenerPanel_(), resumen = hojaPorNombre_(panel, CONFIG.HOJA_RESUMEN);
  if (!resumen || resumen.getLastRow() < 2) throw new Error('No hay resultados consolidados para generar el reporte.');

  var carpeta = obtenerSubcarpeta_(obtenerCarpetaDestino_(), p);
  var ss = obtenerReporteEnvioExistente_(p);
  if (!ss) {
    ss = SpreadsheetApp.create(nombreReporteEnvio_(p));
    DriveApp.getFileById(ss.getId()).moveTo(carpeta);
  }
  var shResumen = hojaPorNombre_(ss, CONFIG.HOJA_REPORTE_ENVIO_RESUMEN);
  if (!shResumen) {
    var inicial = ss.getSheets()[0];
    if (ss.getSheets().length === 1 && inicial.getLastRow() <= 1 && inicial.getLastColumn() <= 1) {
      shResumen = inicial.setName(CONFIG.HOJA_REPORTE_ENVIO_RESUMEN);
    } else shResumen = ss.insertSheet(CONFIG.HOJA_REPORTE_ENVIO_RESUMEN, 0);
  }
  var shDetalle = hojaPorNombre_(ss, CONFIG.HOJA_REPORTE_ENVIO_DETALLE) || ss.insertSheet(CONFIG.HOJA_REPORTE_ENVIO_DETALLE);
  limpiarHojaReporte_(shResumen); limpiarHojaReporte_(shDetalle);

  var just = mapaJustificacionesReporte_(panel), noAut = {}, auto = {}, detalle = [], excluidos = 0;
  var datos = resumen.getRange(2, 1, resumen.getLastRow() - 1, resumen.getLastColumn()).getDisplayValues();
  datos.forEach(function (r) {
    var filaPeriodo = String(r[12] || '').trim();
    if (filaPeriodo !== p) return;
    var resultado = String(r[10] || ''), ap = String(r[8] || '').trim(), sol = String(r[6] || '').trim();
    var apValido = !!ap && ap.toUpperCase() !== 'NULL' && ap.indexOf('@') !== -1;
    var solValido = !!sol && sol.toUpperCase() !== 'NULL' && sol.indexOf('@') !== -1;
    var categoria = '';
    if (resultado.indexOf('sin autorización') !== -1) {
      if (!apValido) { excluidos++; return; }
      categoria = 'Aprobador no autorizado';
    } else if (resultado.indexOf('duplicidad') !== -1) {
      if (!apValido || !solValido || ap.toLowerCase() !== sol.toLowerCase()) { excluidos++; return; }
      categoria = 'Autoaprobación';
    } else return;

    var j = just[claveCaso_(ap, r[1])] || just[claveCaso_(sol, r[1])] || { estado: 'Pendiente de usuario', evidencia: '' };
    detalle.push([p, categoria, r[0], r[1], r[2], r[3], r[4], r[5], sol, r[7], ap, r[9], resultado, j.estado, j.evidencia]);
    var mapa = categoria === 'Aprobador no autorizado' ? noAut : auto;
    var x = mapa[ap.toLowerCase()] || (mapa[ap.toLowerCase()] = { persona: ap, rol: r[9] || r[7], casos: 0, zonas: {}, tipos: {} });
    x.casos++; if (r[4]) x.zonas[r[4]] = true; if (r[0]) x.tipos[r[0]] = true;
  });

  detalle.sort(function (a, b) { return a[1] === b[1] ? String(a[10]).localeCompare(String(b[10])) : String(a[1]).localeCompare(String(b[1])); });
  function tabla(m) {
    return Object.keys(m).map(function (k) { var x = m[k]; return [x.persona, x.rol || '(sin rol)', x.casos, Object.keys(x.zonas).join(', '), Object.keys(x.tipos).join(', ')]; })
      .sort(function (a, b) { return b[2] - a[2]; });
  }
  var t1 = tabla(noAut), t2 = tabla(auto), W = 5, out = [];
  function completa(a) { while (a.length < W) a.push(''); return a; }
  out.push(completa(['REPORTE SOX PARA ENVÍO — ' + p]));
  out.push(completa(['Documento único consolidado; no requiere pestañas por persona.']));
  out.push(completa(['Registros incluidos', detalle.length, 'Casos de datos incompletos excluidos del envío', excluidos]));
  out.push(completa([]));
  var h1 = out.length + 1;
  out.push(completa(['1. APROBADORES NO AUTORIZADOS', 'Rol', 'Registros', 'Zonas', 'Tipos']));
  if (t1.length) t1.forEach(function (r) { out.push(completa(r)); }); else out.push(completa(['Sin casos confirmados']));
  out.push(completa([]));
  var h2 = out.length + 1;
  out.push(completa(['2. PRUEBA DE CONSISTENCIA — solicitante = aprobador', 'Rol', 'Registros', 'Zonas', 'Tipos']));
  if (t2.length) t2.forEach(function (r) { out.push(completa(r)); }); else out.push(completa(['Sin casos confirmados']));
  shResumen.getRange(1, 1, out.length, W).setValues(out);
  shResumen.getRange(1, 1, 1, W).merge().setBackground('#1B0A5A').setFontColor('#ffffff').setFontWeight('bold').setFontSize(15);
  shResumen.getRange(2, 1, 1, W).merge().setFontColor('#666666').setFontStyle('italic');
  [h1, h2].forEach(function (n) { shResumen.getRange(n, 1, 1, W).setBackground('#ED1650').setFontColor('#ffffff').setFontWeight('bold'); });
  shResumen.setFrozenRows(3); shResumen.setColumnWidth(1, 330); shResumen.setColumnWidth(2, 220); shResumen.setColumnWidth(3, 100); shResumen.setColumnWidth(4, 210); shResumen.setColumnWidth(5, 150);

  var headers = ['Período', 'Hallazgo', 'Tipo', 'Rate ID', 'Región', 'Origen', 'Zona', 'Destino', 'Solicitante', 'Rol solicitante', 'Aprobador', 'Rol aprobador', 'Resultado del control', 'Estado evidencia', 'Link evidencia'];
  shDetalle.getRange(1, 1, 1, headers.length).setValues([headers]).setBackground('#1B0A5A').setFontColor('#ffffff').setFontWeight('bold');
  if (detalle.length) {
    shDetalle.getRange(2, 1, detalle.length, headers.length).setNumberFormat('@').setValues(detalle);
    shDetalle.getRange(1, 1, detalle.length + 1, headers.length).createFilter();
    shDetalle.getRange(1, 1, detalle.length + 1, headers.length).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY, true, false);
  }
  shDetalle.setFrozenRows(1); shDetalle.autoResizeColumns(1, headers.length);
  [2, 9, 10, 11, 12, 13, 15].forEach(function (c) { shDetalle.setColumnWidth(c, c === 13 ? 260 : 210); });
  SpreadsheetApp.flush();
  return { spreadsheet: ss, url: ss.getUrl(), id: ss.getId(), periodo: p, registros: detalle.length,
    noAutorizados: t1.reduce(function (s, r) { return s + r[2]; }, 0),
    autoAprobaciones: t2.reduce(function (s, r) { return s + r[2]; }, 0), excluidos: excluidos };
}

function generarReporteEnvioTrimestral(periodo) {
  if (!esAdminActual_()) return { ok: false, mensaje: 'Solo admins.' };
  try {
    var r = generarReporteEnvioTrimestral_(periodo);
    return { ok: true, periodo: r.periodo, url: r.url, registros: r.registros,
      noAutorizados: r.noAutorizados, autoAprobaciones: r.autoAprobaciones, excluidos: r.excluidos };
  } catch (e) { return { ok: false, mensaje: 'No se pudo generar el reporte para envío: ' + e.message }; }
}

function exportarSpreadsheetBlob_(ss, nombre) {
  var mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  var url = 'https://www.googleapis.com/drive/v3/files/' + ss.getId() + '/export?mimeType=' + encodeURIComponent(mime);
  var res = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
  if (res.getResponseCode() !== 200) throw new Error('Drive respondió ' + res.getResponseCode() + '.');
  return res.getBlob().setName(nombre);
}

function exportarReporteEnvioExcel(periodo) {
  if (!esAdminActual_()) return { ok: false, mensaje: 'Solo admins.' };
  try {
    var r = generarReporteEnvioTrimestral_(periodo), nombre = 'Reporte_SOX_' + r.periodo.replace(/\s+/g, '_') + '.xlsx';
    var blob = exportarSpreadsheetBlob_(r.spreadsheet, nombre);
    return { ok: true, nombre: nombre, mimeType: blob.getContentType(), base64: Utilities.base64Encode(blob.getBytes()), url: r.url };
  } catch (e) { return { ok: false, mensaje: 'No se pudo generar el Excel para envío: ' + e.message }; }
}

// Exportación interna del panel maestro; no es el archivo que se distribuye.
function exportarPanelExcel() {
  if (!esAdminActual_()) return { ok: false, mensaje: 'Solo admins.' };
  try {
    var panel = obtenerPanel_();
    var mime = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
    var url = 'https://www.googleapis.com/drive/v3/files/' + panel.getId() + '/export?mimeType=' + encodeURIComponent(mime);
    var res = UrlFetchApp.fetch(url, { headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() }, muteHttpExceptions: true });
    if (res.getResponseCode() !== 200) throw new Error('Drive respondió ' + res.getResponseCode() + '.');
    var fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmm');
    return { ok: true, nombre: 'Control_SOX_' + fecha + '.xlsx', mimeType: mime,
      base64: Utilities.base64Encode(res.getBlob().getBytes()) };
  } catch (e) { return { ok: false, mensaje: 'No se pudo generar el Excel: ' + e.message }; }
}

/* ==================== 13. EXPLORADOR DEL REPOSITORIO DRIVE ============ */
// Expone únicamente el contenido de la carpeta SOX configurada. No permite
// navegar por otras carpetas del Drive del usuario ni modificar archivos.
function listarRepositorioSOX(carpetaId) {
  if (!esAdminActual_()) return { ok: false, mensaje: 'Solo los administradores SOX pueden consultar el repositorio.' };
  try {
    var raiz = obtenerCarpetaDestino_();
    var carpeta = raiz;
    if (carpetaId && String(carpetaId) !== raiz.getId()) {
      carpeta = DriveApp.getFolderById(String(carpetaId));
      if (!carpetaPerteneceA_(carpeta, raiz.getId())) {
        return { ok: false, mensaje: 'La carpeta solicitada no pertenece al repositorio SOX.' };
      }
    }

    var carpetas = [], itC = carpeta.getFolders();
    while (itC.hasNext() && carpetas.length < 100) {
      var c = itC.next();
      carpetas.push({ id: c.getId(), nombre: c.getName(), url: c.getUrl(), actualizado: fechaDrive_(c.getLastUpdated()) });
    }
    carpetas.sort(function (a, b) { return b.nombre.localeCompare(a.nombre); });

    var archivos = [], itA = carpeta.getFiles();
    while (itA.hasNext() && archivos.length < 300) {
      var a = itA.next();
      archivos.push({
        id: a.getId(), nombre: a.getName(), url: a.getUrl(), mimeType: a.getMimeType(),
        tipo: tipoArchivoDrive_(a.getMimeType(), a.getName()), tamano: tamanoDrive_(a.getSize()),
        actualizado: fechaDrive_(a.getLastUpdated())
      });
    }
    archivos.sort(function (a, b) { return String(b.actualizado).localeCompare(String(a.actualizado)); });

    var migas = [{ id: raiz.getId(), nombre: 'Repositorio SOX' }];
    if (carpeta.getId() !== raiz.getId()) migas.push({ id: carpeta.getId(), nombre: carpeta.getName() });
    return {
      ok: true, raizId: raiz.getId(), carpetaId: carpeta.getId(), nombre: carpeta.getName(),
      url: carpeta.getUrl(), migas: migas, carpetas: carpetas, archivos: archivos,
      truncado: carpetas.length >= 100 || archivos.length >= 300
    };
  } catch (e) { return { ok: false, mensaje: 'No se pudo consultar Drive: ' + e.message }; }
}

function carpetaPerteneceA_(carpeta, raizId) {
  if (carpeta.getId() === raizId) return true;
  var actuales = [carpeta], vistos = {}, profundidad = 0;
  while (actuales.length && profundidad < 12) {
    var siguientes = [];
    for (var i = 0; i < actuales.length; i++) {
      var padres = actuales[i].getParents();
      while (padres.hasNext()) {
        var p = padres.next();
        if (p.getId() === raizId) return true;
        if (!vistos[p.getId()]) { vistos[p.getId()] = true; siguientes.push(p); }
      }
    }
    actuales = siguientes; profundidad++;
  }
  return false;
}

function fechaDrive_(fecha) {
  return Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
}
function tamanoDrive_(bytes) {
  var n = Number(bytes || 0);
  if (n < 1024) return n + ' B';
  if (n < 1048576) return Math.round(n / 1024) + ' KB';
  return (Math.round(n / 104857.6) / 10) + ' MB';
}
function tipoArchivoDrive_(mime, nombre) {
  var m = String(mime || ''), n = String(nombre || '').toLowerCase();
  if (m.indexOf('spreadsheet') !== -1 || /\.xlsx?$/.test(n)) return 'Hoja de cálculo';
  if (m.indexOf('pdf') !== -1 || /\.pdf$/.test(n)) return 'PDF';
  if (m.indexOf('zip') !== -1 || /\.zip$/.test(n)) return 'ZIP';
  if (m.indexOf('document') !== -1) return 'Documento';
  return 'Archivo';
}
