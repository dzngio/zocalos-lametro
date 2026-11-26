/**
 * Zócalos La Metro: API de sincronización en red
 * ------------------------------------------------
 * Pegá este archivo en el editor de Apps Script de la planilla
 * (Extensiones → Apps Script) y publicalo como Aplicación web:
 *   Ejecutar como: Yo
 *   Quién tiene acceso: Cualquier usuario
 *
 * Flujo:
 *   control.html  --POST (text/plain, JSON)-->  doPost  -> guarda estado + historial
 *   zocalo.html   --GET cada 1,5 s---------->  doGet   -> devuelve el estado vigente
 *
 * Para que sea rápido y no gaste cuota, el estado vigente se guarda en
 * CacheService. La planilla queda como registro visible y como respaldo si
 * la caché se vacía.
 */

/* ===================== CONFIGURACIÓN ===================== */

// Clave opcional para escribir. Si la completás, el panel tiene que tener la
// misma clave (engranaje → Clave de escritura). Recomendado: el panel es una
// página pública y cualquiera que conozca la URL del Web App podría mandar
// cambios a tu zócalo.
const TOKEN = '';

const SHEET_ESTADO = 'ESTADO_ACTUAL';
const SHEET_HISTORIAL = 'HISTORIAL';

const HEAD_ESTADO = [
  'timestamp', 'visible', 'tipoZocalo', 'colorPrincipal', 'volanta', 'titular',
  'subtitulo', 'ticker', 'qrActivo', 'qrTexto', 'qrDescuento', 'versionId', 'json'
];
const HEAD_HISTORIAL = [
  'fecha', 'accion', 'visible', 'tipoZocalo', 'colorPrincipal', 'volanta',
  'titular', 'subtitulo', 'qrActivo', 'versionId', 'origen'
];

const CACHE_ESTADO = 'estado';   // {versionId, timestamp, state}
const CACHE_BEAT = 'beat';       // último latido de OBS {id, shown, t}
const CACHE_TTL = 21600;         // 6 h, el máximo de CacheService

/* ===================== GET: leer estado ===================== */

/**
 * GET ?t=<cache-buster>
 *   &src=obs&id=<id>&shown=0|1   (lo manda zocalo.html: registra que OBS está vivo)
 *   &src=panel                    (lo manda control.html: recibe además el latido de OBS)
 */
function doGet(e) {
  const p = (e && e.parameter) || {};
  const cache = CacheService.getScriptCache();

  if (p.src === 'obs') {
    cache.put(CACHE_BEAT, JSON.stringify({
      id: String(p.id || '').slice(0, 12),
      shown: p.shown === '1',
      t: Date.now()
    }), 120);
  }

  let actual = cache.get(CACHE_ESTADO);
  if (!actual) {
    actual = leerEstadoDePlanilla_();
    if (actual) cache.put(CACHE_ESTADO, actual, CACHE_TTL);
  }

  const out = Object.assign({ ok: true, versionId: 0, timestamp: 0, state: null },
    actual ? JSON.parse(actual) : {});
  out.now = Date.now();

  if (p.src !== 'obs') {
    const beat = cache.get(CACHE_BEAT);
    out.obs = beat ? JSON.parse(beat) : null;
  }
  return json_(out);
}

/* ===================== POST: guardar estado ===================== */

/**
 * Body (text/plain para evitar el preflight CORS):
 *   { "accion": "ENTRAR" | "SALIR" | "ACTUALIZAR", "state": {...}, "token": "...", "origen": "panel" }
 */
function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(8000);

    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    if (TOKEN && body.token !== TOKEN) return json_({ ok: false, error: 'Clave de escritura incorrecta' });

    const s = body.state;
    if (!s || typeof s !== 'object') return json_({ ok: false, error: 'Falta el estado' });

    const accion = ['ENTRAR', 'SALIR', 'ACTUALIZAR'].indexOf(String(body.accion).toUpperCase()) >= 0
      ? String(body.accion).toUpperCase() : 'ACTUALIZAR';

    // versionId: contador incremental, el zócalo solo re-renderiza cuando cambia
    const props = PropertiesService.getScriptProperties();
    const versionId = Number(props.getProperty('versionId') || 0) + 1;
    props.setProperty('versionId', String(versionId));

    const ahora = new Date();
    const registro = { versionId: versionId, timestamp: ahora.getTime(), state: s };
    CacheService.getScriptCache().put(CACHE_ESTADO, JSON.stringify(registro), CACHE_TTL);

    const persona = s.modo === 'persona';
    const titular = persona ? s.nombre : s.titular;
    const subtitulo = persona ? [s.handle, s.rol].filter(String).join(' · ') : '';

    const ss = SpreadsheetApp.getActiveSpreadsheet();

    // Hoja 1: una sola fila con el estado vigente (fila 2)
    hoja_(ss, SHEET_ESTADO, HEAD_ESTADO)
      .getRange(2, 1, 1, HEAD_ESTADO.length)
      .setValues([[
        ahora, !!s.visible, txt_(s.modo), txt_(s.tema), txt_(s.kicker), txt_(titular),
        txt_(subtitulo), txt_(s.ticker), !!s.qrOn, txt_([s.qrL1, s.qrL2].filter(String).join(' / ')),
        txt_(s.qrPie), versionId, txt_(JSON.stringify(s))
      ]]);

    // Hoja 2: log de todo lo que pasó
    hoja_(ss, SHEET_HISTORIAL, HEAD_HISTORIAL).appendRow([
      ahora, accion, !!s.visible, txt_(s.modo), txt_(s.tema), txt_(s.kicker),
      txt_(titular), txt_(subtitulo), !!s.qrOn, versionId, txt_(body.origen || 'panel')
    ]);

    return json_({ ok: true, versionId: versionId, timestamp: registro.timestamp });
  } catch (err) {
    return json_({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (_) {}
  }
}

/* ===================== Utilidades ===================== */

/** Corré esta función una vez desde el editor para crear las hojas y autorizar el script. */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  hoja_(ss, SHEET_ESTADO, HEAD_ESTADO);
  hoja_(ss, SHEET_HISTORIAL, HEAD_HISTORIAL);
}

/** Borra la caché (si editaste a mano la fila de ESTADO_ACTUAL). */
function limpiarCache() {
  CacheService.getScriptCache().removeAll([CACHE_ESTADO, CACHE_BEAT]);
}

function hoja_(ss, nombre, encabezados) {
  let sh = ss.getSheetByName(nombre);
  if (!sh) sh = ss.insertSheet(nombre);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, encabezados.length).setValues([encabezados]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function leerEstadoDePlanilla_() {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_ESTADO);
  if (!sh || sh.getLastRow() < 2) return null;
  const fila = sh.getRange(2, 1, 1, HEAD_ESTADO.length).getValues()[0];
  const col = n => fila[HEAD_ESTADO.indexOf(n)];
  try {
    return JSON.stringify({
      versionId: Number(col('versionId')) || 0,
      timestamp: new Date(col('timestamp')).getTime() || 0,
      state: JSON.parse(String(col('json')).replace(/^'/, ''))
    });
  } catch (_) { return null; }
}

/**
 * Pasa a texto y evita que algo como "=IMPORTXML(...)" se guarde como fórmula
 * (el endpoint es público, así que no hay que confiar en lo que llega).
 */
function txt_(v) {
  const s = v == null ? '' : String(v).slice(0, 5000);
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
