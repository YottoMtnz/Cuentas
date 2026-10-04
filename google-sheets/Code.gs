/**
 * CUENTAS · Base de datos en Google Sheets, con usuarios
 * ------------------------------------------------------
 * Pega este archivo completo en Extensiones > Apps Script de tu hoja de
 * cálculo y sigue los pasos de CONFIGURAR-GOOGLE-SHEETS.txt.
 *
 * Cada usuario crea su cuenta desde la app (usuario + contraseña) y sus datos
 * quedan separados por usuario en las pestañas Movimientos, Metas y Config.
 * Las pestañas ocultas _usuarios y _sesiones guardan las cuentas (la contraseña
 * se guarda solo como huella, nunca en claro).
 */

// 1) CAMBIA ESTO: código para poder CREAR cuentas (8 a 64 caracteres: letras,
//    números, punto, guion o guion bajo). El mismo va en google-config.txt.
const CODIGO_REGISTRO = 'PON-AQUI-UN-CODIGO-LARGO';

const HOJA_MOV = 'Movimientos', HOJA_METAS = 'Metas', HOJA_CFG = 'Config';
const HOJA_USR = '_usuarios', HOJA_SES = '_sesiones';
const COLS_MOV = ['usuario', 'id', 'fecha', 'tipo', 'categoria', 'descripcion', 'monto', 'deducible', 'id_meta', 'ref_banco', 'bruto', 'retencion'];
const FMT_MOV = ['@', '0', '@', '@', '@', '@', '#,##0.00', '@', '0', '@', '#,##0.00', '#,##0.00'];
const COLS_META = ['usuario', 'id', 'nombre', 'objetivo'];
const FMT_META = ['@', '0', '@', '#,##0.00'];
const COLS_CFG = ['usuario', 'seccion', 'clave', 'valor'];
const FMT_CFG = ['@', '@', '@', '@'];
const COLS_USR = ['usuario', 'sal', 'huella', 'rev', 'creado'];
const FMT_USR = ['@', '@', '@', '0', '@'];
const COLS_SES = ['huella', 'usuario', 'vence'];
const FMT_SES = ['@', '@', '0'];
const TIPO_A_TXT = { i: 'Ingreso', g: 'Gasto', s: 'Ahorro' };
const CFG_CLAVES = ['cur', 'purl', 'theme'];
const TAX_CLAVES = ['year', 'fs', 'kids', 'od', 'n65', 'item', 'adj', 'w2', 'se', 'exp', 'oth', 'wh', 'est', 'st'];
const SESION_MS = 90 * 24 * 3600 * 1000;
const MAX_FALLOS = 5, BLOQUEO_S = 900;
const RE_USUARIO = /^[a-z0-9._-]{3,30}$/, RE_CLAVE = /^[0-9a-f]{64}$/;

/* ---------- Puntos de entrada ---------- */

function doGet() {
  return json_({ ok: true, app: 'cuentas' });
}

function doPost(e) {
  let res;
  try {
    if (!/^[A-Za-z0-9._-]{8,64}$/.test(CODIGO_REGISTRO) || CODIGO_REGISTRO === 'PON-AQUI-UN-CODIGO-LARGO') {
      return json_({ ok: false, error: 'config' });
    }
    const req = JSON.parse(e.postData.contents);
    if (!req || typeof req.action !== 'string') return json_({ ok: false, error: 'accion' });
    const lock = LockService.getScriptLock();
    lock.waitLock(25000);
    try {
      if (req.action === 'register') res = registrar_(req);
      else if (req.action === 'login') res = entrar_(req);
      else {
        const u = usuarioDeSesion_(req.s);
        if (!u) res = { ok: false, error: 'sesion' };
        else if (req.action === 'load') res = cargar_(u);
        else if (req.action === 'save') res = guardar_(u, req);
        else if (req.action === 'logout') { cerrarSesion_(req.s); res = { ok: true }; }
        else if (req.action === 'passwd') res = cambiarClave_(u, req);
        else res = { ok: false, error: 'accion' };
      }
    } finally {
      lock.releaseLock();
    }
  } catch (err) {
    res = { ok: false, error: 'interno' };
  }
  return json_(res);
}

/* Si editas la hoja a mano, sube la versión de ese usuario para que la app lo note. */
function onEdit(e) {
  const r = e && e.range;
  if (!r) return;
  const sh = r.getSheet(), n = sh.getName();
  if (n !== HOJA_MOV && n !== HOJA_METAS && n !== HOJA_CFG) return;
  const vistos = {};
  sh.getRange(r.getRow(), 1, r.getNumRows(), 1).getValues().forEach(function (f) {
    const u = aTexto_(f[0]).trim().toLowerCase();
    if (u) vistos[u] = true;
  });
  const lista = Object.keys(vistos);
  const filas = leerFilas_(HOJA_USR, COLS_USR.length);
  filas.forEach(function (f, i) {
    if (!lista.length || vistos[aTexto_(f[0])]) {
      hoja_(HOJA_USR, COLS_USR, FMT_USR).getRange(i + 2, 4).setValue(num_(f[3]) + 1);
    }
  });
}

/* Ejecútala una vez a mano: crea las pestañas y pide los permisos. */
function iniciar() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  hoja_(HOJA_MOV, COLS_MOV, FMT_MOV);
  hoja_(HOJA_METAS, COLS_META, FMT_META);
  hoja_(HOJA_CFG, COLS_CFG, FMT_CFG);
  hoja_(HOJA_USR, COLS_USR, FMT_USR).hideSheet();
  hoja_(HOJA_SES, COLS_SES, FMT_SES).hideSheet();
  ['Hoja 1', 'Sheet1', 'Hoja1'].forEach(function (n) {
    const sh = ss.getSheetByName(n);
    if (sh && sh.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });
  ss.toast('Listo. Ahora: Implementar > Nueva implementación > Aplicación web.');
}

/* ---------- Cuentas y sesiones ---------- */

function registrar_(req) {
  if (bloqueado_('codigo')) return { ok: false, error: 'bloqueado' };
  if (req.code !== CODIGO_REGISTRO) { fallo_('codigo'); return { ok: false, error: 'codigo' }; }
  const u = String(req.u || '').trim().toLowerCase();
  if (!RE_USUARIO.test(u) || !RE_CLAVE.test(String(req.k || ''))) return { ok: false, error: 'usuario' };
  if (buscarUsuario_(u) >= 0) return { ok: false, error: 'existe' };
  const sal = Utilities.getUuid() + Utilities.getUuid();
  agregarFila_(HOJA_USR, COLS_USR, FMT_USR, [u, sal, huella_(sal + req.k), 0, new Date().toISOString()]);
  return { ok: true, s: nuevaSesion_(u) };
}

function entrar_(req) {
  const u = String(req.u || '').trim().toLowerCase();
  if (!RE_USUARIO.test(u) || !RE_CLAVE.test(String(req.k || ''))) return { ok: false, error: 'credenciales' };
  if (bloqueado_('u:' + u)) return { ok: false, error: 'bloqueado' };
  const i = buscarUsuario_(u);
  const f = i >= 0 ? leerFilas_(HOJA_USR, COLS_USR.length)[i] : null;
  if (!f || huella_(aTexto_(f[1]) + req.k) !== aTexto_(f[2])) {
    fallo_('u:' + u);
    return { ok: false, error: 'credenciales' };
  }
  CacheService.getScriptCache().remove('f:u:' + u);
  return { ok: true, s: nuevaSesion_(u) };
}

function cambiarClave_(u, req) {
  if (!RE_CLAVE.test(String(req.o || '')) || !RE_CLAVE.test(String(req.n || ''))) return { ok: false, error: 'credenciales' };
  if (bloqueado_('u:' + u)) return { ok: false, error: 'bloqueado' };
  const i = buscarUsuario_(u), f = leerFilas_(HOJA_USR, COLS_USR.length)[i];
  if (huella_(aTexto_(f[1]) + req.o) !== aTexto_(f[2])) { fallo_('u:' + u); return { ok: false, error: 'credenciales' }; }
  const sal = Utilities.getUuid() + Utilities.getUuid(), sh = hoja_(HOJA_USR, COLS_USR, FMT_USR);
  sh.getRange(i + 2, 2, 1, 2).setValues([[sal, huella_(sal + req.n)]]);
  quitarSesiones_(function (s) { return aTexto_(s[1]) === u; });
  return { ok: true, s: nuevaSesion_(u) };
}

function nuevaSesion_(u) {
  quitarSesiones_(function (s) { return num_(s[2]) < Date.now(); });
  const tok = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
  agregarFila_(HOJA_SES, COLS_SES, FMT_SES, [huella_(tok), u, Date.now() + SESION_MS]);
  return tok;
}
function usuarioDeSesion_(tok) {
  if (typeof tok !== 'string' || tok.length < 32 || tok.length > 128) return '';
  const h = huella_(tok);
  const filas = leerFilas_(HOJA_SES, COLS_SES.length);
  for (let i = 0; i < filas.length; i++) {
    if (aTexto_(filas[i][0]) === h) return num_(filas[i][2]) > Date.now() ? aTexto_(filas[i][1]) : '';
  }
  return '';
}
function cerrarSesion_(tok) {
  const h = huella_(String(tok || ''));
  quitarSesiones_(function (s) { return aTexto_(s[0]) === h; });
}
function quitarSesiones_(cond) {
  const sh = hoja_(HOJA_SES, COLS_SES, FMT_SES);
  const filas = leerFilas_(HOJA_SES, COLS_SES.length);
  for (let i = filas.length - 1; i >= 0; i--) if (cond(filas[i])) sh.deleteRow(i + 2);
}
function buscarUsuario_(u) {
  const filas = leerFilas_(HOJA_USR, COLS_USR.length);
  for (let i = 0; i < filas.length; i++) if (aTexto_(filas[i][0]) === u) return i;
  return -1;
}
function revUsuario_(u) {
  const i = buscarUsuario_(u);
  return i < 0 ? 0 : num_(leerFilas_(HOJA_USR, COLS_USR.length)[i][3]);
}
function subirRev_(u) {
  const i = buscarUsuario_(u), v = revUsuario_(u) + 1;
  hoja_(HOJA_USR, COLS_USR, FMT_USR).getRange(i + 2, 4).setValue(v);
  return v;
}
function bloqueado_(clave) {
  return num_(CacheService.getScriptCache().get('f:' + clave)) >= (clave === 'codigo' ? 10 : MAX_FALLOS);
}
function fallo_(clave) {
  const c = CacheService.getScriptCache(), k = 'f:' + clave;
  c.put(k, String(num_(c.get(k)) + 1), BLOQUEO_S);
}
function huella_(s) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, s, Utilities.Charset.UTF_8)
    .map(function (b) { return ('0' + (b & 255).toString(16)).slice(-2); }).join('');
}

/* ---------- Cargar y guardar (solo los datos de ese usuario) ---------- */

function cargar_(u) {
  const rev = revUsuario_(u);
  if (rev === 0) return { ok: true, rev: 0, data: null };
  const c = cfgDesdeFilas_(suyas_(HOJA_CFG, COLS_CFG.length, u));
  return {
    ok: true,
    rev: rev,
    data: {
      cfg: c.cfg,
      tx: movDesdeFilas_(suyas_(HOJA_MOV, COLS_MOV.length, u)),
      goals: metasDesdeFilas_(suyas_(HOJA_METAS, COLS_META.length, u)),
      tax: c.tax
    }
  };
}

function guardar_(u, req) {
  const rev = revUsuario_(u);
  if (Number(req.baseRev) !== rev) {
    const c = cargar_(u);
    return { ok: false, conflict: true, rev: c.rev, data: c.data };
  }
  const d = req.data;
  if (!d || !Array.isArray(d.tx) || !Array.isArray(d.goals) || d.tx.length > 20000 || d.goals.length > 500) return { ok: false, error: 'datos' };
  reemplazar_(HOJA_MOV, COLS_MOV, FMT_MOV, u, filasMov_(d.tx));
  reemplazar_(HOJA_METAS, COLS_META, FMT_META, u, filasMetas_(d.goals));
  reemplazar_(HOJA_CFG, COLS_CFG, FMT_CFG, u, filasCfg_(d.cfg, d.tax));
  return { ok: true, rev: subirRev_(u) };
}

/* ---------- Conversión datos <-> filas (sin el usuario en la primera columna) ---------- */

function filasMov_(tx) {
  return (tx || []).map(function (x) {
    return [x.id, x.d, TIPO_A_TXT[x.t] || '', x.c || '', x.x || '', x.a, x.k ? 'Sí' : '', x.g || '', x.b || '', x.gb || '', x.wh || ''];
  });
}
function filasMetas_(goals) {
  return (goals || []).map(function (g) { return [g.id, g.n, g.a]; });
}
function filasCfg_(cfg, tax) {
  const f = [];
  CFG_CLAVES.forEach(function (k) { f.push(['app', k, cfg && cfg[k] != null ? cfg[k] : '']); });
  TAX_CLAVES.forEach(function (k) { f.push(['impuestos', k, tax && tax[k] != null ? tax[k] : '']); });
  return f;
}
function movDesdeFilas_(filas) {
  const tx = [];
  filas.forEach(function (r, i) {
    const fecha = aTexto_(r[1]).trim().slice(0, 10);
    const t = tipoDesdeTxt_(r[2]);
    if (!fecha || !t) return;
    const x = { id: num_(r[0]) || (Date.now() + i), d: fecha, t: t, c: aTexto_(r[3]), x: aTexto_(r[4]), a: num_(r[5]) };
    if (/^(s[ií]|1|true|x)$/i.test(aTexto_(r[6]).trim())) x.k = 1;
    if (num_(r[7])) x.g = num_(r[7]);
    if (aTexto_(r[8])) x.b = aTexto_(r[8]);
    if (num_(r[9])) x.gb = num_(r[9]);
    if (num_(r[10])) x.wh = num_(r[10]);
    tx.push(x);
  });
  return tx;
}
function metasDesdeFilas_(filas) {
  const out = [];
  filas.forEach(function (r, i) {
    const n = aTexto_(r[1]).trim();
    if (!n) return;
    out.push({ id: num_(r[0]) || (Date.now() + i), n: n, a: num_(r[2]) });
  });
  return out;
}
function cfgDesdeFilas_(filas) {
  const cfg = {}, tax = {};
  filas.forEach(function (r) {
    const s = aTexto_(r[0]).trim().toLowerCase(), k = aTexto_(r[1]).trim();
    if (!k) return;
    if (s === 'app' && CFG_CLAVES.indexOf(k) >= 0) cfg[k] = aTexto_(r[2]);
    else if (s === 'impuestos' && TAX_CLAVES.indexOf(k) >= 0) tax[k] = typeof r[2] === 'number' ? r[2] : aTexto_(r[2]);
  });
  return { cfg: cfg, tax: tax };
}
function tipoDesdeTxt_(v) {
  const s = aTexto_(v).trim().toLowerCase();
  if (s === 'ingreso' || s === 'i') return 'i';
  if (s === 'gasto' || s === 'g') return 'g';
  if (s === 'ahorro' || s === 's') return 's';
  return '';
}
function aTexto_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return v == null ? '' : String(v);
}
function num_(v) {
  if (typeof v === 'number') return isFinite(v) ? v : 0;
  const n = parseFloat(String(v == null ? '' : v).replace(/,/g, ''));
  return isFinite(n) ? n : 0;
}

/* ---------- Acceso a la hoja ---------- */

function hoja_(nombre, cols, fmt) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(nombre);
  if (!sh) {
    sh = ss.insertSheet(nombre);
    sh.getRange(1, 1, 1, cols.length).setNumberFormat('@').setValues([cols]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function leerFilas_(nombre, n) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(nombre);
  if (!sh || sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, n).getValues();
}

/* Filas de un usuario, sin la columna "usuario". */
function suyas_(nombre, n, u) {
  return leerFilas_(nombre, n)
    .filter(function (r) { return aTexto_(r[0]).trim().toLowerCase() === u; })
    .map(function (r) { return r.slice(1); });
}

function agregarFila_(nombre, cols, fmt, fila) {
  const sh = hoja_(nombre, cols, fmt);
  const rg = sh.getRange(sh.getLastRow() + 1, 1, 1, cols.length);
  rg.setNumberFormats([fmt]);
  rg.setValues([fila]);
}

/* Quita las filas de ese usuario y escribe las nuevas; las de los demás quedan igual. */
function reemplazar_(nombre, cols, fmt, u, filas) {
  const sh = hoja_(nombre, cols, fmt);
  const otras = leerFilas_(nombre, cols.length).filter(function (r) { return aTexto_(r[0]).trim().toLowerCase() !== u; });
  const todas = otras.concat(filas.map(function (f) { return [u].concat(f); }));
  sh.clearContents();
  sh.getRange(1, 1, 1, cols.length).setNumberFormat('@').setValues([cols]).setFontWeight('bold');
  sh.setFrozenRows(1);
  if (todas.length) {
    const rg = sh.getRange(2, 1, todas.length, cols.length);
    // Formato ANTES de escribir: evita que Sheets convierta fechas o ejecute fórmulas ("=...").
    rg.setNumberFormats(todas.map(function () { return fmt; }));
    rg.setValues(todas);
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
