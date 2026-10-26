/**
 * DATOS FIJOS — SERVIDOR COMPLETO (todo en un solo archivo)
 * Incluye: códigos RDN de un solo uso, solicitudes de pago, PIN de 4 dígitos y avisos.
 *
 * CÓMO USARLO (proyecto nuevo de Apps Script, sin nada más):
 *  1) Borra todo lo que haya en Código.gs y pega este archivo completo. Guarda (Ctrl+S).
 *  2) Arriba, en el selector de funciones, elige "configurar" y pulsa Ejecutar.
 *     Google pedirá permisos: Revisar permisos > tu cuenta > Avanzado > Ir a (proyecto) > Permitir.
 *     Al terminar, en Ver > Registros verás el enlace de tu hoja de Google creada sola.
 *  3) Implementar > Nueva implementación > engranaje > Aplicación web.
 *       Ejecutar como: Yo        Quién tiene acceso: Cualquier persona
 *     Pulsa Implementar y copia la URL que termina en /exec.
 *  4) En tu index.html busca  const API_URL = '...'  y pon ahí tu URL nueva.
 *
 * AVISOS (opcionales): Configuración del proyecto > Propiedades del script:
 *   NOTIFY_EMAIL    = tucorreo@gmail.com        (aviso por correo)
 *   CALLMEBOT_PHONE = 584123334053              (aviso por WhatsApp con CallMeBot)
 *   CALLMEBOT_KEY   = la apikey que te dé CallMeBot
 *     (guarda +34 644 51 95 23 en tus contactos y envíale: I allow callmebot to send me messages)
 * Si cambias la clave de administrador en la página, cambia ADMIN_HASH en esas mismas propiedades.
 */

var ADMIN_HASH_DEFECTO = '2a969ff6de4bc4bbd7de67f841000663c4cc716662e57ae798d4d56afbdb3c06';
var HOJA_CODIGOS = 'Codigos';
var ALFABETO = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';

// ---------- punto de entrada de la página ----------
function doGet() {
  return ContentService.createTextOutput('Datos Fijos: servidor activo');
}
function doPost(e) {
  var salida;
  try {
    var d = JSON.parse(e.postData.contents);
    salida = solicitudesApi_(d) || codigosApi_(d);
  } catch (err) { salida = { ok: false, error: 'servidor' }; }
  return ContentService.createTextOutput(JSON.stringify(salida)).setMimeType(ContentService.MimeType.JSON);
}

// Ejecuta ESTA función una vez para dar permisos y crear la hoja
function configurar() {
  var ss = getSS_();
  codigosHoja_(); solHoja_();
  Logger.log('Listo. Tu hoja de Google: ' + ss.getUrl());
}

function getSS_() {
  var pr = PropertiesService.getScriptProperties();
  var id = pr.getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) { ss = SpreadsheetApp.create('Datos Fijos - Códigos y solicitudes'); }
  pr.setProperty('SHEET_ID', ss.getId());
  return ss;
}

// ---------- códigos RDN de un solo uso (los que ya usaba tu página) ----------
function codigosHoja_() {
  var ss = getSS_(), sh = ss.getSheetByName(HOJA_CODIGOS);
  if (!sh) {
    sh = ss.insertSheet(HOJA_CODIGOS);
    sh.appendRow(['Codigo', 'Estado', 'Generado', 'Usado', 'Nota']);
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('@');
  }
  return sh;
}
function normCodigo_(c) { return String(c || '').toUpperCase().replace(/[^A-Z0-9]/g, ''); }
function fechaVe_(dt) { return Utilities.formatDate(dt, 'America/Caracas', 'dd/MM/yy HH:mm'); }

function codigosApi_(d) {
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return { ok: false, error: 'ocupado' }; }
  try {
    var sh = codigosHoja_();
    var last = sh.getLastRow();
    var vals = last > 1 ? sh.getRange(2, 1, last - 1, 5).getValues() : [];
    function fila(codigo) {
      var n = normCodigo_(codigo);
      for (var i = 0; i < vals.length; i++) if (normCodigo_(vals[i][0]) === n) return i + 2;
      return 0;
    }
    if (d.accion === 'canjear') {
      var f = fila(d.codigo);
      if (!f) return { ok: false, error: 'invalido' };
      if (vals[f - 2][1] === 'Usado') return { ok: false, error: 'usado' };
      sh.getRange(f, 2).setValue('Usado');
      sh.getRange(f, 4).setValue(fechaVe_(new Date()));
      return { ok: true };
    }
    if (!esAdmin_(d.h)) return { ok: false, error: 'auth' };
    if (d.accion === 'listar') {
      var filas = [];
      for (var i = vals.length - 1; i >= 0 && filas.length < 300; i--)
        filas.push({ c: String(vals[i][0]), e: String(vals[i][1]), g: String(vals[i][2]), u: String(vals[i][3]), n: String(vals[i][4]) });
      return { ok: true, filas: filas };
    }
    if (d.accion === 'generar') {
      var codigo;
      do {
        var raw = '';
        for (var k = 0; k < 12; k++) raw += ALFABETO.charAt(Math.floor(Math.random() * 32));
        codigo = 'RDN-' + raw.match(/.{4}/g).join('-');
      } while (fila(codigo));
      sh.appendRow([codigo, 'Disponible', fechaVe_(new Date()), '', '']);
      return { ok: true, codigo: codigo };
    }
    if (d.accion === 'nota') {
      var fn = fila(d.codigo); if (!fn) return { ok: false, error: 'invalido' };
      sh.getRange(fn, 5).setValue(limpiar_(d.nota, 60));
      return { ok: true };
    }
    if (d.accion === 'restablecer') {
      var fr = fila(d.codigo); if (!fr) return { ok: false, error: 'invalido' };
      sh.getRange(fr, 2).setValue('Disponible'); sh.getRange(fr, 4).setValue('');
      return { ok: true };
    }
    return { ok: false, error: 'accion' };
  } finally { lock.releaseLock(); }
}

// ---------- solicitudes de pago + PIN de 4 dígitos ----------
var SOL_HOJA = 'Solicitudes';
var SOL_MAX_INTENTOS = 5;          // intentos fallidos de PIN antes de bloquear la solicitud
var SOL_VIGENCIA_HORAS = 48;       // el PIN vence a las 48 h de aprobado
var SOL_MAX_PENDIENTES = 100;      // freno contra spam de solicitudes

function solicitudesApi_(d) {
  if (!d || typeof d.accion !== 'string' || d.accion.indexOf('sol_') !== 0) return null;
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return { ok: false, error: 'ocupado' }; }
  try {
    switch (d.accion) {
      case 'sol_crear':    return solCrear_(d);
      case 'sol_estado':   return solEstado_(d);
      case 'sol_canjear':  return solCanjear_(d);
      case 'sol_liberar':  return solLiberar_(d);
      case 'sol_listar':   return solListar_(d);
      case 'sol_aprobar':  return solAprobar_(d);
      case 'sol_rechazar': return solRechazar_(d);
    }
    return { ok: false, error: 'accion' };
  } finally { lock.releaseLock(); }
}

// ---------- avisos al administrador ----------
function notificar_(texto) {
  var pr = PropertiesService.getScriptProperties();
  try {
    var key = pr.getProperty('CALLMEBOT_KEY'), tel = pr.getProperty('CALLMEBOT_PHONE') || '584123334053';
    if (key) UrlFetchApp.fetch('https://api.callmebot.com/whatsapp.php?phone=' + encodeURIComponent(tel) +
      '&text=' + encodeURIComponent(texto) + '&apikey=' + encodeURIComponent(key), { muteHttpExceptions: true });
  } catch (e) {}
  try {
    var mail = pr.getProperty('NOTIFY_EMAIL');
    if (mail) MailApp.sendEmail(mail, 'Datos Fijos: aviso', texto);
  } catch (e) {}
}

// ---------- utilidades ----------
function solHoja_() {
  var ss = getSS_();
  var sh = ss.getSheetByName(SOL_HOJA);
  if (!sh) {
    sh = ss.insertSheet(SOL_HOJA);
    sh.appendRow(['ID', 'Token', 'Fecha', 'Nombre', 'Referencia', 'Estado', 'PIN', 'Intentos', 'Aprobado', 'Usado']);
    sh.setFrozenRows(1);
    sh.getRange('A:B').setNumberFormat('@');
    sh.getRange('E:E').setNumberFormat('@');
    sh.getRange('G:G').setNumberFormat('@'); // el PIN se guarda como texto (conserva ceros: 0042)
  }
  return sh;
}
function esAdmin_(h) {
  var real = PropertiesService.getScriptProperties().getProperty('ADMIN_HASH') || ADMIN_HASH_DEFECTO;
  return !!real && typeof h === 'string' && h.length === real.length && h === real;
}
function limpiar_(v, max) {
  return String(v == null ? '' : v).replace(/[\r\n\t]/g, ' ').replace(/^[=+\-@]+/, '').trim().slice(0, max);
}
function idCorto_() {
  var a = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789', r = '';
  for (var i = 0; i < 6; i++) r += a.charAt(Math.floor(Math.random() * a.length));
  return r;
}
function pin4_() {
  var n = parseInt(Utilities.getUuid().replace(/-/g, '').slice(0, 8), 16) % 10000;
  return ('0000' + n).slice(-4);
}
// Devuelve {fila: nº de fila en la hoja, v: valores} o null
function buscar_(sh, id) {
  var last = sh.getLastRow(); if (last < 2) return null;
  var vals = sh.getRange(2, 1, last - 1, 10).getValues();
  for (var i = 0; i < vals.length; i++) if (String(vals[i][0]) === String(id)) return { fila: i + 2, v: vals[i] };
  return null;
}
function tokenOk_(reg, token) {
  return reg && typeof token === 'string' && token.length > 10 && String(reg.v[1]) === token;
}

// ---------- cliente ----------
function solCrear_(d) {
  var nombre = limpiar_(d.nombre, 40), ref = String(d.ref || '').replace(/\D/g, '').slice(0, 12);
  if (!nombre || ref.length < 4) return { ok: false, error: 'datos' };
  var sh = solHoja_(), last = sh.getLastRow(), pend = 0;
  if (last > 1) {
    var est = sh.getRange(2, 6, last - 1, 1).getValues();
    for (var i = 0; i < est.length; i++) if (est[i][0] === 'pendiente') pend++;
  }
  if (pend >= SOL_MAX_PENDIENTES) return { ok: false, error: 'saturado' };
  var id = idCorto_(), token = Utilities.getUuid().replace(/-/g, '');
  sh.appendRow([id, token, new Date().toISOString(), nombre, ref, 'pendiente', '', 0, '', '']);
  sh.getRange(sh.getLastRow(), 5).setNumberFormat('@').setValue(ref);
  notificar_('Nueva solicitud de pago de ' + nombre + ' (ref ' + ref + '). Entra a tu panel (#admin) para aceptarla.');
  return { ok: true, id: id, token: token };
}
function solEstado_(d) {
  var reg = buscar_(solHoja_(), d.id);
  if (!reg || !tokenOk_(reg, d.token)) return { ok: false, error: 'no_existe' };
  var est = reg.v[5];
  if (est === 'aprobado' && solVencido_(reg)) est = 'rechazado';
  return { ok: true, estado: est, pin: est === 'aprobado' ? String(reg.v[6]) : '' };
}
function solVencido_(reg) {
  var ap = new Date(reg.v[8]).getTime();
  return !ap || (Date.now() - ap) > SOL_VIGENCIA_HORAS * 3600 * 1000;
}
function solCanjear_(d) {
  var sh = solHoja_(), reg = buscar_(sh, d.id);
  if (!reg || !tokenOk_(reg, d.token)) return { ok: false, error: 'no_existe' };
  var est = reg.v[5], intentos = Number(reg.v[7]) || 0;
  if (est === 'usado') return { ok: false, error: 'usado' };
  if (est === 'bloqueado' || intentos >= SOL_MAX_INTENTOS) return { ok: false, error: 'bloqueado' };
  if (est !== 'aprobado') return { ok: false, error: est === 'rechazado' ? 'rechazado' : 'pendiente' };
  if (solVencido_(reg)) return { ok: false, error: 'expirado' };
  var pin = String(d.pin || '').replace(/\D/g, '');
  if (pin !== String(reg.v[6])) {
    intentos++;
    sh.getRange(reg.fila, 8).setValue(intentos);
    if (intentos >= SOL_MAX_INTENTOS) { sh.getRange(reg.fila, 6).setValue('bloqueado'); return { ok: false, error: 'bloqueado' }; }
    return { ok: false, error: 'pin', restantes: SOL_MAX_INTENTOS - intentos };
  }
  // PIN correcto: se quema de inmediato (un solo uso)
  sh.getRange(reg.fila, 6).setValue('usado');
  sh.getRange(reg.fila, 10).setValue(new Date().toISOString());
  notificar_(String(reg.v[3]) + ' usó su PIN y está descargando las 3 hojas.');
  return { ok: true };
}
// Si la descarga falló en el teléfono del cliente, se le devuelve el PIN (mismo token, un solo intento extra)
function solLiberar_(d) {
  var sh = solHoja_(), reg = buscar_(sh, d.id);
  if (!reg || !tokenOk_(reg, d.token) || reg.v[5] !== 'usado') return { ok: false, error: 'no_existe' };
  var usado = new Date(reg.v[9]).getTime();
  if (!usado || Date.now() - usado > 10 * 60 * 1000) return { ok: false, error: 'tarde' }; // solo dentro de 10 min
  sh.getRange(reg.fila, 6).setValue('aprobado');
  sh.getRange(reg.fila, 10).setValue('');
  return { ok: true };
}

// ---------- administrador ----------
function solListar_(d) {
  if (!esAdmin_(d.h)) return { ok: false, error: 'auth' };
  var sh = solHoja_(), last = sh.getLastRow(), filas = [];
  if (last > 1) {
    var vals = sh.getRange(2, 1, last - 1, 10).getValues();
    for (var i = vals.length - 1; i >= 0 && filas.length < 40; i--) {
      var v = vals[i];
      filas.push({ id: String(v[0]), fecha: String(v[2]), nombre: String(v[3]), ref: String(v[4]), estado: String(v[5]),
                   pin: v[5] === 'aprobado' ? String(v[6]) : '' });
    }
  }
  // pendientes primero
  filas.sort(function (a, b) { return (a.estado === 'pendiente' ? 0 : 1) - (b.estado === 'pendiente' ? 0 : 1); });
  return { ok: true, filas: filas };
}
function solAprobar_(d) {
  if (!esAdmin_(d.h)) return { ok: false, error: 'auth' };
  var sh = solHoja_(), reg = buscar_(sh, d.id);
  if (!reg) return { ok: false, error: 'no_existe' };
  if (reg.v[5] !== 'pendiente') return { ok: false, error: 'estado' };
  var pin = pin4_();
  sh.getRange(reg.fila, 7).setNumberFormat('@').setValue(pin);
  sh.getRange(reg.fila, 6).setValue('aprobado');
  sh.getRange(reg.fila, 8).setValue(0);
  sh.getRange(reg.fila, 9).setValue(new Date().toISOString());
  return { ok: true, pin: pin };
}
function solRechazar_(d) {
  if (!esAdmin_(d.h)) return { ok: false, error: 'auth' };
  var sh = solHoja_(), reg = buscar_(sh, d.id);
  if (!reg) return { ok: false, error: 'no_existe' };
  if (reg.v[5] !== 'pendiente') return { ok: false, error: 'estado' };
  sh.getRange(reg.fila, 6).setValue('rechazado');
  return { ok: true };
}
