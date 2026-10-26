/**
 * DATOS FIJOS — SERVIDOR (solicitudes de pago + PIN de 4 dígitos de un solo uso)
 *
 * ACTUALIZAR (ya tienes el proyecto de Apps Script):
 *  1) Borra todo en Código.gs y pega este archivo completo. Guarda (Ctrl+S).
 *  2) Ejecuta la función "configurar" una vez (agrega la columna Teléfono a tu hoja).
 *  3) Implementar > Administrar implementaciones > editar (lápiz) > Versión: Nueva versión > Implementar.
 *     La URL /exec NO cambia.
 *
 * AVISOS (opcionales): Configuración del proyecto > Propiedades del script:
 *   NOTIFY_EMAIL    = tucorreo@gmail.com
 *   CALLMEBOT_PHONE = 584123334053
 *   CALLMEBOT_KEY   = la apikey que te dé CallMeBot
 * Si cambias la clave de administrador en la página, cambia ADMIN_HASH en esas mismas propiedades.
 */

var ADMIN_HASH_DEFECTO = '2a969ff6de4bc4bbd7de67f841000663c4cc716662e57ae798d4d56afbdb3c06';
var SOL_HOJA = 'Solicitudes';
var NCOL = 11;                      // ID, Token, Fecha, Nombre, Referencia, Estado, PIN, Intentos, Aprobado, Usado, Teléfono
var SOL_MAX_INTENTOS = 5;           // intentos fallidos antes de bloquear
var SOL_VIGENCIA_HORAS = 48;        // el PIN vence a las 48 h de aprobado
var SOL_MAX_PENDIENTES = 100;       // freno contra spam de solicitudes

// ---------- entrada de la página ----------
function doGet() { return ContentService.createTextOutput('Datos Fijos: servidor activo'); }
function doPost(e) {
  var salida;
  try {
    var d = JSON.parse(e.postData.contents);
    salida = solicitudesApi_(d) || { ok: false, error: 'accion' };
  } catch (err) { salida = { ok: false, error: 'servidor' }; }
  return ContentService.createTextOutput(JSON.stringify(salida)).setMimeType(ContentService.MimeType.JSON);
}
// Ejecuta ESTA función una vez para dar permisos y preparar la hoja
function configurar() {
  var ss = getSS_(); solHoja_();
  Logger.log('Listo. Tu hoja de Google: ' + ss.getUrl());
}
function getSS_() {
  var pr = PropertiesService.getScriptProperties();
  var id = pr.getProperty('SHEET_ID');
  if (id) return SpreadsheetApp.openById(id);
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) ss = SpreadsheetApp.create('Datos Fijos - Solicitudes');
  pr.setProperty('SHEET_ID', ss.getId());
  return ss;
}

function solicitudesApi_(d) {
  if (!d || typeof d.accion !== 'string' || d.accion.indexOf('sol_') !== 0) return null;
  var lock = LockService.getScriptLock();
  try { lock.waitLock(10000); } catch (e) { return { ok: false, error: 'ocupado' }; }
  try {
    switch (d.accion) {
      case 'sol_crear':        return solCrear_(d);
      case 'sol_estado':       return solEstado_(d);
      case 'sol_canjear':      return solCanjear_(d);
      case 'sol_canjear_tel':  return solCanjearTel_(d);
      case 'sol_liberar':      return solLiberar_(d);
      case 'sol_listar':       return solListar_(d);
      case 'sol_aprobar':      return solAprobar_(d);
      case 'sol_rechazar':     return solRechazar_(d);
      case 'sol_nuevopin':     return solNuevoPin_(d);
      case 'sol_restablecer':  return solRestablecer_(d);
      case 'sol_anular':       return solAnular_(d);
      case 'sol_reasignar':    return solReasignar_(d);
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
  var ss = getSS_(), sh = ss.getSheetByName(SOL_HOJA);
  if (!sh) {
    sh = ss.insertSheet(SOL_HOJA);
    sh.appendRow(['ID', 'Token', 'Fecha', 'Nombre', 'Referencia', 'Estado', 'PIN', 'Intentos', 'Aprobado', 'Usado', 'Teléfono']);
    sh.setFrozenRows(1);
    sh.getRange('A:B').setNumberFormat('@');
    sh.getRange('E:E').setNumberFormat('@');
    sh.getRange('G:G').setNumberFormat('@');   // el PIN se guarda como texto (conserva ceros: 0042)
    sh.getRange('K:K').setNumberFormat('@');
  } else if (!sh.getRange(1, 11).getValue()) {
    sh.getRange('K:K').setNumberFormat('@');
    sh.getRange(1, 11).setValue('Teléfono');
  }
  return sh;
}
function esAdmin_(h) {
  var real = PropertiesService.getScriptProperties().getProperty('ADMIN_HASH') || ADMIN_HASH_DEFECTO;
  return typeof h === 'string' && h.length === real.length && h === real;
}
function limpiar_(v, max) {
  return String(v == null ? '' : v).replace(/[\r\n\t]/g, ' ').replace(/^[=+\-@]+/, '').trim().slice(0, max);
}
// Teléfono a formato internacional sin "+" (0412-3334053 -> 584123334053). '' si no es válido
function tel_(v) {
  var d = String(v || '').replace(/\D/g, '');
  if (d.indexOf('00') === 0) d = d.slice(2);
  if (d.charAt(0) === '0') d = '58' + d.slice(1);
  else if (d.length === 10 && d.charAt(0) === '4') d = '58' + d;
  return (d.length >= 11 && d.length <= 15) ? d : '';
}
function idCorto_() {
  var a = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789', r = '';
  for (var i = 0; i < 6; i++) r += a.charAt(Math.floor(Math.random() * a.length));
  return r;
}
function token_() { return Utilities.getUuid().replace(/-/g, ''); }
function pin4_() {
  var n = parseInt(Utilities.getUuid().replace(/-/g, '').slice(0, 8), 16) % 10000;
  return ('0000' + n).slice(-4);
}
function todas_(sh) {
  var last = sh.getLastRow();
  return last > 1 ? sh.getRange(2, 1, last - 1, NCOL).getValues() : [];
}
function buscar_(sh, id) {
  var vals = todas_(sh);
  for (var i = 0; i < vals.length; i++) if (String(vals[i][0]) === String(id)) return { fila: i + 2, v: vals[i] };
  return null;
}
function tokenOk_(reg, token) { return reg && typeof token === 'string' && token.length > 10 && String(reg.v[1]) === token; }
function solVencido_(reg) {
  var ap = new Date(reg.v[8]).getTime();
  return !ap || (Date.now() - ap) > SOL_VIGENCIA_HORAS * 3600 * 1000;
}
function ponerPin_(sh, reg, pin) {
  sh.getRange(reg.fila, 7).setNumberFormat('@').setValue(pin);
  sh.getRange(reg.fila, 6).setValue('aprobado');
  sh.getRange(reg.fila, 8).setValue(0);
  sh.getRange(reg.fila, 9).setValue(new Date().toISOString());
  sh.getRange(reg.fila, 10).setValue('');
  var tel = String(reg.v[10] || '');
  if (tel) CacheService.getScriptCache().remove('ft_' + tel);   // limpia bloqueos por intentos
}
function marcarUsado_(sh, reg) {
  sh.getRange(reg.fila, 6).setValue('usado');
  sh.getRange(reg.fila, 10).setValue(new Date().toISOString());
  notificar_(String(reg.v[3]) + ' usó su PIN y está descargando las 3 hojas.');
}

// ---------- cliente ----------
function solCrear_(d) {
  var nombre = limpiar_(d.nombre, 40), ref = String(d.ref || '').replace(/\D/g, '').slice(0, 12), tel = tel_(d.tel);
  if (!nombre || ref.length < 4 || !tel) return { ok: false, error: 'datos' };
  var sh = solHoja_(), vals = todas_(sh), pend = 0;
  for (var i = 0; i < vals.length; i++) if (vals[i][5] === 'pendiente') pend++;
  if (pend >= SOL_MAX_PENDIENTES) return { ok: false, error: 'saturado' };
  var id = idCorto_(), token = token_();
  sh.appendRow([id, token, new Date().toISOString(), nombre, '', 'pendiente', '', 0, '', '', '']);
  var f = sh.getLastRow();
  sh.getRange(f, 5).setNumberFormat('@').setValue(ref);
  sh.getRange(f, 11).setNumberFormat('@').setValue(tel);
  notificar_('Nueva solicitud de pago de ' + nombre + ' (ref ' + ref + ', tel +' + tel + '). Entra a tu panel (#admin) para aceptarla.');
  return { ok: true, id: id, token: token };
}
function solEstado_(d) {
  var reg = buscar_(solHoja_(), d.id);
  if (!reg || !tokenOk_(reg, d.token)) return { ok: false, error: 'no_existe' };
  var est = reg.v[5];
  if (est === 'aprobado' && solVencido_(reg)) est = 'rechazado';
  return { ok: true, estado: est, pin: est === 'aprobado' ? String(reg.v[6]) : '' };
}
function solCanjear_(d) {
  var sh = solHoja_(), reg = buscar_(sh, d.id);
  if (!reg || !tokenOk_(reg, d.token)) return { ok: false, error: 'no_existe' };
  var est = reg.v[5], intentos = Number(reg.v[7]) || 0;
  if (est === 'usado') return { ok: false, error: 'usado' };
  if (est === 'bloqueado' || intentos >= SOL_MAX_INTENTOS) return { ok: false, error: 'bloqueado' };
  if (est !== 'aprobado') return { ok: false, error: est === 'pendiente' ? 'pendiente' : est };   // rechazado / anulado
  if (solVencido_(reg)) return { ok: false, error: 'expirado' };
  var pin = String(d.pin || '').replace(/\D/g, '');
  if (pin !== String(reg.v[6])) {
    intentos++;
    sh.getRange(reg.fila, 8).setValue(intentos);
    if (intentos >= SOL_MAX_INTENTOS) { sh.getRange(reg.fila, 6).setValue('bloqueado'); return { ok: false, error: 'bloqueado' }; }
    return { ok: false, error: 'pin', restantes: SOL_MAX_INTENTOS - intentos };
  }
  marcarUsado_(sh, reg);   // PIN correcto: se quema de inmediato
  return { ok: true };
}
// Para quien cerró la página o cambió de teléfono: PIN + número de teléfono de la solicitud
function solCanjearTel_(d) {
  var tel = tel_(d.tel), pin = String(d.pin || '').replace(/\D/g, '');
  if (!tel || pin.length !== 4) return { ok: false, error: 'datos' };
  var cache = CacheService.getScriptCache(), key = 'ft_' + tel, fallos = Number(cache.get(key)) || 0;
  if (fallos >= SOL_MAX_INTENTOS) return { ok: false, error: 'bloqueado' };
  var sh = solHoja_(), vals = todas_(sh), usadoCoincide = false, vencido = false;
  for (var i = vals.length - 1; i >= 0; i--) {
    var v = vals[i];
    if (String(v[10]) !== tel || String(v[6]) !== pin) continue;
    var reg = { fila: i + 2, v: v };
    if (v[5] === 'aprobado') {
      if (solVencido_(reg)) { vencido = true; continue; }
      marcarUsado_(sh, reg); cache.remove(key);
      return { ok: true, id: String(v[0]), token: String(v[1]) };
    }
    if (v[5] === 'usado') usadoCoincide = true;
  }
  if (usadoCoincide) return { ok: false, error: 'usado' };
  if (vencido) return { ok: false, error: 'expirado' };
  fallos++; cache.put(key, String(fallos), 21600);
  if (fallos >= SOL_MAX_INTENTOS) return { ok: false, error: 'bloqueado' };
  return { ok: false, error: 'pin', restantes: SOL_MAX_INTENTOS - fallos };
}
// Si la descarga falló en el teléfono del cliente, se le devuelve el PIN (solo dentro de 10 min)
function solLiberar_(d) {
  var sh = solHoja_(), reg = buscar_(sh, d.id);
  if (!reg || !tokenOk_(reg, d.token) || reg.v[5] !== 'usado') return { ok: false, error: 'no_existe' };
  var usado = new Date(reg.v[9]).getTime();
  if (!usado || Date.now() - usado > 10 * 60 * 1000) return { ok: false, error: 'tarde' };
  sh.getRange(reg.fila, 6).setValue('aprobado');
  sh.getRange(reg.fila, 10).setValue('');
  return { ok: true };
}

// ---------- administrador ----------
function solListar_(d) {
  if (!esAdmin_(d.h)) return { ok: false, error: 'auth' };
  var vals = todas_(solHoja_()), filas = [];
  for (var i = vals.length - 1; i >= 0 && filas.length < 40; i--) {
    var v = vals[i], conPin = (v[5] === 'aprobado' || v[5] === 'usado');
    filas.push({ id: String(v[0]), fecha: String(v[2]), nombre: String(v[3]), ref: String(v[4]), estado: String(v[5]),
                 pin: conPin ? String(v[6]) : '', usado: String(v[9] || ''), tel: String(v[10] || '') });
  }
  filas.sort(function (a, b) { return (a.estado === 'pendiente' ? 0 : 1) - (b.estado === 'pendiente' ? 0 : 1); });
  return { ok: true, filas: filas };
}
function adminReg_(d) {
  if (!esAdmin_(d.h)) return { error: 'auth' };
  var sh = solHoja_(), reg = buscar_(sh, d.id);
  if (!reg) return { error: 'no_existe' };
  return { sh: sh, reg: reg };
}
function solAprobar_(d) {
  var a = adminReg_(d); if (a.error) return { ok: false, error: a.error };
  if (a.reg.v[5] !== 'pendiente') return { ok: false, error: 'estado' };
  var pin = pin4_(); ponerPin_(a.sh, a.reg, pin);
  return { ok: true, pin: pin };
}
function solRechazar_(d) {
  var a = adminReg_(d); if (a.error) return { ok: false, error: a.error };
  if (a.reg.v[5] !== 'pendiente') return { ok: false, error: 'estado' };
  a.sh.getRange(a.reg.fila, 6).setValue('rechazado');
  return { ok: true };
}
// Cambia el PIN por uno nuevo (queda disponible, intentos en 0, vigencia reiniciada)
function solNuevoPin_(d) {
  var a = adminReg_(d); if (a.error) return { ok: false, error: a.error };
  if (a.reg.v[5] === 'pendiente') return { ok: false, error: 'estado' };
  var pin = pin4_(); ponerPin_(a.sh, a.reg, pin);
  return { ok: true, pin: pin };
}
// Deja disponible otra vez el MISMO PIN (si estaba usado, bloqueado o anulado)
function solRestablecer_(d) {
  var a = adminReg_(d); if (a.error) return { ok: false, error: a.error };
  var est = a.reg.v[5];
  if (est === 'pendiente' || est === 'aprobado') return { ok: false, error: 'estado' };
  var pin = String(a.reg.v[6] || '');
  if (!/^\d{4}$/.test(pin)) pin = pin4_();
  ponerPin_(a.sh, a.reg, pin);
  return { ok: true, pin: pin };
}
// Borra el PIN (ya no sirve)
function solAnular_(d) {
  var a = adminReg_(d); if (a.error) return { ok: false, error: a.error };
  if (a.reg.v[5] !== 'aprobado') return { ok: false, error: 'estado' };
  a.sh.getRange(a.reg.fila, 6).setValue('anulado');
  a.sh.getRange(a.reg.fila, 7).setValue('');
  return { ok: true };
}
// Pasa el mismo PIN a otro cliente: cambia teléfono (y nombre), invalida la sesión del anterior y lo deja disponible
function solReasignar_(d) {
  var a = adminReg_(d); if (a.error) return { ok: false, error: a.error };
  var est = a.reg.v[5];
  if (est !== 'aprobado' && est !== 'usado' && est !== 'bloqueado') return { ok: false, error: 'estado' };
  var tel = tel_(d.tel); if (!tel) return { ok: false, error: 'datos' };
  var pin = String(a.reg.v[6] || ''); if (!/^\d{4}$/.test(pin)) pin = pin4_();
  var nombre = limpiar_(d.nombre, 40);
  if (nombre) a.sh.getRange(a.reg.fila, 4).setValue(nombre);
  a.sh.getRange(a.reg.fila, 11).setNumberFormat('@').setValue(tel);
  a.sh.getRange(a.reg.fila, 2).setValue(token_());
  a.reg.v[10] = tel;
  ponerPin_(a.sh, a.reg, pin);
  return { ok: true, pin: pin };
}
