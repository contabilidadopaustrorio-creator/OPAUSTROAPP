// =========================================================
// OPAUSTRO – Apps Script Unificado v4.0
// Proyecto: 1Zq5VjpFE2v6MQBro27khmffsLeJSMFtrMy7x3OCFHtmWBCILv3yMxjuQ
// Sheet usuarios: 1D4DzASd5yh-tMhKqJvshFcOV4gathx9yAsZ0xHV1F3I
// =========================================================

const SHEETS = {
  EQUIPOS:      'Equipos',
  MOVIMIENTOS:  'Movimientos',
  MANTENIMIENTO:'Mantenimiento'
};

const USUARIOS_SHEET_ID = '1D4DzASd5yh-tMhKqJvshFcOV4gathx9yAsZ0xHV1F3I';
const DRIVE_FOLDER_NAME    = 'Control Cabinets Fotos';
const DRIVE_CONTRATOS_NAME = 'Captaciones Contratos';
const APP_ROOT_NAME        = 'APP OPAUSTRO';
const TZ                   = 'America/Guayaquil';
const SESSION_TTL_MS       = 8 * 60 * 60 * 1000; // 8 horas

// Mapa rol → módulos permitidos
const ROL_MODULOS = {
  admin:        ['portal','logistica','ventas','gerencia'],
  gerencial:    ['portal','logistica','ventas','gerencia'],
  costos:       ['portal','gerencia'],
  ventas:       ['portal','ventas'],
  mapa:         ['portal','ventas'],
  logistica:    ['portal','logistica'],
  tecnico:      ['portal','logistica'],
  coordinacion: ['portal','logistica','ventas']
};

// Solo admin puede escribir datos
const ROL_PUEDE_ESCRIBIR = { admin: true };

// URLs de módulos en GitHub Pages
const MODULE_URLS = {
  logistica: 'https://contabilidadopaustrorio-creator.github.io/OPAUSTRO-LOGISTICA/',
  ventas:    'https://contabilidadopaustrorio-creator.github.io/OPAUSTRO-VENTAS/',
  gerencia:  'https://contabilidadopaustrorio-creator.github.io/OPAUSTRO-GERENCIA/',
  portal:    'https://contabilidadopaustrorio-creator.github.io/OPAUSTROAPP/'
};

function jsonResponse(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ============================================================
// doGet
// ============================================================
function doGet(e) {
  try {
    var p      = (e && e.parameter) || {};
    var action = p.action || '';

    if (action === 'login')           return jsonResponse(handleLogin(p.usuario, p.clave));
    if (action === 'validateSession') return jsonResponse(handleValidateSession(p.token));
    if (action === 'logout')          { handleLogout(p.token); return jsonResponse({ ok: true }); }

    var sess = requireSession(p.token);
    if (sess.error) return jsonResponse(sess);

    if (action === 'getData' || action === '') {
      ensureSheets();
      return jsonResponse({
        ok: true,
        db: { equipos: getEquipos(), movimientos: getMovimientos(), mantenimientos: getMantenimientos() },
        session: buildSessPublic(sess)
      });
    }
    if (action === 'getUsuarios')         return jsonResponse({ usuarios: getUsuarios() });
    if (action === 'getUsuariosGerencial')return jsonResponse({ usuarios: getUsuariosGerencial() });
    if (action === 'getUsuariosMapa')     return jsonResponse({ usuarios: getUsuariosMapa() });
    if (action === 'getUsuariosCostos')   return jsonResponse({ usuarios: getUsuariosCostos() });
    if (action === 'getCfg') {
      return jsonResponse({ ok: true, valor: PropertiesService.getScriptProperties().getProperty('cfg_' + p.clave) || '' });
    }

    return jsonResponse({ error: 'Acción no reconocida: ' + action });
  } catch (err) {
    return jsonResponse({ error: err.message });
  }
}

// ============================================================
// doPost – solo admin puede modificar datos
// ============================================================
function doPost(e) {
  try {
    var payload = JSON.parse(e.postData.contents);
    var action  = payload.action;

    if (action === 'login') return jsonResponse(handleLogin(payload.usuario, payload.clave));

    var sess = requireSession(payload.token || payload.sessionToken);
    if (sess.error) return jsonResponse({ ok: false, error: sess.error });

    if (!sess.puedeEscribir) {
      return jsonResponse({ ok: false, error: 'Acceso de solo lectura. Solo el administrador puede modificar datos.' });
    }

    ensureSheets();
    var data = payload.data || {};

    if      (action === 'entrada')          { var fe = payload.fotosEntrada || data.fotosEntrada || {}; data.fotosEntrada = subirFotosADrive(data.placa,'entrada',fe); guardarMovimiento(data); upsertEquipo(buildEquipoFromMov(data)); }
    else if (action === 'planificarSalida') { data.tipo='SALIDA'; data.estado='PLANIFICADO'; guardarMovimiento(data); }
    else if (action === 'cancelarPlan')     { cancelarPlan(data.placa, data.ts); }
    else if (action === 'actualizarFotos') {
      var inF = payload.fotosSalida||{}; var nF = subirFotosADrive(payload.placa,'salida',inF); var mF = {};
      ['f','t','i'].forEach(function(k){ if(nF[k]) mF[k]=nF[k]; else if(inF[k]&&String(inF[k]).indexOf('base64,')===-1) mF[k]=inF[k]; });
      guardarFotosSalidaEquipo(payload.placa, mF);
      upsertEquipoCampo(payload.placa, { tsUltMov: payload.ts||new Date().toISOString() });
    }
    else if (action === 'despacharSalida') { despacharSalida(data, payload.planTs||data.tsPlanificacion||'', payload.fotosContrato||{}); }
    else if (action === 'salida')          { data.tipo='SALIDA'; data.estado='DESPACHADO'; guardarMovimiento(data); marcarEquipoMercado(data,{}); }
    else if (action === 'finalizarMant')   { finalizarMantenimiento(payload); }
    else if (action === 'bajaEquipo')      { guardarMovimientoSiNoExiste(data); eliminarEquipo(data.placa); }
    else if (action === 'corregirEtapa')   { corregirEtapa(payload); }
    else if (action === 'actualizarContrato') { actualizarContrato(payload); }
    else if (action === 'enviarQuito') {
      guardarMovimientoSiNoExiste(data);
      upsertEquipoCampo(data.placa, { estado:'MANTENIMIENTO', etapa:data.etapa||'C', cliente:'', tsEtapa:data.ts, tsUltMov:data.ts, trabajos:data.trabajos||['ENVIO QUITO'] });
    }
    else if (action === 'cargarBase')  { (payload.equipos||[]).forEach(function(eq){ upsertEquipo(eq); }); }
    else if (action === 'saveCfg')     { PropertiesService.getScriptProperties().setProperty('cfg_'+payload.clave, String(payload.valor)); }
    else { throw new Error('Acción no reconocida: ' + action); }

    return jsonResponse({ ok: true });
  } catch (err) {
    return jsonResponse({ ok: false, error: err.message });
  }
}

// ============================================================
// AUTH / SESSION
// ============================================================

function handleLogin(rawUsuario, rawClave) {
  if (!rawUsuario || !rawClave) return { ok: false, error: 'Usuario y clave requeridos.' };
  var u    = normUser(rawUsuario);
  var pass = String(rawClave).trim();
  purgeExpiredSessions();
  var found = buscarEnSheet(u, pass);
  if (!found) return { ok: false, error: 'Usuario o clave incorrectos.' };
  var rol           = (found.rol || 'logistica').toLowerCase();
  var modulos       = ROL_MODULOS[rol] || ['portal'];
  var puedeEscribir = !!ROL_PUEDE_ESCRIBIR[rol];
  var token   = generateToken();
  var expires = Date.now() + SESSION_TTL_MS;
  PropertiesService.getScriptProperties().setProperty('sess_' + token, JSON.stringify({
    nombre: found.usuario, rol: rol, modulos: modulos,
    puedeEscribir: puedeEscribir, ts: Date.now(), exp: expires
  }));
  return {
    ok: true, token: token, nombre: found.usuario, rol: rol,
    modulos: modulos, puedeEscribir: puedeEscribir,
    moduleUrls: buildModuleUrls(modulos, token),
    expiresIn: SESSION_TTL_MS
  };
}

function buildModuleUrls(modulos, token) {
  var urls = {};
  modulos.forEach(function(m) { if (MODULE_URLS[m]) urls[m] = MODULE_URLS[m] + '#t=' + token; });
  return urls;
}

function handleValidateSession(token) {
  if (!token) return { valid: false, error: 'Token requerido.' };
  var sess = getSession(token);
  if (!sess) return { valid: false, error: 'Sesión inválida o expirada.' };
  return { valid: true, nombre: sess.nombre, rol: sess.rol, modulos: sess.modulos, puedeEscribir: sess.puedeEscribir };
}

function handleLogout(token) {
  if (token) PropertiesService.getScriptProperties().deleteProperty('sess_' + token);
}

function requireSession(token) {
  if (!token) return { error: 'Token de sesión requerido.' };
  var sess = getSession(token);
  if (!sess) return { error: 'Sesión inválida o expirada. Inicia sesión nuevamente.' };
  return sess;
}

function getSession(token) {
  if (!token) return null;
  var raw = PropertiesService.getScriptProperties().getProperty('sess_' + token);
  if (!raw) return null;
  try {
    var sess = JSON.parse(raw);
    if (Date.now() > sess.exp) { PropertiesService.getScriptProperties().deleteProperty('sess_' + token); return null; }
    return sess;
  } catch(e) { return null; }
}

function buildSessPublic(sess) {
  return { nombre: sess.nombre, rol: sess.rol, modulos: sess.modulos, puedeEscribir: sess.puedeEscribir };
}

function generateToken() {
  var chars = 'ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789';
  var t = '';
  for (var i = 0; i < 56; i++) t += chars.charAt(Math.floor(Math.random() * chars.length));
  return t;
}

function purgeExpiredSessions() {
  try {
    var props = PropertiesService.getScriptProperties().getProperties();
    var now = Date.now(); var del = [];
    Object.keys(props).forEach(function(k) {
      if (k.indexOf('sess_') !== 0) return;
      try { var s = JSON.parse(props[k]); if (now > s.exp) del.push(k); } catch(e) { del.push(k); }
    });
    if (del.length) PropertiesService.getScriptProperties().deleteProperties(del);
  } catch(e) {}
}

function buscarEnSheet(normUsuario, clave) {
  var tabs = ['usuarios','cabinets','gerencial','mapa','costos'];
  var ss;
  try { ss = SpreadsheetApp.openById(USUARIOS_SHEET_ID); } catch(e) { return null; }
  for (var i = 0; i < tabs.length; i++) {
    var sh = ss.getSheetByName(tabs[i]);
    if (!sh) continue;
    var last = sh.getLastRow();
    if (last < 2) continue;
    var rows = sh.getRange(2, 1, last-1, 3).getValues();
    for (var j = 0; j < rows.length; j++) {
      if (!rows[j][0]) continue;
      if (normUser(rows[j][0]) === normUsuario && String(rows[j][1]).trim() === clave) {
        return { usuario: String(rows[j][0]).trim().toUpperCase(), rol: String(rows[j][2]||'logistica').trim().toLowerCase() };
      }
    }
  }
  return null;
}

function getUsuarios()          { return leerUsuariosTab('cabinets'); }
function getUsuariosGerencial() { return leerUsuariosTab('gerencial'); }
function getUsuariosMapa()      { return leerUsuariosTab('mapa'); }
function getUsuariosCostos()    { return leerUsuariosTab('costos'); }

function leerUsuariosTab(tab) {
  try {
    var ss = SpreadsheetApp.openById(USUARIOS_SHEET_ID);
    var sh = ss.getSheetByName(tab);
    if (!sh) return [];
    var last = sh.getLastRow();
    if (last < 2) return [];
    return sh.getRange(2,1,last-1,3).getValues()
      .filter(function(r){ return !!r[0]; })
      .map(function(r){ return { usuario:String(r[0]).trim().toUpperCase(), clave:String(r[1]).trim(), rol:String(r[2]||'').trim().toLowerCase() }; });
  } catch(e) { return []; }
}

// ============================================================
// DATOS
// ============================================================

function getEquipos() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.EQUIPOS);
  var last = sh.getLastRow();
  if (last < 2) return {};
  var result = {};
  sh.getRange(2,1,last-1,15).getValues().forEach(function(r) {
    if (!r[0]) return;
    var placa = String(r[0]);
    result[placa] = {
      placa:placa, modelo:r[1]||'', estado:r[2]||'', etapa:r[3]||'',
      obs:r[4]||'', cliente:r[5]||'', delegado:r[6]||'',
      tsIngresoBodega:fmtTs(r[7]), tsEtapa:fmtTs(r[8]),
      tsSalidaMercado:fmtTs(r[9]), tsUltMov:fmtTs(r[10]),
      trabajos:splitList(r[11]), fotosEntrada:parseJson(r[12]),
      fotosSalida:parseJson(r[13]), fotosContrato:parseJson(r[14]), histEtapas:[]
    };
  });
  return result;
}

function getMovimientos() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.MOVIMIENTOS);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var start = Math.max(2, last-999);
  return sh.getRange(start,1,last-start+1,16).getValues()
    .filter(function(r){ return !!r[0]; })
    .map(function(r){
      return { tipo:r[0], ts:fmtTs(r[1]), placa:String(r[2]||''), modelo:r[3]||'',
        codCli:r[4]||'', nomCli:r[5]||'', cliente:r[6]||'', delegado:r[7]||'',
        fechaPlan:fmtDate(r[8]), numPlan:r[9]||null, etapa:r[10]||'',
        estado:r[11]||'', trabajos:splitList(r[12]), obs:r[13]||'',
        fotosEntrada:parseJson(r[14]), fotosContrato:parseJson(r[15]) };
    });
}

function getMantenimientos() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.MANTENIMIENTO);
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2,1,last-1,7).getValues()
    .filter(function(r){ return !!r[0]; })
    .map(function(r){ return { placa:String(r[0]), modelo:r[1]||'', etapa:r[2]||'', tsInicio:fmtTs(r[3]), tsFin:fmtTs(r[4]), dias:r[5]||0, trabajos:splitList(r[6]) }; });
}

function guardarMovimiento(mov) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.MOVIMIENTOS).appendRow([
    mov.tipo||'', mov.ts||new Date().toISOString(), mov.placa||'', mov.modelo||'',
    mov.codCli||'', mov.nomCli||'', mov.cliente||'', mov.delegado||'',
    mov.fechaPlan||'', mov.numPlan||'', mov.etapa||'', mov.estado||'',
    Array.isArray(mov.trabajos)?mov.trabajos.join('|'):'', mov.obs||'',
    JSON.stringify(mov.fotosEntrada||{}), JSON.stringify(mov.fotosContrato||{})
  ]);
}

function guardarMovimientoSiNoExiste(mov) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.MOVIMIENTOS);
  if (findMovimientoRow(sh, mov.placa, mov.ts) === -1) guardarMovimiento(mov);
}

function findMovimientoRow(sh, placa, ts) {
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var rows = sh.getRange(2,1,last-1,3).getValues();
  for (var i = rows.length-1; i >= 0; i--) {
    if (String(rows[i][0])==='SALIDA' && String(rows[i][2])===String(placa) && fmtTs(rows[i][1])===String(ts)) return i+2;
  }
  return -1;
}

function cancelarPlan(placa, ts) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.MOVIMIENTOS);
  var row = findMovimientoRow(sh, placa, ts);
  if (row !== -1) sh.deleteRow(row);
}

function buildEquipoFromMov(mov) {
  return { placa:mov.placa, modelo:mov.modelo||'', estado:mov.estado||'MANTENIMIENTO', etapa:mov.etapa||'', obs:mov.obs||'', cliente:'', delegado:mov.delegado||'', tsIngresoBodega:mov.ts||'', tsEtapa:mov.ts||'', tsSalidaMercado:'', tsUltMov:mov.ts||'', trabajos:mov.trabajos||[], fotosEntrada:mov.fotosEntrada||{}, fotosSalida:{}, fotosContrato:{} };
}

function upsertEquipo(eq) {
  if (!eq||!eq.placa) return;
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.EQUIPOS);
  var row = findEquipoRow(sh, eq.placa);
  var values = [eq.placa, eq.modelo||'', eq.estado||'', eq.etapa||'', eq.obs||'', eq.cliente||'', eq.delegado||'', eq.tsIngresoBodega||'', eq.tsEtapa||'', eq.tsSalidaMercado||'', eq.tsUltMov||'', Array.isArray(eq.trabajos)?eq.trabajos.join('|'):(eq.trabajos||''), JSON.stringify(eq.fotosEntrada||{}), JSON.stringify(eq.fotosSalida||{}), JSON.stringify(eq.fotosContrato||{})];
  if (row === -1) sh.appendRow(values);
  else sh.getRange(row,1,1,values.length).setValues([values]);
}

function upsertEquipoCampo(placa, campos) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.EQUIPOS);
  var row = findEquipoRow(sh, placa);
  if (row === -1) throw new Error('No existe la placa '+placa+' en Equipos');
  var cols = {modelo:2,estado:3,etapa:4,obs:5,cliente:6,delegado:7,tsIngresoBodega:8,tsEtapa:9,tsSalidaMercado:10,tsUltMov:11,trabajos:12,fotosContrato:15};
  Object.keys(campos).forEach(function(k) {
    if (!cols[k]) return;
    var value = campos[k];
    if (Array.isArray(value)) value = value.join('|');
    sh.getRange(row, cols[k]).setValue(value);
  });
}

function eliminarEquipo(placa) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.EQUIPOS);
  var row = findEquipoRow(sh, placa);
  if (row !== -1) sh.deleteRow(row);
}

function findEquipoRow(sh, placa) {
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var values = sh.getRange(2,1,last-1,1).getValues();
  for (var i = 0; i < values.length; i++) if (String(values[i][0])===String(placa)) return i+2;
  return -1;
}

function despacharSalida(data, planTs, fotosContratoRaw) {
  var urlsContrato = subirContratoADrive(data.placa, fotosContratoRaw||{});
  data.fotosContrato = urlsContrato;
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.MOVIMIENTOS);
  var row = findMovimientoRow(sh, data.placa, planTs);
  var despachoTs = data.ts||new Date().toISOString();
  if (row === -1) { data.tipo='SALIDA'; data.estado='DESPACHADO'; data.ts=despachoTs; guardarMovimiento(data); }
  else {
    sh.getRange(row,2).setValue(despachoTs); sh.getRange(row,12).setValue('DESPACHADO');
    if (Object.keys(urlsContrato).length) sh.getRange(row,16).setValue(JSON.stringify(urlsContrato));
  }
  marcarEquipoMercado(data, urlsContrato);
}

function marcarEquipoMercado(data, fotosContrato) {
  upsertEquipoCampo(data.placa, { estado:'MERCADO', cliente:(data.codCli||'')+' - '+(data.nomCli||''), tsSalidaMercado:data.ts||new Date().toISOString(), tsUltMov:data.ts||new Date().toISOString() });
  if (fotosContrato&&Object.keys(fotosContrato).length) guardarFotosContratoEquipo(data.placa, fotosContrato);
}

function finalizarMantenimiento(payload) {
  var fotos = subirFotosADrive(payload.placa,'salida',payload.fotosSalida||{});
  upsertEquipoCampo(payload.placa, {estado:'BODEGA_OK',etapa:'',trabajos:'',tsUltMov:payload.ts||new Date().toISOString()});
  if (Object.keys(fotos).length) guardarFotosSalidaEquipo(payload.placa, fotos);
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.MANTENIMIENTO).appendRow([payload.placa||'',payload.modelo||'',payload.etapa||'',payload.tsInicio||'',payload.ts||'',payload.dias||0,Array.isArray(payload.trabajos)?payload.trabajos.join('|'):'']);
}

function corregirEtapa(payload) {
  var ts = payload.ts||new Date().toISOString();
  var campos = {estado:payload.estado||'MANTENIMIENTO',etapa:payload.etapa||'',tsUltMov:ts};
  if ((payload.estado||'')==='DAÑADO') campos.tsEtapa='';
  upsertEquipoCampo(payload.placa, campos);
}

function actualizarContrato(payload) {
  var incoming = payload.fotosContrato||{};
  var nuevas = subirContratoADrive(payload.placa, incoming);
  var merged = {};
  ['f','r'].forEach(function(k){
    if(nuevas[k]) merged[k]=nuevas[k];
    else if(incoming[k]&&String(incoming[k]).indexOf('base64,')===-1) merged[k]=incoming[k];
  });
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.MOVIMIENTOS);
  var row = findMovimientoRow(sh, payload.placa, payload.ts);
  if (row !== -1) sh.getRange(row,16).setValue(JSON.stringify(merged));
  guardarFotosContratoEquipo(payload.placa, merged);
}

function subirFotosADrive(placa, tipo, fotos) {
  if (!fotos||!placa) return {};
  var root=getOrCreateFolder(DRIVE_FOLDER_NAME), folder=getOrCreateFolder(String(placa),root);
  var labels={f:'frontal',t:'trasera',i:'interior',p:'placa'};
  var stamp=Utilities.formatDate(new Date(),TZ,'yyyy-MM-dd_HHmmss'), urls={};
  Object.keys(fotos).forEach(function(key){
    var value=fotos[key];
    if(!value||value.indexOf('base64,')===-1) return;
    var blob=Utilities.newBlob(Utilities.base64Decode(value.split('base64,')[1]),'image/jpeg',tipo+'_'+(labels[key]||key)+'_'+stamp+'.jpg');
    var file=folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);
    urls[key]=file.getUrl();
  });
  return urls;
}

function subirContratoADrive(placa, fotos) {
  if (!fotos||!placa) return {};
  var root=getOrCreateFolder(DRIVE_CONTRATOS_NAME), folder=getOrCreateFolder(String(placa),root);
  var labels={f:'frontal',r:'reverso'};
  var stamp=Utilities.formatDate(new Date(),TZ,'yyyy-MM-dd_HHmmss'), urls={};
  Object.keys(fotos).forEach(function(key){
    var value=fotos[key];
    if(!value||value.indexOf('base64,')===-1) return;
    var blob=Utilities.newBlob(Utilities.base64Decode(value.split('base64,')[1]),'image/jpeg','contrato_'+(labels[key]||key)+'_'+stamp+'.jpg');
    var file=folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK,DriveApp.Permission.VIEW);
    urls[key]=file.getUrl();
  });
  return urls;
}

function getOrCreateFolder(name, parent) {
  var base=parent||getAppRoot(), found=base.getFoldersByName(name);
  return found.hasNext()?found.next():base.createFolder(name);
}

function getAppRoot() {
  var found=DriveApp.getFoldersByName(APP_ROOT_NAME);
  return found.hasNext()?found.next():DriveApp.getRootFolder().createFolder(APP_ROOT_NAME);
}

function guardarFotosSalidaEquipo(placa,urls){var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.EQUIPOS);var row=findEquipoRow(sh,placa);if(row!==-1)sh.getRange(row,14).setValue(JSON.stringify(urls));}
function guardarFotosContratoEquipo(placa,urls){var sh=SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.EQUIPOS);var row=findEquipoRow(sh,placa);if(row!==-1)sh.getRange(row,15).setValue(JSON.stringify(urls));}

function ensureSheets() {
  var ss=SpreadsheetApp.getActiveSpreadsheet(), defs={};
  defs[SHEETS.EQUIPOS]      =['placa','modelo','estado','etapa','obs','cliente','delegado','tsIngresoBodega','tsEtapa','tsSalidaMercado','tsUltMov','trabajos','fotosEntrada','fotosSalida','fotosContrato'];
  defs[SHEETS.MOVIMIENTOS]  =['tipo','ts','placa','modelo','codCli','nomCli','cliente','delegado','fechaPlan','numPlan','etapa','estado','trabajos','obs','fotosEntrada','fotosContrato'];
  defs[SHEETS.MANTENIMIENTO]=['placa','modelo','etapa','tsInicio','tsFin','dias','trabajos'];
  Object.keys(defs).forEach(function(name){ var sh=ss.getSheetByName(name); if(!sh){sh=ss.insertSheet(name);sh.appendRow(defs[name]);sh.setFrozenRows(1);} });
}

function normUser(v){return String(v||'').trim().toUpperCase().normalize('NFD').replace(/[̀-ͯ]/g,'');}
function splitList(v){return v?String(v).split('|').filter(Boolean):[];}
function parseJson(v){try{return JSON.parse(v||'{}');}catch(e){return {};}}
function fmtTs(v){return !v?'':(v instanceof Date?v.toISOString():String(v));}
function fmtDate(v){if(!v)return '';if(typeof v==='string'&&/^\d{4}-\d{2}-\d{2}/.test(v))return v.substring(0,10);return Utilities.formatDate(new Date(v),TZ,'yyyy-MM-dd');}
