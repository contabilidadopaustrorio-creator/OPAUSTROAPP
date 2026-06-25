const SHEETS = {
  EQUIPOS: 'Equipos',
  MOVIMIENTOS: 'Movimientos',
  MANTENIMIENTO: 'Mantenimiento'
};

const DRIVE_FOLDER_NAME = 'Control Cabinets Fotos';
const DRIVE_CONTRATOS_NAME = 'Captaciones Contratos';
const APP_ROOT_NAME = 'APP OPAUSTRO';  // carpeta principal que contiene todo
const TZ = 'America/Guayaquil';
const USUARIOS_SHEET_ID = '15McucuG8iCpDVPAyGVW9-RsrdpP2IdOMa03kjL-vtr8';
const USUARIOS_TAB = 'cabinets';
const USUARIOS_TAB_GERENCIAL = 'gerencial';
const USUARIOS_TAB_MAPA = 'mapa';
const USUARIOS_TAB_COSTOS = 'costos';

function doGet(e) {
  try {
    var action = (e && e.parameter && e.parameter.action) || '';
    if (action === 'getUsuariosGerencial') {
      return jsonResponse({usuarios: getUsuariosGerencial()});
    }
    if (action === 'getUsuariosMapa') {
      return jsonResponse({usuarios: getUsuariosMapa()});
    }
    if (action === 'getUsuariosCostos') {
      return jsonResponse({usuarios: getUsuariosCostos()});
    }
    if (action === 'getUsuarios') {
      return jsonResponse({usuarios: getUsuarios()});
    }
    ensureSheets();
    return jsonResponse({db: {
      equipos: getEquipos(),
      movimientos: getMovimientos(),
      mantenimientos: getMantenimientos()
    }, usuarios: getUsuarios()});
  } catch (err) {
    return jsonResponse({error: err.message});
  }
}

function getUsuarios() {
  return leerUsuariosTab(USUARIOS_TAB);
}

function getUsuariosGerencial() {
  return leerUsuariosTab(USUARIOS_TAB_GERENCIAL);
}

function getUsuariosMapa() {
  return leerUsuariosTab(USUARIOS_TAB_MAPA);
}

function getUsuariosCostos() {
  return leerUsuariosTab(USUARIOS_TAB_COSTOS);
}

function leerUsuariosTab(tab) {
  try {
    var ss = SpreadsheetApp.openById(USUARIOS_SHEET_ID);
    var sh = ss.getSheetByName(tab);
    if (!sh) return [];
    var last = sh.getLastRow();
    if (last < 2) return [];
    return sh.getRange(2, 1, last - 1, 3).getValues()
      .filter(function(r) { return !!r[0]; })
      .map(function(r) {
        return {
          usuario: String(r[0]).trim().toUpperCase(),
          clave: String(r[1]).trim(),
          rol: String(r[2]).trim().toLowerCase()
        };
      });
  } catch(e) {
    return [];
  }
}

function doGetUsuariosGerencial() {
  return jsonResponse({usuarios: getUsuariosGerencial()});
}

function doPost(e) {
  try {
    ensureSheets();
    var payload = JSON.parse(e.postData.contents);
    var action = payload.action;
    var data = payload.data || {};

    if (action === 'entrada') {
      var fotosEntrada = payload.fotosEntrada || data.fotosEntrada || {};
      data.fotosEntrada = subirFotosADrive(data.placa, 'entrada', fotosEntrada);
      guardarMovimiento(data);
      upsertEquipo(buildEquipoFromMov(data));
    } else if (action === 'planificarSalida') {
      data.tipo = 'SALIDA';
      data.estado = 'PLANIFICADO';
      guardarMovimiento(data);
    } else if (action === 'cancelarPlan') {
      cancelarPlan(data.placa, data.ts);
    } else if (action === 'actualizarFotos') {
      var incomingF = payload.fotosSalida || {};
      var nuevasF = subirFotosADrive(payload.placa, 'salida', incomingF); // solo sube las base64 nuevas
      var mergedF = {};
      ['f', 't', 'i'].forEach(function(k) {
        if (nuevasF[k]) mergedF[k] = nuevasF[k];                                  // foto nueva subida
        else if (incomingF[k] && String(incomingF[k]).indexOf('base64,') === -1) mergedF[k] = incomingF[k]; // URL existente que se conserva
      });
      guardarFotosSalidaEquipo(payload.placa, mergedF);
      upsertEquipoCampo(payload.placa, {tsUltMov: payload.ts || new Date().toISOString()});
    } else if (action === 'despacharSalida') {
      despacharSalida(data, payload.planTs || data.tsPlanificacion || '', payload.fotosContrato || {});
    } else if (action === 'salida') {
      data.tipo = 'SALIDA';
      data.estado = 'DESPACHADO';
      guardarMovimiento(data);
      marcarEquipoMercado(data, {});
    } else if (action === 'finalizarMant') {
      finalizarMantenimiento(payload);
    } else if (action === 'bajaEquipo') {
      guardarMovimientoSiNoExiste(data);
      eliminarEquipo(data.placa);
    } else if (action === 'corregirEtapa') {
      corregirEtapa(payload);
    } else if (action === 'actualizarContrato') {
      actualizarContrato(payload);
    } else if (action === 'enviarQuito') {
      guardarMovimientoSiNoExiste(data);
      upsertEquipoCampo(data.placa, {
        estado: 'MANTENIMIENTO', etapa: data.etapa || 'C',
        cliente: '', tsEtapa: data.ts, tsUltMov: data.ts,
        trabajos: data.trabajos || ['ENVIO QUITO']
      });
    } else if (action === 'cargarBase') {
      (payload.equipos || []).forEach(function(eq) { upsertEquipo(eq); });
    } else {
      throw new Error('Accion no reconocida: ' + action);
    }

    return jsonResponse({ok: true});
  } catch (err) {
    return jsonResponse({ok: false, error: err.message});
  }
}

function despacharSalida(data, planTs, fotosContratoRaw) {
  // Subir fotos del contrato a Drive en carpeta "Captaciones Contratos/<placa>/"
  var urlsContrato = subirContratoADrive(data.placa, fotosContratoRaw || {});
  data.fotosContrato = urlsContrato;

  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.MOVIMIENTOS);
  var row = findMovimientoRow(sh, data.placa, planTs);
  var despachoTs = data.ts || new Date().toISOString();

  if (row === -1) {
    data.tipo = 'SALIDA';
    data.estado = 'DESPACHADO';
    data.ts = despachoTs;
    guardarMovimiento(data);
  } else {
    sh.getRange(row, 2).setValue(despachoTs);
    sh.getRange(row, 12).setValue('DESPACHADO');
    if (Object.keys(urlsContrato).length) {
      sh.getRange(row, 16).setValue(JSON.stringify(urlsContrato));
    }
  }
  marcarEquipoMercado(data, urlsContrato);
}

function marcarEquipoMercado(data, fotosContrato) {
  upsertEquipoCampo(data.placa, {
    estado: 'MERCADO',
    cliente: (data.codCli || '') + ' - ' + (data.nomCli || ''),
    tsSalidaMercado: data.ts || new Date().toISOString(),
    tsUltMov: data.ts || new Date().toISOString()
  });
  if (fotosContrato && Object.keys(fotosContrato).length) {
    guardarFotosContratoEquipo(data.placa, fotosContrato);
  }
}

function finalizarMantenimiento(payload) {
  var fotos = subirFotosADrive(payload.placa, 'salida', payload.fotosSalida || {});
  upsertEquipoCampo(payload.placa, {
    estado: 'BODEGA_OK', etapa: '', trabajos: '',
    tsUltMov: payload.ts || new Date().toISOString()
  });
  if (Object.keys(fotos).length) guardarFotosSalidaEquipo(payload.placa, fotos);
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.MANTENIMIENTO).appendRow([
    payload.placa || '', payload.modelo || '', payload.etapa || '',
    payload.tsInicio || '', payload.ts || '', payload.dias || 0,
    Array.isArray(payload.trabajos) ? payload.trabajos.join('|') : ''
  ]);
}

// Corrige la etapa de un equipo en mantenimiento si se registró por error.
// A/B/C: solo cambia la etapa (conserva tsEtapa, no reinicia los días).
// DAÑADO: saca el equipo de mantenimiento y limpia tsEtapa.
function corregirEtapa(payload) {
  var ts = payload.ts || new Date().toISOString();
  var campos = {
    estado: payload.estado || 'MANTENIMIENTO',
    etapa: payload.etapa || '',
    tsUltMov: ts
  };
  if ((payload.estado || '') === 'DAÑADO') campos.tsEtapa = '';
  upsertEquipoCampo(payload.placa, campos);
}

// Sube o reemplaza las fotos del contrato de un despacho ya registrado.
// Actualiza el movimiento (col 16) y el equipo (col 15). Conserva las fotos
// que ya estaban (URL) y sube solo las nuevas (base64).
function actualizarContrato(payload) {
  var placa = payload.placa;
  var ts = payload.ts;
  var incoming = payload.fotosContrato || {};
  var nuevas = subirContratoADrive(placa, incoming); // sube solo las base64 nuevas
  var merged = {};
  ['f', 'r'].forEach(function(k) {
    if (nuevas[k]) merged[k] = nuevas[k];                                          // foto nueva subida
    else if (incoming[k] && String(incoming[k]).indexOf('base64,') === -1) merged[k] = incoming[k]; // URL existente
  });
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.MOVIMIENTOS);
  var row = findMovimientoRow(sh, placa, ts);
  if (row !== -1) sh.getRange(row, 16).setValue(JSON.stringify(merged));
  guardarFotosContratoEquipo(placa, merged);
}

function getEquipos() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.EQUIPOS);
  var last = sh.getLastRow();
  if (last < 2) return {};
  var result = {};
  sh.getRange(2, 1, last - 1, 15).getValues().forEach(function(r) {
    if (!r[0]) return;
    var placa = String(r[0]);
    result[placa] = {
      placa: placa, modelo: r[1] || '', estado: r[2] || '', etapa: r[3] || '',
      obs: r[4] || '', cliente: r[5] || '', delegado: r[6] || '',
      tsIngresoBodega: fmtTs(r[7]), tsEtapa: fmtTs(r[8]),
      tsSalidaMercado: fmtTs(r[9]), tsUltMov: fmtTs(r[10]),
      trabajos: splitList(r[11]), fotosEntrada: parseJson(r[12]),
      fotosSalida: parseJson(r[13]), fotosContrato: parseJson(r[14]),
      histEtapas: []
    };
  });
  return result;
}

function getMovimientos() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.MOVIMIENTOS);
  var last = sh.getLastRow();
  if (last < 2) return [];
  var start = Math.max(2, last - 999);
  return sh.getRange(start, 1, last - start + 1, 16).getValues()
    .filter(function(r) { return !!r[0]; })
    .map(function(r) {
      return {
        tipo: r[0], ts: fmtTs(r[1]), placa: String(r[2] || ''), modelo: r[3] || '',
        codCli: r[4] || '', nomCli: r[5] || '', cliente: r[6] || '', delegado: r[7] || '',
        fechaPlan: fmtDate(r[8]), numPlan: r[9] || null, etapa: r[10] || '',
        estado: r[11] || '', trabajos: splitList(r[12]), obs: r[13] || '',
        fotosEntrada: parseJson(r[14]), fotosContrato: parseJson(r[15])
      };
    });
}

function getMantenimientos() {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.MANTENIMIENTO);
  var last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, 7).getValues()
    .filter(function(r) { return !!r[0]; })
    .map(function(r) {
      return {placa: String(r[0]), modelo: r[1] || '', etapa: r[2] || '',
        tsInicio: fmtTs(r[3]), tsFin: fmtTs(r[4]), dias: r[5] || 0, trabajos: splitList(r[6])};
    });
}

function guardarMovimiento(mov) {
  SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.MOVIMIENTOS).appendRow([
    mov.tipo || '', mov.ts || new Date().toISOString(), mov.placa || '', mov.modelo || '',
    mov.codCli || '', mov.nomCli || '', mov.cliente || '', mov.delegado || '',
    mov.fechaPlan || '', mov.numPlan || '', mov.etapa || '', mov.estado || '',
    Array.isArray(mov.trabajos) ? mov.trabajos.join('|') : '', mov.obs || '',
    JSON.stringify(mov.fotosEntrada || {}),
    JSON.stringify(mov.fotosContrato || {})
  ]);
}

function guardarMovimientoSiNoExiste(mov) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.MOVIMIENTOS);
  if (findMovimientoRow(sh, mov.placa, mov.ts) === -1) guardarMovimiento(mov);
}

function findMovimientoRow(sh, placa, ts) {
  var last = sh.getLastRow();
  if (last < 2) return -1;
  var rows = sh.getRange(2, 1, last - 1, 3).getValues();
  for (var i = rows.length - 1; i >= 0; i--) {
    if (String(rows[i][0]) === 'SALIDA' && String(rows[i][2]) === String(placa) && fmtTs(rows[i][1]) === String(ts)) return i + 2;
  }
  return -1;
}

function cancelarPlan(placa, ts) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.MOVIMIENTOS);
  var row = findMovimientoRow(sh, placa, ts);
  if (row !== -1) sh.deleteRow(row);
}

function buildEquipoFromMov(mov) {
  return {
    placa: mov.placa, modelo: mov.modelo || '', estado: mov.estado || 'MANTENIMIENTO',
    etapa: mov.etapa || '', obs: mov.obs || '', cliente: '', delegado: mov.delegado || '',
    tsIngresoBodega: mov.ts || '', tsEtapa: mov.ts || '', tsSalidaMercado: '',
    tsUltMov: mov.ts || '', trabajos: mov.trabajos || [],
    fotosEntrada: mov.fotosEntrada || {}, fotosSalida: {}, fotosContrato: {}
  };
}

function upsertEquipo(eq) {
  if (!eq || !eq.placa) return;
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.EQUIPOS);
  var row = findEquipoRow(sh, eq.placa);
  var values = [eq.placa, eq.modelo || '', eq.estado || '', eq.etapa || '', eq.obs || '',
    eq.cliente || '', eq.delegado || '', eq.tsIngresoBodega || '', eq.tsEtapa || '',
    eq.tsSalidaMercado || '', eq.tsUltMov || '',
    Array.isArray(eq.trabajos) ? eq.trabajos.join('|') : (eq.trabajos || ''),
    JSON.stringify(eq.fotosEntrada || {}), JSON.stringify(eq.fotosSalida || {}),
    JSON.stringify(eq.fotosContrato || {})];
  if (row === -1) sh.appendRow(values);
  else sh.getRange(row, 1, 1, values.length).setValues([values]);
}

function upsertEquipoCampo(placa, campos) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.EQUIPOS);
  var row = findEquipoRow(sh, placa);
  if (row === -1) throw new Error('No existe la placa ' + placa + ' en Equipos');
  var cols = {modelo: 2, estado: 3, etapa: 4, obs: 5, cliente: 6, delegado: 7,
    tsIngresoBodega: 8, tsEtapa: 9, tsSalidaMercado: 10, tsUltMov: 11, trabajos: 12,
    fotosContrato: 15};
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
  var values = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) if (String(values[i][0]) === String(placa)) return i + 2;
  return -1;
}

function subirFotosADrive(placa, tipo, fotos) {
  if (!fotos || !placa) return {};
  var root = getOrCreateFolder(DRIVE_FOLDER_NAME);
  var folder = getOrCreateFolder(String(placa), root);
  var labels = {f: 'frontal', t: 'trasera', i: 'interior', p: 'placa'};
  var stamp = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd_HHmmss');
  var urls = {};
  Object.keys(fotos).forEach(function(key) {
    var value = fotos[key];
    if (!value || value.indexOf('base64,') === -1) return;
    var name = tipo + '_' + (labels[key] || key) + '_' + stamp + '.jpg';
    var blob = Utilities.newBlob(Utilities.base64Decode(value.split('base64,')[1]), 'image/jpeg', name);
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    urls[key] = file.getUrl();
  });
  return urls;
}

function subirContratoADrive(placa, fotos) {
  if (!fotos || !placa) return {};
  var root = getOrCreateFolder(DRIVE_CONTRATOS_NAME);
  var folder = getOrCreateFolder(String(placa), root);
  var labels = {f: 'frontal', r: 'reverso'};
  var stamp = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd_HHmmss');
  var urls = {};
  Object.keys(fotos).forEach(function(key) {
    var value = fotos[key];
    if (!value || value.indexOf('base64,') === -1) return;
    var name = 'contrato_' + (labels[key] || key) + '_' + stamp + '.jpg';
    var blob = Utilities.newBlob(Utilities.base64Decode(value.split('base64,')[1]), 'image/jpeg', name);
    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    urls[key] = file.getUrl();
  });
  return urls;
}

function getOrCreateFolder(name, parent) {
  var base = parent || getAppRoot();
  var found = base.getFoldersByName(name);
  return found.hasNext() ? found.next() : base.createFolder(name);
}

// Carpeta principal "APP OPAUSTRO". La busca en todo el Drive (donde sea
// que la hayas movido); si no existe la crea en la raíz una sola vez.
function getAppRoot() {
  var found = DriveApp.getFoldersByName(APP_ROOT_NAME);
  return found.hasNext() ? found.next() : DriveApp.getRootFolder().createFolder(APP_ROOT_NAME);
}

function guardarFotosSalidaEquipo(placa, urls) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.EQUIPOS);
  var row = findEquipoRow(sh, placa);
  if (row !== -1) sh.getRange(row, 14).setValue(JSON.stringify(urls));
}

function guardarFotosContratoEquipo(placa, urls) {
  var sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEETS.EQUIPOS);
  var row = findEquipoRow(sh, placa);
  if (row !== -1) sh.getRange(row, 15).setValue(JSON.stringify(urls));
}

function ensureSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var defs = {};
  defs[SHEETS.EQUIPOS] = ['placa','modelo','estado','etapa','obs','cliente','delegado',
    'tsIngresoBodega','tsEtapa','tsSalidaMercado','tsUltMov','trabajos',
    'fotosEntrada','fotosSalida','fotosContrato'];
  defs[SHEETS.MOVIMIENTOS] = ['tipo','ts','placa','modelo','codCli','nomCli','cliente',
    'delegado','fechaPlan','numPlan','etapa','estado','trabajos','obs',
    'fotosEntrada','fotosContrato'];
  defs[SHEETS.MANTENIMIENTO] = ['placa','modelo','etapa','tsInicio','tsFin','dias','trabajos'];
  Object.keys(defs).forEach(function(name) {
    var sh = ss.getSheetByName(name);
    if (!sh) { sh = ss.insertSheet(name); sh.appendRow(defs[name]); sh.setFrozenRows(1); }
  });
}

function splitList(value) { return value ? String(value).split('|').filter(Boolean) : []; }
function parseJson(value) { try { return JSON.parse(value || '{}'); } catch (e) { return {}; } }
function fmtTs(value) { return !value ? '' : (value instanceof Date ? value.toISOString() : String(value)); }
function fmtDate(value) {
  if (!value) return '';
  if (typeof value === 'string' && /^\d{4}-\d{2}-\d{2}/.test(value)) return value.substring(0, 10);
  return Utilities.formatDate(new Date(value), TZ, 'yyyy-MM-dd');
}
function jsonResponse(data) { return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON); }
