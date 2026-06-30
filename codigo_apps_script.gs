// BD 8D - VIDUSA | Google Apps Script Web App
// Deploy: Execute as "Me" → Anyone (no sign-in required)
// CORS is handled automatically by the GAS runtime for public deployments.

var SHEET_NAME          = 'BD_8D';
var USUARIOS_SHEET_NAME = 'USUARIOS';
var USUARIOS_HEADERS    = ['username','password','nombre','rol','fraccionamiento'];

var DEFAULT_USERS = [
  ['admin',       'admin123',  'Administrador',    'Admin',  ''],
  ['jose_agustin','vidusa2024','José Agustín',     'Editor', ''],
  ['ramiro',      'vidusa2024','Ramiro Fernández', 'Viewer', ''],
];

var HEADERS = [
  'ID_Registro','Folio','Fraccionamiento','Fecha_Revision','Superintendente',
  'Residente','Facilitador_BPO','Equipo_Trabajo','Descripcion_Problema',
  'Accion_Contencion','Causa_Raiz','Ponderacion_Causa',
  'Accion_Correctiva_1','Responsable_Accion_1','Fecha_Programada_1','Fecha_Realizacion_1',
  'Accion_Correctiva_2','Responsable_Accion_2','Fecha_Programada_2','Fecha_Realizacion_2',
  'Accion_Correctiva_3','Responsable_Accion_3','Fecha_Programada_3','Fecha_Realizacion_3',
  'Accion_Correctiva_4','Responsable_Accion_4','Fecha_Programada_4','Fecha_Realizacion_4',
  'Accion_Correctiva_5','Responsable_Accion_5','Fecha_Programada_5','Fecha_Realizacion_5',
  'Accion_Correctiva_6','Responsable_Accion_6','Fecha_Programada_6','Fecha_Realizacion_6',
  'Accion_Correctiva_7','Responsable_Accion_7','Fecha_Programada_7','Fecha_Realizacion_7',
  'Accion_Correctiva_8','Responsable_Accion_8','Fecha_Programada_8','Fecha_Realizacion_8',
  'Comentarios_Accion','Resultado_Accion','Verificacion_D6','FECHA_DE_REVISION_D6',
  'SATISFACCION_O_NO_SATISFACCION_D6','REVISION_1_D6','ResidenteD6',
  'Superintendente_D6','Facilitador_PMO_D6','Jefe_de_Calidad_D6','Estatus_Folio_D6',
  'Fecha_Cierre_D6','Evidencia_Link_D6','D7_Documentos_estandarizados',
  'Manual_de_Procesos','Plano','Modificacion_de_presupuesto','Responsable_D7',
  'Fecha_D7','D8_Evaluacion_de_efectividad','Fecha_de_revision_D8',
  'Satisfaccion_O_no_SatisfaccionD8','ResidenteD8','SuperintendenteD8',
  'Facilitador_PMO_D8','Jefe_de_Calidad_D8','Estatus_Folio_D8','Fecha_Cierre_D8',
  'Evidencia_Link_D8','Creado_Por'
];

// ── SHA-256 hex — solo para comparación de fallback en AUTH ──────
function sha256Hex(str) {
  var bytes = Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256, String(str), Utilities.Charset.UTF_8
  );
  return bytes.map(function(b) {
    return ('0' + (b < 0 ? b + 256 : b).toString(16)).slice(-2);
  }).join('');
}

// ── Response helper ───────────────────────────────────────────────
function jsonOut(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── Sheet helpers ─────────────────────────────────────────────────
function getUsuariosSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(USUARIOS_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(USUARIOS_SHEET_NAME);
    sheet.appendRow(USUARIOS_HEADERS);
    sheet.setFrozenRows(1);
    for (var i = 0; i < DEFAULT_USERS.length; i++) {
      sheet.appendRow(DEFAULT_USERS[i]);
    }
  }
  return sheet;
}

function getSheet() {
  var ss    = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) { sheet = ss.insertSheet(SHEET_NAME); }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function rowToObj(row) {
  var tz  = SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone();
  var obj = {};
  for (var i = 0; i < HEADERS.length; i++) {
    var v = row[i];
    obj[HEADERS[i]] = v instanceof Date
      ? Utilities.formatDate(v, tz, 'yyyy-MM-dd')
      : (v === null || v === undefined ? '' : String(v));
  }
  return obj;
}

function findRowById(sheet, id) {
  var lr = sheet.getLastRow();
  if (lr < 2) return -1;
  var ids = sheet.getRange(2, 1, lr - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return -1;
}

// ── doGet — return all records ────────────────────────────────────
function doGet(e) {
  try {
    var sheet   = getSheet();
    var lr      = sheet.getLastRow();
    var records = [];
    if (lr >= 2) {
      var values = sheet.getRange(2, 1, lr - 1, HEADERS.length).getValues();
      for (var i = 0; i < values.length; i++) {
        var obj = rowToObj(values[i]);
        if (obj.ID_Registro !== '') records.push(obj);
      }
    }
    return jsonOut({ status: 'ok', data: records });
  } catch (err) {
    return jsonOut({ status: 'error', message: err.message });
  }
}

// ── doPost — AUTH | CREATE | UPDATE | DELETE ──────────────────────
function doPost(e) {
  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonOut({ status: 'error', message: 'JSON inválido' });
  }

  var method = body._method ? String(body._method).toUpperCase() : 'POST';

  // AUTH: plain-text username/password lookup
  if (method === 'AUTH') {
    try {
      var uname  = String(body.username || '').trim();
      var upass  = String(body.password || '');
      var usheet = getUsuariosSheet();
      var ulr    = usheet.getLastRow();
      if (ulr < 2) return jsonOut({ status: 'error', message: 'Sin usuarios registrados' });
      var urows  = usheet.getRange(2, 1, ulr - 1, USUARIOS_HEADERS.length).getValues();
      for (var i = 0; i < urows.length; i++) {
        var urow   = urows[i];
        var stored = String(urow[1]);
        if (String(urow[0]).trim() !== uname) continue;
        // Intento 1: contraseña en texto plano (hoja normal)
        // Intento 2: hoja tiene hash SHA-256 (sembrada por versión anterior)
        var match = (stored === upass) || (stored === sha256Hex(upass));
        if (match) {
          return jsonOut({
            status: 'ok',
            user: {
              username: String(urow[0]).trim(),
              fullname: String(urow[2]),
              role:     String(urow[3]),
              fracc:    String(urow[4])
            }
          });
        }
        return jsonOut({ status: 'error', message: 'Usuario o contraseña incorrectos' });
      }
      return jsonOut({ status: 'error', message: 'Usuario o contraseña incorrectos' });
    } catch (err) {
      return jsonOut({ status: 'error', message: err.message });
    }
  }

  // CRUD operations
  try {
    var sheet = getSheet();

    if (method === 'POST') {
      if (!body.ID_Registro) return jsonOut({ status: 'error', message: 'ID_Registro requerido' });
      var newRow = HEADERS.map(function(h) { return body[h] !== undefined ? body[h] : ''; });
      sheet.appendRow(newRow);
      return jsonOut({ status: 'ok', action: 'created', id: body.ID_Registro });
    }

    if (method === 'PUT') {
      var id = body.ID_Registro;
      if (!id) return jsonOut({ status: 'error', message: 'ID_Registro requerido' });
      var rowNum = findRowById(sheet, id);
      if (rowNum === -1) return jsonOut({ status: 'error', message: 'No encontrado: ' + id });
      var updRow = HEADERS.map(function(h) { return body[h] !== undefined ? body[h] : ''; });
      sheet.getRange(rowNum, 1, 1, HEADERS.length).setValues([updRow]);
      return jsonOut({ status: 'ok', action: 'updated', id: id });
    }

    if (method === 'DELETE') {
      var delId = body.ID_Registro;
      if (!delId) return jsonOut({ status: 'error', message: 'ID_Registro requerido' });
      var delRow = findRowById(sheet, delId);
      if (delRow === -1) return jsonOut({ status: 'error', message: 'No encontrado: ' + delId });
      sheet.deleteRow(delRow);
      return jsonOut({ status: 'ok', action: 'deleted', id: delId });
    }

    return jsonOut({ status: 'error', message: 'Método desconocido: ' + method });

  } catch (err) {
    return jsonOut({ status: 'error', message: err.message });
  }
}
