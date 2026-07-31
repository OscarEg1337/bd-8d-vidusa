// BD 8D - VIDUSA | Google Apps Script Web App
// Deploy: Execute as "Me" → Anyone (no sign-in required)
// CORS is handled automatically by the GAS runtime for public deployments.

var SHEET_NAME          = 'BD_8D';
var DESPIVOTADA_SHEET_NAME = 'BD_8D_Despivotada';
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
  'Evidencia_Link_D8','Creado_Por','Votacion_D4_JSON',
  'Bitacora_Accion_1','Fecha_Bitacora_1','Bitacora_Accion_2','Fecha_Bitacora_2',
  'Bitacora_Accion_3','Fecha_Bitacora_3','Bitacora_Accion_4','Fecha_Bitacora_4',
  'Bitacora_Accion_5','Fecha_Bitacora_5','Bitacora_Accion_6','Fecha_Bitacora_6',
  'Bitacora_Accion_7','Fecha_Bitacora_7','Bitacora_Accion_8','Fecha_Bitacora_8'
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
    // Hoja vacía: escribir encabezados
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  } else {
    // Hoja con datos: verificar si la fila 1 coincide con HEADERS
    var currentHeaders = sheet.getRange(1, 1, 1, HEADERS.length).getValues()[0];
    var headersMatch = HEADERS.every(function(h, i) { return currentHeaders[i] === h; });
    if (!headersMatch) {
      // Encabezados viejos: reemplazar fila 1 con los nuevos
      sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
      sheet.setFrozenRows(1);
    }
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

// ── DESPIVOTAR ──────────────────────────────────────────────────────
// Genera/actualiza la hoja "BD_8D_Despivotada": una fila por cada Acción
// Correctiva (1-8) que tenga contenido, con los datos base del folio
// repetidos — pensada para consumirse en Power BI. Se regenera completa
// cada vez (no se acumula), así siempre refleja el estado actual de BD_8D.
var DESPIVOTADA_BASE_FIELDS = [
  'ID_Registro','Folio','Fraccionamiento','Fecha_Revision','Superintendente',
  'Residente','Facilitador_BPO','Descripcion_Problema','Accion_Contencion',
  'Causa_Raiz','Ponderacion_Causa','Estatus_Folio_D6','Estatus_Folio_D8','Creado_Por'
];

function generarDespivotada() {
  var sheet = getSheet();
  var lr = sheet.getLastRow();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var out = ss.getSheetByName(DESPIVOTADA_SHEET_NAME);
  if (!out) out = ss.insertSheet(DESPIVOTADA_SHEET_NAME);
  out.clear();

  var outHeaders = DESPIVOTADA_BASE_FIELDS.concat([
    'Numero_Accion','Accion_Correctiva','Responsable_Accion',
    'Fecha_Programada','Fecha_Realizacion','Bitacora_Accion','Fecha_Bitacora'
  ]);

  var rows = [];
  if (lr >= 2) {
    var values = sheet.getRange(2, 1, lr - 1, HEADERS.length).getValues();
    values.forEach(function(row) {
      var obj = rowToObj(row);
      if (!obj.ID_Registro) return;
      var base = DESPIVOTADA_BASE_FIELDS.map(function(f) { return obj[f] || ''; });
      for (var n = 1; n <= 8; n++) {
        var accion = obj['Accion_Correctiva_' + n];
        if (!accion) continue; // solo acciones con contenido capturado
        rows.push(base.concat([
          n,
          accion,
          obj['Responsable_Accion_' + n] || '',
          obj['Fecha_Programada_' + n] || '',
          obj['Fecha_Realizacion_' + n] || '',
          obj['Bitacora_Accion_' + n] || '',
          obj['Fecha_Bitacora_' + n] || ''
        ]));
      }
    });
  }

  out.getRange(1, 1, 1, outHeaders.length).setValues([outHeaders]);
  out.setFrozenRows(1);
  if (rows.length) out.getRange(2, 1, rows.length, outHeaders.length).setValues(rows);
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

  // ADD_USER: append new user to USUARIOS sheet
  if (method === 'ADD_USER') {
    try {
      var newUname  = String(body.username || '').trim();
      var newUpass  = String(body.password || '').trim();
      var newNombre = String(body.nombre   || '').trim();
      var newRol    = String(body.rol      || '').trim();
      if (!newUname || !newUpass || !newNombre || !newRol) {
        return jsonOut({ status: 'error', message: 'Todos los campos son requeridos' });
      }
      var validRoles = ['Admin', 'Editor', 'Viewer'];
      if (validRoles.indexOf(newRol) === -1) {
        return jsonOut({ status: 'error', message: 'Rol inválido. Usa Admin, Editor o Viewer' });
      }
      var usheet2 = getUsuariosSheet();
      var ulr2    = usheet2.getLastRow();
      if (ulr2 >= 2) {
        var existing = usheet2.getRange(2, 1, ulr2 - 1, 1).getValues();
        for (var j = 0; j < existing.length; j++) {
          if (String(existing[j][0]).trim().toLowerCase() === newUname.toLowerCase()) {
            return jsonOut({ status: 'error', message: 'El username "' + newUname + '" ya existe' });
          }
        }
      }
      usheet2.appendRow([newUname, newUpass, newNombre, newRol, '']);
      return jsonOut({ status: 'ok', message: 'Usuario creado' });
    } catch (err) {
      return jsonOut({ status: 'error', message: err.message });
    }
  }

  // CRUD operations
  try {
    var sheet = getSheet();

    if (method === 'POST') {
      // El ID_Registro/Folio se asignan aquí bajo lock, no se confía en lo que
      // manda el navegador: si dos usuarios crean un registro casi al mismo tiempo,
      // cada uno con su propia copia local desactualizada, ambos podrían calcular
      // el mismo número y terminar con dos filas con el mismo ID_Registro.
      var lock = LockService.getScriptLock();
      lock.waitLock(15000);
      try {
        var lr2 = sheet.getLastRow();
        var existingIds = lr2 >= 2
          ? sheet.getRange(2, 1, lr2 - 1, 1).getValues().map(function(row) { return String(row[0]); })
          : [];
        var yr = new Date().getFullYear();
        var n = lr2;
        var folio, idReg;
        do {
          folio = String(n).padStart(3, '0');
          idReg = '8D-' + yr + '-' + folio;
          n++;
        } while (existingIds.indexOf(idReg) !== -1);
        var newRow = HEADERS.map(function(h) {
          if (h === 'ID_Registro') return idReg;
          if (h === 'Folio') return folio;
          return body[h] !== undefined ? body[h] : '';
        });
        sheet.appendRow(newRow);
        generarDespivotada();
        return jsonOut({ status: 'ok', action: 'created', id: idReg, folio: folio });
      } finally {
        lock.releaseLock();
      }
    }

    if (method === 'PUT') {
      var id = body.ID_Registro;
      if (!id) return jsonOut({ status: 'error', message: 'ID_Registro requerido' });
      var rowNum = findRowById(sheet, id);
      if (rowNum === -1) return jsonOut({ status: 'error', message: 'No encontrado: ' + id });
      // Merge: si el valor entrante viene vacío/indefinido pero la hoja ya tiene un dato
      // guardado (p.ej. otro usuario lo capturó después de que este cliente cargó el registro),
      // se conserva el valor existente en vez de borrarlo con una cadena vacía.
      var currentRow = sheet.getRange(rowNum, 1, 1, HEADERS.length).getValues()[0];
      var updRow = HEADERS.map(function(h, i) {
        var incoming = body[h];
        if (incoming === undefined || incoming === '') return currentRow[i];
        return incoming;
      });
      sheet.getRange(rowNum, 1, 1, HEADERS.length).setValues([updRow]);
      generarDespivotada();
      return jsonOut({ status: 'ok', action: 'updated', id: id });
    }

    if (method === 'DELETE') {
      var delId = body.ID_Registro;
      if (!delId) return jsonOut({ status: 'error', message: 'ID_Registro requerido' });
      var delRow = findRowById(sheet, delId);
      if (delRow === -1) return jsonOut({ status: 'error', message: 'No encontrado: ' + delId });
      sheet.deleteRow(delRow);
      generarDespivotada();
      return jsonOut({ status: 'ok', action: 'deleted', id: delId });
    }

    return jsonOut({ status: 'error', message: 'Método desconocido: ' + method });

  } catch (err) {
    return jsonOut({ status: 'error', message: err.message });
  }
}

// ── MANTENIMIENTO: ejecutar UNA VEZ manualmente desde el editor de Apps ────
// Script (seleccionar esta función en el desplegable de arriba → Ejecutar)
// para corregir folios que hayan quedado con el mismo ID_Registro duplicado
// (causa de que al editar un folio se abrieran los datos de otro).
function repararIdsDuplicados() {
  var sheet = getSheet();
  var lr = sheet.getLastRow();
  if (lr < 2) { Logger.log('Sin registros.'); return; }

  var range = sheet.getRange(2, 1, lr - 1, 2); // columnas ID_Registro, Folio
  var values = range.getValues();
  var yr = new Date().getFullYear();

  var maxNum = 0;
  values.forEach(function(row) {
    var m = String(row[0]).match(/(\d+)$/);
    if (m) maxNum = Math.max(maxNum, parseInt(m[1], 10));
  });

  var seen = {};
  var changed = false;
  for (var i = 0; i < values.length; i++) {
    var id = String(values[i][0]);
    if (!id) continue;
    if (seen[id]) {
      maxNum++;
      var newFolio = String(maxNum).padStart(3, '0');
      var newId = '8D-' + yr + '-' + newFolio;
      Logger.log('Duplicado: fila ' + (i + 2) + ' tenía ' + id + ' → reasignado a ' + newId);
      values[i][0] = newId;
      values[i][1] = newFolio;
      changed = true;
    } else {
      seen[id] = true;
    }
  }

  if (changed) {
    range.setValues(values);
    Logger.log('Listo: se corrigieron IDs duplicados. Revisa el registro de ejecución arriba.');
  } else {
    Logger.log('No se encontraron IDs duplicados.');
  }
}
