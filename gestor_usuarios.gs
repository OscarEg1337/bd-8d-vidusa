// GESTOR DE USUARIOS — BD 8D VIDUSA
// 1. Pega este código en un Google Apps Script nuevo (script.google.com)
// 2. Cambia SPREADSHEET_ID por el ID de tu hoja de cálculo
// 3. Despliega: Implementar → Nueva implementación → Aplicación web
//    Ejecutar como: Yo | Quién tiene acceso: Cualquier usuario

var SPREADSHEET_ID    = 'PON_AQUI_EL_ID_DE_TU_SPREADSHEET';
var USUARIOS_SHEET    = 'USUARIOS';
var USUARIOS_HEADERS  = ['username', 'password', 'nombre', 'rol', 'fraccionamiento'];

function doGet() {
  return HtmlService.createHtmlOutput(getHTML())
    .setTitle('Gestor de Usuarios — VIDUSA')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  var body;
  try { body = JSON.parse(e.postData.contents); }
  catch(err) { return jsonOut({ status: 'error', message: 'JSON inválido' }); }

  var action = body.action || '';

  if (action === 'LIST') {
    return jsonOut(listUsers());
  }

  if (action === 'ADD') {
    var uname = String(body.username || '').trim();
    var upass  = String(body.password || '').trim();
    var nombre = String(body.nombre   || '').trim();
    var rol    = String(body.rol      || '').trim();
    var fracc  = String(body.fraccionamiento || '').trim();

    if (!uname || !upass || !nombre || !rol) {
      return jsonOut({ status: 'error', message: 'Todos los campos son requeridos' });
    }
    var validRoles = ['Admin', 'Editor', 'Viewer'];
    if (validRoles.indexOf(rol) === -1) {
      return jsonOut({ status: 'error', message: 'Rol inválido' });
    }

    var sheet = getSheet();
    var lr = sheet.getLastRow();
    if (lr >= 2) {
      var existing = sheet.getRange(2, 1, lr - 1, 1).getValues();
      for (var i = 0; i < existing.length; i++) {
        if (String(existing[i][0]).trim().toLowerCase() === uname.toLowerCase()) {
          return jsonOut({ status: 'error', message: 'El username "' + uname + '" ya existe' });
        }
      }
    }
    sheet.appendRow([uname, upass, nombre, rol, fracc]);
    return jsonOut({ status: 'ok', message: 'Usuario "' + uname + '" creado correctamente' });
  }

  if (action === 'DELETE') {
    var delUser = String(body.username || '').trim();
    if (!delUser) return jsonOut({ status: 'error', message: 'Username requerido' });
    var sheet2 = getSheet();
    var lr2 = sheet2.getLastRow();
    if (lr2 >= 2) {
      var rows = sheet2.getRange(2, 1, lr2 - 1, 1).getValues();
      for (var j = 0; j < rows.length; j++) {
        if (String(rows[j][0]).trim().toLowerCase() === delUser.toLowerCase()) {
          sheet2.deleteRow(j + 2);
          return jsonOut({ status: 'ok', message: 'Usuario eliminado' });
        }
      }
    }
    return jsonOut({ status: 'error', message: 'Usuario no encontrado' });
  }

  return jsonOut({ status: 'error', message: 'Acción desconocida' });
}

function listUsers() {
  var sheet = getSheet();
  var lr = sheet.getLastRow();
  var users = [];
  if (lr >= 2) {
    var rows = sheet.getRange(2, 1, lr - 1, USUARIOS_HEADERS.length).getValues();
    for (var i = 0; i < rows.length; i++) {
      if (rows[i][0]) {
        users.push({
          username: rows[i][0],
          nombre:   rows[i][2],
          rol:      rows[i][3],
          fracc:    rows[i][4]
        });
      }
    }
  }
  return { status: 'ok', users: users };
}

function getSheet() {
  var ss    = SpreadsheetApp.openById(SPREADSHEET_ID);
  var sheet = ss.getSheetByName(USUARIOS_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(USUARIOS_SHEET);
    sheet.appendRow(USUARIOS_HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function jsonOut(data) {
  return ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);
}

function getHTML() {
  return '<!DOCTYPE html><html><head><meta charset="UTF-8"><meta name="viewport" content="width=device-width,initial-scale=1"><title>Gestor Usuarios</title><style>' +
    '*{box-sizing:border-box;margin:0;padding:0}' +
    'body{font-family:system-ui,sans-serif;background:#f0f4f3;min-height:100vh;padding:24px}' +
    'h1{color:#1a3c38;font-size:20px;margin-bottom:20px}' +
    '.card{background:#fff;border-radius:10px;padding:24px;box-shadow:0 2px 8px rgba(0,0,0,.08);margin-bottom:20px}' +
    'h2{font-size:15px;color:#28534E;margin-bottom:16px;padding-bottom:8px;border-bottom:1px solid #e0ebe9}' +
    'label{display:block;font-size:12px;font-weight:600;color:#444;margin-bottom:4px;margin-top:12px}' +
    'input,select{width:100%;padding:9px 12px;border:1px solid #cdd9d7;border-radius:6px;font-size:14px}' +
    'input:focus,select:focus{outline:none;border-color:#28534E}' +
    '.btn{display:inline-block;padding:10px 20px;border:none;border-radius:6px;font-size:14px;font-weight:600;cursor:pointer;margin-top:16px}' +
    '.btn-primary{background:#28534E;color:#fff}' +
    '.btn-primary:hover{background:#1e3e3a}' +
    '.btn-danger{background:#dc3545;color:#fff;padding:5px 12px;font-size:12px;margin:0}' +
    '.msg{padding:10px 14px;border-radius:6px;margin-top:12px;font-size:13px;display:none}' +
    '.msg.ok{background:#d4edda;color:#155724}' +
    '.msg.er{background:#f8d7da;color:#721c24}' +
    'table{width:100%;border-collapse:collapse;font-size:13px}' +
    'th{background:#28534E;color:#fff;padding:8px 12px;text-align:left}' +
    'td{padding:8px 12px;border-bottom:1px solid #eee}' +
    'tr:hover td{background:#f8faf9}' +
    '.badge{display:inline-block;padding:2px 8px;border-radius:10px;font-size:11px;font-weight:700}' +
    '.badge-admin{background:#1a3c38;color:#fff}' +
    '.badge-editor{background:#d4edda;color:#155724}' +
    '.badge-viewer{background:#e2e3e5;color:#383d41}' +
    '</style></head><body>' +
    '<h1>&#128100; Gestor de Usuarios — VIDUSA</h1>' +
    '<div class="card"><h2>Agregar Usuario</h2>' +
    '<label>Username</label><input id="u" placeholder="usuario123">' +
    '<label>Contraseña</label><input id="p" type="password" placeholder="contraseña">' +
    '<label>Nombre Completo</label><input id="n" placeholder="Nombre Apellido">' +
    '<label>Rol</label><select id="r"><option value="">-- Seleccionar --</option><option>Admin</option><option>Editor</option><option>Viewer</option></select>' +
    '<label>Fraccionamiento (opcional)</label><input id="f" placeholder="Ej: Almendros">' +
    '<br><button class="btn btn-primary" onclick="addUser()">&#43; Guardar Usuario</button>' +
    '<div id="msg" class="msg"></div></div>' +
    '<div class="card"><h2>Usuarios Actuales <button class="btn btn-primary" style="padding:5px 12px;font-size:12px;float:right" onclick="loadUsers()">&#8635; Actualizar</button></h2>' +
    '<table><thead><tr><th>Username</th><th>Nombre</th><th>Rol</th><th>Fracc.</th><th></th></tr></thead>' +
    '<tbody id="tbl"><tr><td colspan="5" style="text-align:center;color:#888;padding:20px">Cargando...</td></tr></tbody></table></div>' +
    '<script>' +
    'var URL=window.location.href;' +
    'function msg(txt,type){var el=document.getElementById("msg");el.textContent=txt;el.className="msg "+type;el.style.display="block";setTimeout(function(){el.style.display="none"},4000)}' +
    'function addUser(){' +
    '  var u=document.getElementById("u").value.trim();' +
    '  var p=document.getElementById("p").value.trim();' +
    '  var n=document.getElementById("n").value.trim();' +
    '  var r=document.getElementById("r").value;' +
    '  var f=document.getElementById("f").value.trim();' +
    '  if(!u||!p||!n||!r){msg("Completa todos los campos requeridos","er");return}' +
    '  fetch(URL,{method:"POST",headers:{"Content-Type":"text/plain"},body:JSON.stringify({action:"ADD",username:u,password:p,nombre:n,rol:r,fraccionamiento:f})})' +
    '  .then(function(r){return r.json()})' +
    '  .then(function(d){' +
    '    if(d.status==="ok"){msg(d.message,"ok");document.getElementById("u").value="";document.getElementById("p").value="";document.getElementById("n").value="";document.getElementById("r").value="";document.getElementById("f").value="";loadUsers();}' +
    '    else msg(d.message,"er");' +
    '  }).catch(function(){msg("Error de conexión","er")})' +
    '}' +
    'function loadUsers(){' +
    '  fetch(URL,{method:"POST",headers:{"Content-Type":"text/plain"},body:JSON.stringify({action:"LIST"})})' +
    '  .then(function(r){return r.json()})' +
    '  .then(function(d){' +
    '    var tb=document.getElementById("tbl");' +
    '    if(!d.users||!d.users.length){tb.innerHTML="<tr><td colspan=\'5\' style=\'text-align:center;color:#888;padding:20px\'>Sin usuarios</td></tr>";return}' +
    '    tb.innerHTML=d.users.map(function(u){' +
    '      var bc=u.rol.toLowerCase()==="admin"?"admin":u.rol.toLowerCase()==="editor"?"editor":"viewer";' +
    '      return "<tr><td>"+u.username+"</td><td>"+u.nombre+"</td><td><span class=\'badge badge-"+bc+"\'>"+(u.rol||"")+"</span></td><td>"+(u.fracc||"—")+"</td><td><button class=\'btn btn-danger\' onclick=\'delUser(\\""+u.username+"\\")\'>&#128465;</button></td></tr>";' +
    '    }).join("")' +
    '  }).catch(function(){document.getElementById("tbl").innerHTML="<tr><td colspan=\'5\'>Error al cargar</td></tr>"})' +
    '}' +
    'function delUser(u){' +
    '  if(!confirm("¿Eliminar usuario "+u+"?"))return;' +
    '  fetch(URL,{method:"POST",headers:{"Content-Type":"text/plain"},body:JSON.stringify({action:"DELETE",username:u})})' +
    '  .then(function(r){return r.json()})' +
    '  .then(function(d){msg(d.message,d.status==="ok"?"ok":"er");if(d.status==="ok")loadUsers()})' +
    '  .catch(function(){msg("Error de conexión","er")})' +
    '}' +
    'loadUsers();' +
    '<\/script></body></html>';
}
