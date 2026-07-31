// ─────────────────────────────────────────────────────────
// PowerBI.gs — Despivotado de BD_8D para Power BI
// No modifica BD_8D, USUARIOS, ni el Web App (doGet/doPost).
// Solo lee BD_8D y escribe/actualiza BD_8D_Despivotada.
// ─────────────────────────────────────────────────────────

function despivotarBD8D() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaOrigen = ss.getSheetByName('BD_8D');
  const nombreHojaDestino = 'BD_8D_Despivotada';

  const datos = hojaOrigen.getDataRange().getValues();
  const encabezados = datos[0];
  const filas = datos.slice(1);

  // Columnas identificadoras que se repiten en cada fila de salida
  const columnasId = ['ID_Registro', 'Folio', 'Fraccionamiento', 'Fecha_Revision',
                       'Superintendente', 'Residente', 'Facilitador_BPO',
                       'Descripcion_Problema', 'Causa_Raiz', 'Ponderacion_Causa'];
  const idxId = columnasId.map(nombre => encabezados.indexOf(nombre));

  // Detecta dinámicamente cuántos bloques Accion_Correctiva_N existen
  const numerosAccion = [];
  encabezados.forEach(h => {
    const match = String(h).match(/^Accion_Correctiva_(\d+)$/);
    if (match) numerosAccion.push(Number(match[1]));
  });
  numerosAccion.sort((a, b) => a - b);

  const bloques = numerosAccion.map(n => ({
    num: n,
    idxAccion: encabezados.indexOf('Accion_Correctiva_' + n),
    idxResponsable: encabezados.indexOf('Responsable_Accion_' + n),
    idxFechaProg: encabezados.indexOf('Fecha_Programada_' + n),
    idxFechaReal: encabezados.indexOf('Fecha_Realizacion_' + n),
    idxBitacora: encabezados.indexOf('Bitacora_Accion_' + n),
    idxFechaBitacora: encabezados.indexOf('Fecha_Bitacora_' + n)
  }));

  const encabezadoSalida = [...columnasId, 'Num_Accion', 'Accion_Correctiva',
                             'Responsable_Accion', 'Fecha_Programada', 'Fecha_Realizacion',
                             'Bitacora_Accion', 'Fecha_Bitacora'];
  const salida = [encabezadoSalida];

  filas.forEach(fila => {
    if (fila.every(c => c === '' || c === null)) return;

    const valoresId = idxId.map(idx => fila[idx]);

    bloques.forEach(b => {
      const accion = fila[b.idxAccion];
      if (accion !== '' && accion !== null && accion !== undefined) {
        salida.push([
          ...valoresId,
          b.num,
          accion,
          fila[b.idxResponsable],
          fila[b.idxFechaProg],
          fila[b.idxFechaReal],
          b.idxBitacora >= 0 ? fila[b.idxBitacora] : '',
          b.idxFechaBitacora >= 0 ? fila[b.idxFechaBitacora] : ''
        ]);
      }
    });
  });

  let hojaDestino = ss.getSheetByName(nombreHojaDestino);
  if (!hojaDestino) {
    hojaDestino = ss.insertSheet(nombreHojaDestino);
  } else {
    hojaDestino.clearContents();
  }

  hojaDestino.getRange(1, 1, salida.length, salida[0].length).setValues(salida);

  if (salida.length > 1) {
    const colFechaProg = encabezadoSalida.indexOf('Fecha_Programada') + 1;
    const colFechaReal = encabezadoSalida.indexOf('Fecha_Realizacion') + 1;
    const colFechaBitacora = encabezadoSalida.indexOf('Fecha_Bitacora') + 1;
    hojaDestino.getRange(2, colFechaProg, salida.length - 1, 1).setNumberFormat('dd/mm/yyyy');
    hojaDestino.getRange(2, colFechaReal, salida.length - 1, 1).setNumberFormat('dd/mm/yyyy');
    hojaDestino.getRange(2, colFechaBitacora, salida.length - 1, 1).setNumberFormat('dd/mm/yyyy');
  }

  hojaDestino.setFrozenRows(1);
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('8D Herramientas')
    .addItem('Despivotar acciones correctivas', 'despivotarBD8D')
    .addToUi();
}
