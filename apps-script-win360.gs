/**
 * WIN-360 · Puente entre la web y la hoja de cálculo
 * ---------------------------------------------------
 * 1. Crea una hoja de cálculo en Google Sheets.
 * 2. Renombra la primera pestaña como  Base
 * 3. Extensiones › Apps Script › pega este código.
 * 4. Implementar › Nueva implementación › Aplicación web
 *       Ejecutar como: Yo
 *       Quién tiene acceso: Cualquier usuario
 * 5. Copia la URL que termina en /exec y pégala en win360.html:
 *       const WEBHOOK_URL = "https://script.google.com/.../exec";
 */

const HOJA = 'Base';

const COLUMNAS = ['ID','Fecha','Nombre','Ubicación','Número de contacto','Pisos vivienda',
  'Piso router','Zona con problema','Piso del problema','¿Qué sucede?','¿Qué hace ahí?',
  'Mejora cerca del router','Paredes','¿Qué intentó?','Índice','Resultado','Recomendación',
  'Precio (S/)','Oferta','Estado','Técnico asignado'];

function hoja_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(HOJA) || ss.insertSheet(HOJA);
  if (sh.getLastRow() === 0) {
    sh.appendRow(COLUMNAS);
    sh.getRange(1, 1, 1, COLUMNAS.length).setFontWeight('bold')
      .setBackground('#1B1B3F').setFontColor('#FFFFFF');
    sh.setFrozenRows(1);
  }
  return sh;
}

/** La web envía cada diagnóstico aquí. */
function doPost(e) {
  const sh = hoja_();
  const d = JSON.parse(e.postData.contents);
  const id = sh.getLastRow(); // fila 1 es cabecera, así el ID queda correlativo
  sh.appendRow([
    id, d.fecha, d.nombre, d.ubicacion, d.contacto, d.pisos, d.pisoRouter, d.zona,
    d.pisoProblema, d.sucede, d.hace, d.mejora, d.paredes, d.intento,
    d.indice, d.resultado, d.recomendacion, d.precio, d.oferta || '—',
    d.estado || 'Pendiente', d.tecnico || ''
  ]);
  return ContentService.createTextOutput(JSON.stringify({ ok: true, id: id }))
    .setMimeType(ContentService.MimeType.JSON);
}

/** El panel "Central WIN" lee la hoja desde aquí. */
function doGet(e) {
  const sh = hoja_();
  const valores = sh.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < valores.length; i++) {
    const v = valores[i];
    if (!v[2]) continue;
    rows.push({
      id: v[0], fecha: formatearFecha_(v[1]), nombre: v[2], ubicacion: v[3], contacto: v[4],
      pisos: v[5], pisoRouter: v[6], zona: v[7], pisoProblema: v[8], sucede: v[9],
      hace: v[10], mejora: v[11], paredes: v[12], intento: v[13], indice: v[14],
      resultado: v[15], recomendacion: v[16], precio: v[17], oferta: v[18],
      estado: v[19], tecnico: v[20],
      mesh: String(v[16] || '').indexOf('2') === 0 ? 2 : (String(v[16] || '').indexOf('Mesh') > -1 ? 1 : 0)
    });
  }
  return ContentService.createTextOutput(JSON.stringify({ rows: rows }))
    .setMimeType(ContentService.MimeType.JSON);
}

function formatearFecha_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  return String(v || '');
}
