/**
 * BACKEND para "Mis Gastos - Personal" (Google Apps Script)
 * ---------------------------------------------------------
 * Recibe cada gasto que envía la app y lo agrega como una fila nueva
 * en la pestaña "Gastos" de tu hoja "Timeless - Ventas e Inventario".
 *
 * CÓMO USARLO (resumen — los pasos detallados están en el chat):
 *  1. Abre tu hoja "Timeless - Ventas e Inventario" en Google Sheets.
 *  2. Menú: Extensiones -> Apps Script.
 *  3. Borra el código de ejemplo y pega TODO este archivo.
 *  4. Guarda (icono de disquete).
 *  5. Implementar -> Nueva implementación -> tipo "Aplicación web".
 *     - Ejecutar como: Yo
 *     - Quién tiene acceso: Cualquier usuario
 *  6. Implementar -> Autoriza los permisos que te pida.
 *  7. Copia la "URL de la aplicación web" (termina en /exec).
 *  8. Pega esa URL en app.js (constante SHEETS_WEBHOOK_URL).
 *
 * Como el script está creado DESDE la hoja, usa getActiveSpreadsheet()
 * y no necesita ningún ID.
 */

var SHEET_NAME = 'Gastos';

// Recibe cada gasto (POST). Si la ID ya existe en la hoja, ACTUALIZA esa fila
// (así llegan las ediciones hechas en la app, no solo el gasto original);
// si no existe, la agrega como fila nueva.
function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000); // evita filas duplicadas si llegan dos a la vez

    var data = JSON.parse(e.postData.contents);
    var sheet = getOrCreateSheet_();

    // Encabezados la primera vez.
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['ID', 'Fecha', 'Categoría', 'Monto', 'Nota', 'Registrado en']);
    }

    var fecha = data.date ? new Date(data.date) : new Date();
    var fila = data.id ? buscarFila_(sheet, data.id) : -1;

    if (fila > 0) {
      // Ya existe esa ID: es una edición, actualiza la fila en vez de duplicarla.
      sheet.getRange(fila, 2, 1, 4).setValues([[
        fecha,
        data.category || '',
        Number(data.amount) || 0,
        data.note || ''
      ]]);
      return json_({ ok: true, updated: true });
    }

    sheet.appendRow([
      data.id || '',
      fecha,
      data.category || '',
      Number(data.amount) || 0,
      data.note || '',
      new Date()
    ]);

    return json_({ ok: true });
  } catch (err) {
    return json_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (ignore) {}
  }
}

// Permite abrir la URL en el navegador para comprobar que está viva.
function doGet() {
  return json_({ ok: true, service: 'Mis Gastos - Personal', sheet: SHEET_NAME });
}

function getOrCreateSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  return sheet;
}

// Devuelve el número de fila (1-based) donde está esa ID, o -1 si no existe.
function buscarFila_(sheet, id) {
  var last = sheet.getLastRow();
  if (last < 2) return -1;
  var ids = sheet.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2; // +2: la fila 1 es encabezado
  }
  return -1;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
