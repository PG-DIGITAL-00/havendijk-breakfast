// ─── HAVENDIJK — Dashboard reader ───────────────────────────
// Archivo: dashboard.gs dentro del mismo proyecto Apps Script
// Este script devuelve los datos de la Sheet al dashboard de cocina

function doGet(e) {
  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    var lastRow = sheet.getLastRow();

    // Si solo hay encabezados o está vacía, devuelve array vacío
    if (lastRow <= 1) {
      return ContentService
        .createTextOutput(JSON.stringify({ result: 'ok', reservations: [] }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    // Lee todas las filas de datos (desde fila 2, saltando encabezados)
    var data = sheet.getRange(2, 1, lastRow - 1, 11).getValues();

    var reservations = data.map(function(row, index) {
      return {
        id:          index + 1,
        fecha:       row[0],
        horaRegistro: row[1],
        habitacion:  String(row[2]),
        adultos:     row[3],
        ninos:       row[4],
        edadNinos:   row[5],
        tipo:        row[6],
        plato:       row[7],
        horaLlegada: row[8],
        dietetico:   row[9],
        idioma:      row[10],
        status:      'pending'   // el estado se maneja localmente en el dashboard
      };
    });

    // Ordena por hora de llegada
    reservations.sort(function(a, b) {
      return a.horaLlegada.localeCompare(b.horaLlegada);
    });

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'ok', reservations: reservations }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
