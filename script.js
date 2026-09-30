function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(30000);

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Encabezados si está vacía
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(["Fecha y Hora", "Nombre", "Edad", "rfc", "curp", "estudios", "correo", "direccion"]);
    }

    var nombre = (e && e.parameter && e.parameter.nombre) ? e.parameter.nombre : "";
    var edad = (e && e.parameter && e.parameter.edad) ? e.parameter.edad : "";
    var rfc = (e && e.parameter && e.parameter.rfc) ? e.parameter.rfc : "";
    var curp = (e && e.parameter && e.parameter.curp) ? e.parameter.curp : "";
    var estudios = (e && e.parameter && e.parameter.estudios) ? e.parameter.estudios : "";
    var correo = (e && e.parameter && e.parameter.correo) ? e.parameter.correo : "";
    var direccion = (e && e.parameter && e.parameter.direccion) ? e.parameter.direccion : "";


    sheet.appendRow([new Date(), nombre, edad, rfc, curp, estudios, correo, direccion]);

    return ContentService
      .createTextOutput("OK")
      .setMimeType(ContentService.MimeType.TEXT);

  } catch (error) {
    return ContentService
      .createTextOutput("Error: " + error.toString())
      .setMimeType(ContentService.MimeType.TEXT);

  } finally {
    lock.releaseLock();
  }
}