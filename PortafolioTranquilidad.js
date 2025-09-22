function revisarRegistros() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var hoja = ss.getSheetByName("Leadform");
  if (!hoja) {
    Logger.log("Hoja 'Copia de Leadform' no encontrada.");
    return;
  }

  var lastRow = hoja.getLastRow();
  if (lastRow < 2) {
    Logger.log("No hay filas con datos (lastRow < 2).");
    return;
  }

  var headers = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
  var candidates = ['fecha','timestamp','date','created_at','created at','hora','createdat'];
  var fechaCol = -1;
  for (var i = 0; i < headers.length; i++) {
    var h = (headers[i] || '').toString().toLowerCase();
    for (var j = 0; j < candidates.length; j++) {
      if (h.indexOf(candidates[j]) !== -1) { fechaCol = i + 1; break; }
    }
    if (fechaCol !== -1) break;
  }
  if (fechaCol === -1) fechaCol = 8; 

  var datos = hoja.getRange(2, fechaCol, lastRow - 1, 1).getValues();

  var tz = ss.getSpreadsheetTimeZone();
  var hoy = new Date();
  var ayer = new Date(hoy);
  ayer.setDate(hoy.getDate() - 1);
  var anioA = ayer.getFullYear(), mesA = ayer.getMonth(), diaA = ayer.getDate();

  var huboRegistros = false;
  var filasEncontradas = [];
  var filasNoParseables = [];

  for (var r = 0; r < datos.length; r++) {
    var cell = datos[r][0];
    var fecha = null;

    if (cell instanceof Date && !isNaN(cell.getTime())) {
      fecha = cell;
    } else if (typeof cell === 'string' && cell.trim() !== '') {
      var s = cell.trim().replace(/\//g, '-');
      var iso = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})(?:[ T](\d{1,2}):(\d{2})(?::(\d{2}))?(?:\.(\d+))?(Z|[+\-]\d{2}:?\d{2})?)?$/);
      if (iso) {
        var y = parseInt(iso[1], 10);
        var m = parseInt(iso[2], 10) - 1;
        var d = parseInt(iso[3], 10);
        var hh = parseInt(iso[4] || 0, 10);
        var mm = parseInt(iso[5] || 0, 10);
        var ssSec = parseInt(iso[6] || 0, 10);
        var tzpart = iso[8];

        if (tzpart === 'Z') {
          fecha = new Date(Date.UTC(y, m, d, hh, mm, ssSec));
        } else if (tzpart && /^[+\-]\d{2}:?\d{2}$/.test(tzpart)) {
          var off = tzpart.replace(':', '');
          var sign = off[0] === '+' ? 1 : -1;
          var oh = parseInt(off.slice(1, 3), 10);
          var om = parseInt(off.slice(3), 10);
          var utc = Date.UTC(y, m, d, hh, mm, ssSec);
          fecha = new Date(utc - sign * (oh * 3600000 + om * 60000));
        } else {
          fecha = new Date(y, m, d, hh, mm, ssSec);
        }
      } else {
        var parsed = Date.parse(s);
        if (!isNaN(parsed)) fecha = new Date(parsed);
      }
    }

    if (fecha && !isNaN(fecha.getTime())) {
      if (fecha.getFullYear() === anioA && fecha.getMonth() === mesA && fecha.getDate() === diaA) {
        huboRegistros = true;
        filasEncontradas.push({ row: r + 2, fecha: Utilities.formatDate(fecha, tz, "yyyy-MM-dd HH:mm:ss") });
      }
    } else {
      filasNoParseables.push({ row: r + 2, raw: cell });
    }
  }

  if (filasEncontradas.length) Logger.log("Filas encontradas para " + Utilities.formatDate(ayer, tz, "yyyy-MM-dd") + ": " + JSON.stringify(filasEncontradas));
  if (filasNoParseables.length) Logger.log("Filas no parseables (revisar formato): " + JSON.stringify(filasNoParseables));

  if (!huboRegistros) {
    var destinatariosJefes = [
      "anamilena.roa@segurosbolivar.com",
      "diana.cordoba@segurosbolivar.com",
      "nohora.jaimes@segurosbolivar.com",
      "sebastian.daza@segurosbolivar.com",
      "maria.camila.rodriguez@segurosbolivar.com"
    ];
    var urlHoja = ss.getUrl();

    var asuntoJefes = "Alerta: No hubo registros para " + hoja.getName() + " el " + Utilities.formatDate(ayer, tz, "yyyy-MM-dd");
    var mensajeJefes = "No se añadieron registros en la hoja '" + hoja.getName() + "' el " +
                       Utilities.formatDate(ayer, tz, "yyyy-MM-dd") + ".\n\n" +
                       "Puedes revisar la hoja aquí: " + urlHoja;

    MailApp.sendEmail(destinatariosJefes.join(","), asuntoJefes, mensajeJefes);

    var asuntoRebe = "[COPIA] " + asuntoJefes;
    var mensajeRebe = mensajeJefes + "\n\n" +
                      (filasNoParseables.length ? "Atención: hay filas cuyo formato de fecha no se pudo parsear:\n" + JSON.stringify(filasNoParseables) : "");
    MailApp.sendEmail("rebeca.pedrozo@segurosbolivar.com", asuntoRebe, mensajeRebe);

  } else {
    var urlHoja = ss.getUrl();
    var asuntoRebe = "Reporte: Sí hubo registros en PORTAFOLIO TRANQUILIDAD " + hoja.getName() + " el " + Utilities.formatDate(ayer, tz, "yyyy-MM-dd");
    var mensajeRebe = "Se detectaron registros en la hoja '" + hoja.getName() + "' el " + Utilities.formatDate(ayer, tz, "yyyy-MM-dd") + ".\n\n" +
                      "Cantidad de filas detectadas: " + filasEncontradas.length + "\n\n" +
                      "Ejemplos:\n" + JSON.stringify(filasEncontradas.slice(0,5)) + "\n\n" +
                      (filasNoParseables.length ? "También hubo filas no parseables:\n" + JSON.stringify(filasNoParseables) + "\n\n" : "") +
                      "Puedes revisar la hoja aquí: " + urlHoja;

    MailApp.sendEmail("rebeca.pedrozo@segurosbolivar.com", asuntoRebe, mensajeRebe);
  }
}