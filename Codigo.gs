
//////////////////////////////////////////////////////////
/////////// ACTIVAR ESTAS ACCIONES ///////////////////////
//////////////////////////////////////////////////////////

function testing() {
  var sheetHoja = "testing";
  verAsistenciaMeetTEST(sheetHoja);
}

function enviarEmail(cuerpo) {
  var nowTime = new Date();
  var timeFormateado = Utilities.formatDate(nowTime, 'Europe/Madrid', 'MM/dd/yyyy\'T\'HH:mm:ss.SSS');
  GmailApp.sendEmail('jp@liceobritanico.com', 'REPORTE ASISTENCIAS LCB- ' + cuerpo, cuerpo + ' ' + timeFormateado);
};

//////////////////////////////////////////////////////////
//////////// FIN DE LOS TRIGGERS  ////////////////////////
//////////////////////////////////////////////////////////


function verAsistenciaMeetTEST(hoja) {
  try {
    enviarEmail("Arranco proceso " + hoja);
    var start = new Date();
    var startFormateado = Utilities.formatDate(start, 'Europe/Madrid', 'MM/dd/yyyy\'T\'HH:mm:ss.SSS');
    Logger.log("ARRANCO EL SCRIPT: " + start);
    Logger.log("StartFormateado: " + startFormateado);

    var ss = SpreadsheetApp.openByUrl(docURL).getSheetByName(hoja);  // for StandAlone Script
    //var ss=SpreadsheetApp.getActive().getSheetByName(hoja); // for ContainerBound Script
    var colC = ss.getRange("C3:C").getValues(); // Selecciona TODA la ColC
    var filasLlenasLengthC = colC.filter(String).length; // Cuenta cuantos datos tiene ColC
    Logger.log("Col C - ClassID= " + filasLlenasLengthC);

    var colD = ss.getRange("D3:D").getValues(); // Selecciona TODA la ColD
    var filasLlenasLengthD = colD.filter(String).length; // Cuenta cuantos datos tiene ColD
    Logger.log("Col D - Usuarios= " + filasLlenasLengthD);

    var colMeet = ss.getRange("E3:E").getValues(); // Selecciona TODA la ColE
    var filasLlenasLengthMeet = colMeet.filter(String).length; // Cuenta cuantos datos tiene ColE
    Logger.log("Col E - Meet= " + filasLlenasLengthMeet);

    var colAsistencia = ss.getRange("G3:G").getValues(); // Selecciona TODA la ColG
    var filasLlenasLengthAsistencia = colAsistencia.filter(String).length; // Cuenta cuantos datos tiene ColG
    Logger.log("Col G - Asistencia= " + filasLlenasLengthAsistencia);

    if (filasLlenasLengthAsistencia < filasLlenasLengthD) {
      Logger.log("MARCADOR: ARRANCAR PROCESO VERIFICACION");
      var dia = ss.getRange("G1").getValue(); // Fecha escrita en la celda G1
      var fecha = fechaMasUno(dia, 1.5); // Transforma la fecha para que muestre el horario, etc
      // Logger.log("Fecha en Spreadsheet: "+fecha);

      for (var i = filasLlenasLengthAsistencia; i < filasLlenasLengthD; i++) {
        // Logger.log ("Full Meet Code: "+colMeet[i]);
        var meetCode = colMeet[i];

        //////// COMIENZO = Funcion para evitar que el Script "runs out of time".  
        if (isTimeUp_(start)) {
          Logger.log("Time up");
          Logger.log("Create Trigger for a few mins later");

          enviarEmail('Termino Ciclo de ParcialMeet. Registros hechos: ' + filasLlenasLengthAsistencia + '/' + filasLlenasLengthD);
          // setUpTemporaryTrigger('hojaMeetParcial1', 3); ///////////// Crea Trigger para repetir en 3 minutos.
          setUpTemporaryTrigger();
          break;
          Logger.log("MARCADOR: BREAK"); /// DELETE AFTER ///////////////////////
        }
        //////// FIN = Funcion para evitar que el Script "runs out of time".  
        checkMeetCodeTest(meetCode, i, fecha, hoja); // Runs this function
      }
      Logger.log("MARCADOR: AFTER FOR LOOP"); /// DELETE AFTER ///////////////////////

      // RUNS ALWAYS
      // Logger.log("FIN DEL PROCESO. Deberia estar completo");
      var colAsistenciaValues = ss.getRange("G3:G").getValues(); // Selecciona TODA la ColA
      var colAsistenciaCheck = colAsistenciaValues.filter(String).length; // Cuenta cuantos datos tiene ColA
      Logger.log("COL ASISTENCIAS CHECK = " + colAsistenciaCheck);
      Logger.log("FilasLlenasLengthD-Usuarios= " + filasLlenasLengthD);

      if (colAsistenciaCheck + 1 == filasLlenasLengthD || colAsistenciaCheck == filasLlenasLengthD || colAsistenciaCheck - 1 == filasLlenasLengthD) {
        Logger.log("MARCADOR: FIN DEL PROCESO. CORRE: HACER_TODO... Deberia estar completo");
        HACER_TODO();  // Runs this function
      }
    }
    else {
      Logger.log("MARCADOR: NO ARRANCAR PROCESO VERIFICACION")
      // ERROR
      // HACER_TODO(); ACA NO VA...
    }
    Logger.log("MARCADOR: FINAL DE LA FUNCION")

  } catch (e) {
    Logger.log("Error logged: " + e)
  }

}


// Funcion para evitar que el Script "runs out of time". 
// Source: https://www.labnol.org/code/20016-maximum-execution-time-limit
function isTimeUp_(startTime) {
  var now = new Date();
  var nowFormateado = Utilities.formatDate(now, 'Europe/Madrid', 'MM/dd/yyyy\'T\'HH:mm:ss.SSS');
  return now.getTime() - startTime.getTime() > 1740000; // 29.0 minutes
  Logger.log("NOW Time: " + now);
  Logger.log("NOWFORMATEADO: " + nowFormateado);
}


// FUNCION PARA VERIFICAR SI ESTUVO PRESENTE
function checkMeetCodeTest(meetCode, index, searchDate, hoja) {
  try {
    var ss = SpreadsheetApp.getActive().getSheetByName(hoja);
    var colB = ss.getRange("D3:D").getValues(); // Selecciona TODA la ColA
    var filasLlenasLengthB = colB.filter(String).length; // Cuenta cuantos datos tiene ColA

    var durationWeb = 0;
    var durationMovil = 0;
    var parameters = [
      'duration_seconds',
      'device_type',
      'identifier',
    ];

    // universal settings - static
    var userKey = 'all';
    var applicationName = 'meet';
    var emailAddress = colB[index];
    Logger.log("Index / Alumno: " + index + " / " + emailAddress);

    var optionalArgs = {
      event_name: "call_ended",
      endTime: searchDate,
      startTime: fechaRestada(searchDate),
      filters: "identifier==" + emailAddress + ",meeting_code==" + meetCode
    };
    const response = AdminReports.Activities.list(userKey, applicationName, optionalArgs)
    Logger.log("RESPONSE: " + response);  // -----------------------------  RETURNS LONG JSON ----------------------------------------------
    var myItems = response.items;
    Logger.log("MyITEMS ALL = " + myItems);

    if (myItems == null) {            // Si no encuentra "Object property" ITEMS mark ABSENT 
      printAbsent(ss, index + 3, 7)
      Logger.log("Mark as Absent - No 'Items present'")
    }
    else {                            // Si encuentra "Object property" ITEMS mark PRESENT                 
      for (i = 0; i < myItems.length; i++) {
        var insideItemsI = myItems[i]; // Returns OBJECT
        Logger.log("Object insideItems[i]= " + i + "/" + insideItemsI);     // Logs every ITEMS found 
        var theEvent = insideItemsI.events; // Entra en el Object Property EVENTS
        var parametros = theEvent[0].parameters // Entra en el primer Array dentro de EVENTS y Propiedad PARAMETERS dentro de ese array.

        if (parametros); {
          Logger.log("Parametros = " + parametros) // Logs every PARAMETER found with all its object properties.
          var parameterValues = getParameterValues(parametros);
          var duration = parameterValues['duration_seconds'];
          var device = parameterValues['device_type'];
          duration = +duration; // convierte el STRING que devuelce duration a un INTEGER (numero)
          Logger.log("DURATION_S = " + duration);
          Logger.log("DEVICE = " + device);

          if (device === "web") {
            durationWeb = durationWeb + duration;
          } else {
            durationMovil = durationMovil + duration;
          }
          Logger.log("DURATION_WEB Parcial " + i + " = " + durationWeb)
          Logger.log("DURATION_MOVIL Parcial " + i + " = " + durationMovil)
        }
      }

      Logger.log("DURATION WEB TOTAL x USUARIO = " + durationWeb);
      Logger.log("DURATION MOVIL TOTAL x USUARIO = " + durationMovil)
      Logger.log("Escribiendo Resultados en Sheets");

      printPresent(ss, index + 3, 7);
      if (durationWeb > 0) {
        printWebDuration(durationWeb, ss, index + 3, 8);
      }
      if (durationMovil > 0) {
        printMovilDuration(durationMovil, ss, index + 3, 9);
      }

    }

  } catch (e) {
    Logger.log("Error logged: " + e)
  }
  // Fin de la funcion        
}


// * Gets a map of parameter names to values from an array of parameter objects.
// * @param {Array} parameters An array of parameter objects.
// * @return {Object} A map from parameter names to their values.

function getParameterValues(parameters) {
  try {
    return parameters.reduce(function (result, parameter) {
      var name = parameter.name;
      var value;
      if (parameter.intValue !== undefined) {
        value = parameter.intValue;
      } else if (parameter.value !== undefined) {
        value = parameter.value;
      } else if (parameter.stringValue !== undefined) {
        value = parameter.stringValue;
      } else if (parameter.datetimeValue !== undefined) {
        value = new Date(parameter.datetimeValue);
      } else if (parameter.boolValue !== undefined) {
        value = parameter.boolValue;
      }
      result[name] = value;
      //Logger.log("Result getParameterValues ="+result); // Returns [object Object]
      return result;
    }, {});
  } catch (e) {
    Logger.log("Error logged: " + e)
  }
}

function printAbsent(sheet, r, c) {
  var cell = sheet.getRange(r, c);
  cell.setValue("Absent");
}

function printPresent(sheet, r, c) {
  var cell = sheet.getRange(r, c);
  cell.setValue("PRESENT");
}

function printWebDuration(print, sheet, r, c) {
  var cell = sheet.getRange(r, c);
  cell.setValue(print / 60);
}

function printMovilDuration(print, sheet, r, c) {
  var cell = sheet.getRange(r, c);
  cell.setValue(print / 60);
}
