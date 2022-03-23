 // Menu Options 
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('<< HOAKEEN.com >>')
      .addItem('Activar Parcial1 Manualmente', 'hojaMeetParcial1')
      .addItem('Testing', 'testing')
      .addSeparator()
      .addToUi();
}


///////////////////////////////////////////////////////////
/////// GLOBAL VARABLES ///////////////////////////////////
var docURL = "https://docs.google.com/spreadsheets/d/1WliOklNCJQi_pDsRSm7fX_GK1PtHUuk39gw8ZlPyqgI/edit#gid=0";
var newDocID;
var fechaFormat;
var driveFolderID = "1WnWp3GLPBdqcUdfVtjJNSaLFYZ_KGH8U"; // Carpeta "ASISTENCIA A MEET [NOT SHARED]"
var CopyXSedeDriveFolderID = '1oAGW3rWaQuWjK87th0oc6BHxB9wTbMiU' // Carpeta "xSede dentro de [NOT SHARED]"
// var driveFolderID = "1hQTQWhxm6IVKO_AcZiaABhH3NO2EealE"; // Carpeta "ASISTENCIA A MEET"
// var CopyXSedeDriveFolderID = '1yBvxkPz5e50gUND_ad7BIwQv6mYzxL93' // Subcarpeta "\ASISTENCIA X SEDE"


/////////////////////////////////////////////////////////////
/////////// ACTIVAR ESTAS ACCIONES //////////////////////////
/////////////////////////////////////////////////////////////

function hojaMeetParcial1(){
  var sheetHoja = "MeetParcial1";
  verAsistenciaMeet(sheetHoja);
  }

function HACER_TODO(){
  Logger.log("Runs Informe_Duplicar()");
  Informe_Duplicar();
  
  Logger.log("Runs Informe_Copiar_Values()");
  Informe_Copiar_Values();
    
  Logger.log("Runs Informe_Modificar() &  MakeCopySede() at the end");
  Informe_Modificar();
  
  Logger.log("Runs Informe_Duplicar()");
  LimpiarHojasParciales();
  
  enviarEmail('Termino todo el Script');
}

function enviarEmail (cuerpo){
  var nowTime = new Date();
  var timeFormateado = Utilities.formatDate(nowTime, 'Europe/Madrid', 'MM/dd/yyyy\'T\'HH:mm:ss.SSS');
  GmailApp.sendEmail('jp@liceobritanico.com', 'REPORTE ASISTENCIAS LCB- '+cuerpo,cuerpo +' '+ timeFormateado);
};

//////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////
//////////////////////////////////////////////////////////

function verAsistenciaMeet(hoja){
try{
  enviarEmail("Arranco proceso "+hoja);
  var start = new Date();
  var startFormateado = Utilities.formatDate(start, 'Europe/Madrid', 'MM/dd/yyyy\'T\'HH:mm:ss.SSS');
  Logger.log("ARRANCO EL SCRIPT: "+start);
  Logger.log("StartFormateado: "+startFormateado);

    var ss=SpreadsheetApp.openByUrl(docURL).getSheetByName(hoja);  // for StandAlone Script
    //var ss=SpreadsheetApp.getActive().getSheetByName(hoja); // for ContainerBound Script
    var colC = ss.getRange("C3:C").getValues(); // Selecciona TODA la ColC
    var filasLlenasLengthC = colC.filter(String).length; // Cuenta cuantos datos tiene ColC
    Logger.log("Col C - ClassID= "+filasLlenasLengthC);
   
    var colD = ss.getRange("D3:D").getValues(); // Selecciona TODA la ColD
    var filasLlenasLengthD = colD.filter(String).length; // Cuenta cuantos datos tiene ColD
    Logger.log("Col D - Usuarios= "+filasLlenasLengthD);
  
    var colMeet = ss.getRange("E3:E").getValues(); // Selecciona TODA la ColE
    var filasLlenasLengthMeet = colMeet.filter(String).length; // Cuenta cuantos datos tiene ColE
    Logger.log("Col E - Meet= "+filasLlenasLengthMeet);
  
    var colAsistencia = ss.getRange("G3:G").getValues(); // Selecciona TODA la ColG
    var filasLlenasLengthAsistencia = colAsistencia.filter(String).length; // Cuenta cuantos datos tiene ColG
    Logger.log("Col G - Asistencia= "+filasLlenasLengthAsistencia);
  
  
  
  if(filasLlenasLengthAsistencia < filasLlenasLengthD){
   Logger.log("MARCADOR: ARRANCAR PROCESO VERIFICACION");  
 
    var dia = ss.getRange("G1").getValue(); // Fecha escrita en la celda G1
    var fecha = fechaMasUno(dia, 1.5); // Transforma la fecha para que muestre el horario, etc
   // Logger.log("Fecha en Spreadsheet: "+fecha);
  
  
    for (var i = filasLlenasLengthAsistencia; i < filasLlenasLengthD; i++){
     // Logger.log ("Full Meet Code: "+colMeet[i]);
      var meetCode = colMeet[i];
     
      //////// COMIENZO = Funcion para evitar que el Script "runs out of time".  
          if (isTimeUp_(start)) { 
            Logger.log("Time up");
            Logger.log("Create Trigger for a few mins later");
            
            enviarEmail('Termino Ciclo de ParcialMeet. Registros hechos: '+filasLlenasLengthAsistencia + '/'+ filasLlenasLengthD);
            // setUpTemporaryTrigger('hojaMeetParcial1', 3); ///////////// Crea Trigger para repetir en 3 minutos.
            setUpTemporaryTrigger();
            break;
          Logger.log("MARCADOR: BREAK"); /// DELETE AFTER ///////////////////////
          }
      //////// FIN = Funcion para evitar que el Script "runs out of time".  
          
          checkMeetCodeWebMovil(meetCode, i, fecha, hoja); // Runs this function
         
     }     
    Logger.log("MARCADOR: AFTER FOR LOOP"); /// DELETE AFTER ///////////////////////
    
    // RUNS ALWAYS
    // Logger.log("FIN DEL PROCESO. Deberia estar completo");
    var colAsistenciaValues = ss.getRange("G3:G").getValues(); // Selecciona TODA la ColA
    var colAsistenciaCheck = colAsistenciaValues.filter(String).length; // Cuenta cuantos datos tiene ColA
    Logger.log("COL ASISTENCIAS CHECK = " + colAsistenciaCheck);
    Logger.log("FilasLlenasLengthD-Usuarios= "+filasLlenasLengthD);

    // Margen de 5 filas de diferencia
    for (var w = 0; w <6; w++) {
    if(colAsistenciaCheck +w == filasLlenasLengthD ||  colAsistenciaCheck -w == filasLlenasLengthD ){
      Logger.log("MARCADOR: FIN DEL PROCESO 1. CORRE: HACER_TODO... ");
      setUpHacerTodoTrigger() // Runs the HACER_TODO function
      }
     }
    } 
     else {    
       Logger.log ("MARCADOR: NO ARRANCAR PROCESO VERIFICACION 2")   
       // ERROR
       // HACER_TODO(); ACA NO VA...
    }
     Logger.log ("MARCADOR: FINAL DE LA FUNCION 3")       
} catch (e){
  Logger.log("Error logged: "+e)
           }
} 
 

// Funcion para evitar que el Script "runs out of time". 
// Source: https://www.labnol.org/code/20016-maximum-execution-time-limit
function isTimeUp_(startTime) {
  var now = new Date();
  var nowFormateado = Utilities.formatDate(now, 'Europe/Madrid', 'MM/dd/yyyy\'T\'HH:mm:ss.SSS');
  return now.getTime() - startTime.getTime() > 1740000; // 29.0 minutes
  Logger.log ("NOW Time: "+now);
  Logger.log("NOWFORMATEADO: "+ nowFormateado);
}



// FUNCION PARA VERIFICAR SI ESTUVO PRESENTE
function checkMeetCodeWebMovil(meetCode, index, searchDate, hoja) {
try{
  var ss=SpreadsheetApp.getActive().getSheetByName(hoja);
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
   
  Logger.log("Index / Alumno: "+index+" / " + emailAddress);
  
  var optionalArgs = {
      event_name: "call_ended",
      endTime: searchDate,   
      startTime: fechaRestada(searchDate),      
      filters: "identifier==" + emailAddress + ",meeting_code==" + meetCode 
    };

      const response = AdminReports.Activities.list(userKey, applicationName, optionalArgs)
       // Logger.log("RESPONSE: "+response);  
      var myItems = response.items;
      //  Logger.log("MyITEMS ALL = "+myItems);
      if (myItems == null) {                           // Si no encuentra "Object property" ITEMS mark ABSENT 
         printAbsent(ss,index+3,7)
        Logger.log("Mark as Absent - No 'Items present'")
      }
      else {                                           // Si encuentra "Object property" ITEMS mark PRESENT  
        
        for (i = 0; i < myItems.length; i++){
          var insideItemsI = myItems[i]; // Returns OBJECT
        //  Logger.log("Object insideItems[i]= "+i + "/"+insideItemsI);     // Logs every ITEMS found 
          
            var theEvent = insideItemsI.events; // Entra en el Object Property EVENTS
            var parametros = theEvent[0].parameters // Entra en el primer Array dentro de EVENTS y Propiedad PARAMETERS dentro de ese array.
            if(parametros);{
                   Logger.log("Parametros = "+ parametros) // Logs every PARAMETER found with all its object properties.
                  var parameterValues = getParameterValues(parametros);
                  var duration = parameterValues['duration_seconds'];
                  var device = parameterValues['device_type'];
                  duration = +duration; // convierte el STRING que devuelce duration a un INTEGER (numero)
                  Logger.log("DURATION_S = "+duration);
                  Logger.log("DEVICE = "+ device);

                  if(device === "web"){
                  durationWeb = durationWeb + duration;
                   } else {
                  durationMovil = durationMovil + duration;                
                   }
                  Logger.log("DURATION_WEB Parcial "+i +" = "+ durationWeb )
                  Logger.log("DURATION_MOVIL Parcial "+i +" = "+ durationMovil )
          }           
        }
    Logger.log("DURATION WEB TOTAL x USUARIO = "+durationWeb);
    Logger.log("DURATION MOVIL TOTAL x USUARIO = "+durationMovil)
    Logger.log ("Escribiendo Resultados en Sheets");
    
    printPresent(ss,index+3,7);
    if (durationWeb > 0){ 
    printWebDuration(durationWeb, ss,index+3,8);
      }
    if (durationMovil > 0){ 
    printMovilDuration(durationMovil, ss,index+3,9);
      }
   }
} catch (e){
  Logger.log("Error logged: "+e)
          }
// Fin de la funcion  
} 
 

// * Gets a map of parameter names to values from an array of parameter objects.
// * @param {Array} parameters An array of parameter objects.
// * @return {Object} A map from parameter names to their values.
// Exemplo sacardo de Google Developers Website
 
function getParameterValues(parameters) {
try{ 
    return parameters.reduce(function(result, parameter) {
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

} catch (e){
  Logger.log("Error logged: "+e)
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
    cell.setValue(print/60);
}

function printMovilDuration(print, sheet, r, c) {
    var cell = sheet.getRange(r, c);
    cell.setValue(print/60);
}


function fechaRestada(fecha, diferencia){
  var MILLIS_PER_DAY = 1000 * 60 * 60 * 24;
  var now = new Date(fecha);
  var specificDate = new Date(now.getTime() - (0.875 * MILLIS_PER_DAY));
  //var specificDate = new Date(now.getTime() - (0.875 * MILLIS_PER_DAY));
  var timeZone = 'America/Argentina/Buenos_Aires';
  
  var todo = Utilities.formatDate(specificDate, timeZone, 'yyyy-MM-dd\'T\'HH:mm:ss.SSS\'Z\'');
  return  todo;
}

function fechaMasUno(fecha, diferencia){
  var MILLIS_PER_DAY = 1000 * 60 * 60 * 24;
  var now = new Date(fecha);
  var specificDate = new Date(now.getTime() + (diferencia * MILLIS_PER_DAY));
  var timeZone = 'America/Argentina/Buenos_Aires';
  
  var todo = Utilities.formatDate(specificDate, timeZone, 'yyyy-MM-dd\'T\'HH:mm:ss.SSS\'Z\'');
  return  todo;
}



/////////////////////////////////////////////////////////
/////////////// Duplicados y guardar informes
/////////////////////////////////////////////////////////

function Informe_Duplicar() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('DUPLICAR'), true);
  spreadsheet.duplicateActiveSheet();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Copia de DUPLICAR'), true);
  spreadsheet.getRange('A3').activate();
}

function Informe_Copiar_Values() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Copia de DUPLICAR'), true);
  spreadsheet.getRange('A3:L').activate();
  spreadsheet.getRange('A3:L').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getRange('E1:E2').activate();
  spreadsheet.getRange('E1:E2').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);
  spreadsheet.getRange('A3').activate();
}

// FUNCION QUE CREAR UNA NUEVA SPREADSHEET CON LOS PARAMETROS ESPECIFICADOS.
function create_New_Document(sheetName, folder, fileName){
    var ss = SpreadsheetApp.create(fileName+sheetName);
    var id = ss.getId();
    var file = DriveApp.getFileById(id);
    folder.addFile(file);
      
    var newSpreadsheet = SpreadsheetApp.openById(id);
     
    return newSpreadsheet; 
}



function Informe_Modificar() {
try{
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.setActiveSheet(spreadsheet.getSheetByName('Copia de DUPLICAR'), true);
  spreadsheet.getRange('J1:J2').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getActiveRange().breakApart();
  spreadsheet.getActiveRangeList().setBackground(null);
  
  var fecha = spreadsheet.getRange('L10').activate().getValue();
//  Logger.log(fecha);
  var timeZone = 'Europe/Madrid';
  var fechaFormat = Utilities.formatDate(fecha, timeZone, 'dd/MM/yy');
//  Logger.log(fechaFormat);
  
  spreadsheet.getRange('A1:D2').activate();
  spreadsheet.getCurrentCell().setValue(fechaFormat);
  spreadsheet.getActiveSheet().setName(fechaFormat);
  
  spreadsheet.getRange('H4:H').activate();
  spreadsheet.getActiveRangeList().setNumberFormat('0');
  spreadsheet.getRange('H2').activate();
  spreadsheet.getActiveRangeList().setNumberFormat('0.00%');
  
  
  
  ///////// LLAMA A LA FUNCION PARA CREAR NUEVO DOC Y COPIAR TODA LA DATA CON FORMATOS.
  var destinationSS = create_New_Document(fechaFormat, DriveApp.getFolderById(driveFolderID), 'Asistencia TODOS: ');  // should return "newSpreadsheet" 
  var sheetToCopy = sheet.copyTo(destinationSS);  // copia la hoja a la nueva Spreadsheet
  destinationSS.setActiveSheet(destinationSS.getSheets()[1]).setName(fechaFormat);
  destinationSS.setActiveSheet(destinationSS.getSheets()[0]) // Activa la hoja 1
  destinationSS.deleteActiveSheet(); // elimina la hoja 1
  spreadsheet.getRange('A3').activate();
  
  newDocID = destinationSS.getId();
  Logger.log("newDocID = "+ newDocID);
  
  MakeCopySede(newDocID, fechaFormat);
} catch (e){
  Logger.log("Error logged: "+e)
           }  
}


function LimpiarHojasParciales() {
  var spreadsheet=SpreadsheetApp.openByUrl(docURL);
  // var spreadsheet = SpreadsheetApp.getActive();
  // spreadsheet.setActiveSheet(spreadsheet.getSheetByName('MeetParcial1'), true);
  spreadsheet.setActiveSheet(spreadsheet.getSheetByName('MeetParcial1'), true);
  spreadsheet.getRange('G3:I').activate();
  spreadsheet.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
  spreadsheet.getRange('G2').activate();
}



