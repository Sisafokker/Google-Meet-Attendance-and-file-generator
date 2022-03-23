/////////////////////////////////////////////////////////////
/////////// CREACION DE TRIGGERS ////////////////////////////
/////////////////////////////////////////////////////////////

// Source: https://www.youtube.com/watch?v=5BYhGGPQlyA
function setUpMainTrigger(){
ScriptApp.newTrigger("hojaMeetParcial1")
 .timeBased()
 .atHour(6)
 .nearMinute(15)
 .everyDays(1)
 .inTimezone('Europe/Madrid')
 .create();
}

function setUpTemporaryTrigger(){
  var espera = 3 * 60 * 1000; // 3 minutos
  ScriptApp.newTrigger('hojaMeetParcial1')
  .timeBased()
  .after(espera)
  .create()
}


function setUpHacerTodoTrigger(){
  var espera = 3 * 60 * 1000; // 3 minutos
  ScriptApp.newTrigger('HACER_TODO')
  .timeBased()
  .after(espera)
  .create()
}

function setUpOneTimeTrigger(){
 ScriptApp.newTrigger("hojaMeetParcial1")
 .timeBased()
 .atHour(13)
 .nearMinute(45)
 .everyDays(1)
 .inTimezone('Europe/Madrid')
 .create();

}

function testing(){
  var sheetHoja = "testing";
  verAsistenciaMeetTEST(sheetHoja);
  }

function enviarEmail (cuerpo){
  var nowTime = new Date();
  var timeFormateado = Utilities.formatDate(nowTime, 'Europe/Madrid', 'MM/dd/yyyy\'T\'HH:mm:ss.SSS');
  GmailApp.sendEmail('jp@liceobritanico.com', 'REPORTE ASISTENCIAS LCB- '+cuerpo,cuerpo +' '+ timeFormateado);
};


function one(){
  var nowTime = new Date();
  var espera1 = 7 * 60 * 1000;
    Logger.log(nowTime);
    Logger.log(espera1);
  var timeFormateado = Utilities.formatDate(nowTime, 'Europe/Madrid', 'MM/dd/yyyy\'T\'HH:mm:ss.SSS');
  GmailApp.sendEmail('jp@liceobritanico.com', 'Testeando Triggers automaticos', timeFormateado);
};


function correrDiferencias(){
  var comienzo = new Date();
  var startTime= Utilities.formatDate(comienzo, "GMT-3", 'dd-MM-yyyy HH:mm:ss.SSS')
  Logger.log(comienzo);
  Logger.log(startTime);
  var minOne = 60000;
  var minFive = 300000;
  var minSix = 360000;
  isTimeUp(comienzo,startTime, minFive);
}

function isTimeUp(comienzo, startTime, limite) {
    var fin = new Date();
    var endTime= Utilities.formatDate(fin, "GMT-3", 'dd-MM-yyyy HH:mm:ss.SSS')
    var diferencia = (fin.getTime() - comienzo.getTime())
    Logger.log("startTime: "+startTime);
    Logger.log('endTime: '+endTime);
    Logger.log("diferencia: "+diferencia);

    if(diferencia < limite ) {
    Logger.log("Quedan: "+(limite - diferencia));
    Logger.log("Continuar Loop");
    } 
    else {
    Logger.log("Imprimir, Frenar y crear TimeTrigger en 5 min");  
    } 
}

