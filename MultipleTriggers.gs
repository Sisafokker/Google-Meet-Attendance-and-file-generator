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

function setUpOneTimeTrigger() {
  ScriptApp.newTrigger("hojaMeetParcial1")
    .timeBased()
    .atHour(13)
    .nearMinute(45)
    .everyDays(1)
    .inTimezone('Europe/Madrid')
    .create();
}

function one (){
var nowTime = new Date();
var espera1 = 7 * 60 * 1000;
  Logger.log(nowTime);
  Logger.log(espera1);
var timeFormateado = Utilities.formatDate(nowTime, 'Europe/Madrid', 'MM/dd/yyyy\'T\'HH:mm:ss.SSS');
GmailApp.sendEmail('jp@liceobritanico.com', 'Testeando Triggers automaticos', timeFormateado);
};