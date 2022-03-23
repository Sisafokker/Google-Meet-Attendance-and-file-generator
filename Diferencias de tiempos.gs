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


function fechaRestada(fecha, diferencia) {
  var MILLIS_PER_DAY = 1000 * 60 * 60 * 24;
  var now = new Date(fecha);
  var specificDate = new Date(now.getTime() - (0.875 * MILLIS_PER_DAY));
  //var specificDate = new Date(now.getTime() - (0.875 * MILLIS_PER_DAY));
  var timeZone = 'America/Argentina/Buenos_Aires';
  var todo = Utilities.formatDate(specificDate, timeZone, 'yyyy-MM-dd\'T\'HH:mm:ss.SSS\'Z\'');
  return todo;
}

function fechaMasUno(fecha, diferencia) {
  var MILLIS_PER_DAY = 1000 * 60 * 60 * 24;
  var now = new Date(fecha);
  var specificDate = new Date(now.getTime() + (diferencia * MILLIS_PER_DAY));
  var timeZone = 'America/Argentina/Buenos_Aires';

  var todo = Utilities.formatDate(specificDate, timeZone, 'yyyy-MM-dd\'T\'HH:mm:ss.SSS\'Z\'');
  return todo;
}



