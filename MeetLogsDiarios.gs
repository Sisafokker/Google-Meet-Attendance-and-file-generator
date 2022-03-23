var list =[]

function checkMeetTesting(index, searchDate) {
try{
  var userKey = 'all';
  var applicationName = 'meet';
  var optionalArgs = {
      event_name: "call_ended",
    };
  var parameters = [
                    'meeting_code'
                  ];

 list.push(['Codigo Sucio','Fecha Sucia','Usuario']);

 const response = AdminReports.Activities.list(userKey, applicationName, optionalArgs)
 var myItems = response.items;
  if (myItems && myItems.length > 0) {
    // console.log (myItems);
    for (i = 0; i < myItems.length; i++) {
      console.log(i);
      var row = [] // Array para la fila
      var datos = myItems[i];
      var theEvent = datos.events; // Entra en el Object Property EVENTS
      var parametros = theEvent[0].parameters // Entra en el primer Array dentro de EVENTS y Propiedad PARAMETERS dentro de ese array.
      var correo = datos.actor.email;

      if(parametros);{
            var parameterValues = getParameterValues(parametros);
            var codigoMeet = parameterValues['meeting_code'];                      
          
          console.log(correo);
          if(correo && !correo.includes("st.")){ // Similar a indexOf . Si correo NO contiene "st."
               row.push(codigoMeet, datos.id.time, correo) // Empujamos fila
          console.log(row);
            }
      }
   
      if (row && row.length>0){
          console.log("Pushing row:" + row)    
          list.push(row); // Empujamos todo el array filas
     }
 }
   console.log("---FIN---");
   console.log("List Length =", list.length);
   console.log("List [0] Length =", list[0].length);
   console.log("---FIN---");

   if(list.length > 0) {
       SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Meet Diarios Test").getRange(2,5,list.length,list[1].length).setValues(list);
      }
  }
   else {
    console.log('No logins found.');
  }
} catch (e){
  console.log("Error logged: "+e)
  }    
} 


// Va,falta conseguir el meet code
function checkMeetOriginal(index, searchDate) {
 try{
    var userKey = 'all';
    var applicationName = 'meet';
    var optionalArgs = {
        event_name: "call_ended",
      };
    const response = AdminReports.Activities.list(userKey, applicationName, optionalArgs)
    var myItems = response.items;
  
    if (myItems && myItems.length > 0) {
      console.log (myItems);
      console.log('XYZ:');
      for (i = 0; i < myItems.length; i++) {
        var dato = myItems[i];
        //console.log('%s: %s (%s)', dato.id.time, dato.actor.email, dato.events[0].name);
        console.log("--------------------");
        console.log(i);
        console.log('%s: %s',dato.id.time, dato.actor.email);
        console.log("--------------------");
        }
     } else {
      console.log('Nothing found.');
     }
   } catch (e){
    console.log("Error logged: "+e)
    }    
} 