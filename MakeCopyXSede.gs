function MANUAL_RUN (){
var manualSheetID = '1Yb4nBcFtbpi76WpndQmk9iS-Z_risgWi-3jjg25QCSw'
var manualFecha = '05/05/20'
MakeCopySede(manualSheetID ,manualFecha);
//deleteRowsExcept('Belgrano');
}

// GLOBAL
const allSedes = ["Adrogue","B.Norte","Belgrano","Boedo","Central","Flores","Hurlingham","Lanus","Martinez","Moreno","Palomar","Quilmes","San Martin","V.Crespo","V.Urquiza","Virtual"]



function MakeCopySede(thisURL, thisSheet) {
 try{ 
  var ss = SpreadsheetApp.openById(thisURL);
  var sheet = ss.getSheetByName(thisSheet);
  var sheetID = sheet.getSheetId();
  
  var lastRow = sheet.getLastRow(); // with data
  Logger.log("lastRow = " +lastRow)
  var maxRow = sheet.getMaxRows() // in the sheet
  Logger.log("maxRow = " +maxRow);
  
  var lastColumn = sheet.getLastColumn(); // with data
  var maxColumn = sheet.getMaxColumns(); // in the sheet
  
  
  var rangeColC = sheet.getRange(4, 2, lastRow).getValues();
  Logger.log("RANGE-C ="+rangeColC);
  
  var rows = sheet.getDataRange();  // returns RANGE
  Logger.log("ROWS ="+rows);
  
  var numRows = rows.getNumRows(); // returns 9094 (cantidad total de rows)
  Logger.log("NumROWS ="+numRows);
  
//  var values = rows.getValues();
//  Logger.log("VALUES ="+values); // toda la data
  
  
// sheet.deleteRows(lastRow + 1, maxRow - lastRow);
// var insertUnique = sheet.getRange("M4").setFormula("=UNIQUE(B4:B)");  
 
 for (i = 0; i < allSedes.length; i++){
   Logger.log("Filtrando Sede = "+allSedes[i]);
   filterSedes(thisURL, thisSheet, allSedes[i], sheetID); // Filtra perfecto pero aparentemente es complejo agarrar el range de la data filtrada
//     makeArrayPorSede(thisURL, thisSheet, allSedes[i])  
  
  }
} catch (e){
  Logger.log("Error logged: "+e)
           }     
}


function filterSedes (thisURL, thisSheet, sedePuntual, sheetID){
try{
  var ss = SpreadsheetApp.openById(thisURL);
  var sheet = ss.getSheetByName(thisSheet);
 
 // 1. Crear filtro y aplicar criterio
  var filter = ss.getRange('A3:L').createFilter();
  //var filter = ss.getDataRange().createFilter();
  var filterCriteria = SpreadsheetApp.newFilterCriteria();
  var defineCriteria = filterCriteria.whenTextEqualTo(sedePuntual);
  var applyfilter = filter.setColumnFilterCriteria(2, filterCriteria)
  
  // 2. Guardar info filtrada en un rango
  var returnArray = visibleData(thisURL, sheetID, sedePuntual); // returns array with visibleData to paste
  Logger.log("returnArray.length ="+returnArray.length);
  Logger.log("returnArray[0].length ="+returnArray[0].length);
  
  // 3. Crear nueva hoja
  var destinationSS = create_New_SS(thisSheet, DriveApp.getFolderById(CopyXSedeDriveFolderID), sedePuntual); // driveFolderID es una variable GLOBAL declarada en otra pagina.
  // returns  var newSpreadsheet = SpreadsheetApp.openById(id);
  
  // 4. Pegar Rango info filtrada en nueva hoja
  var destSheet = destinationSS.getSheetByName(thisSheet);
  destSheet.getRange(1, 1, returnArray.length, returnArray[0].length).setValues(returnArray);
  
  // 5. Modifica el formato de la nueva planilla
  modificaFormatosHoja(destSheet); // DA ERROR, hay que mejorarlo
  
  // 5 Remover Filtro del file original para que pueda pueda reiniciarse el loop sin dar error
  var removeFilter = filter.remove()
  //var removeFilter =  ss.getDataRange().createFilter().removeColumnFilterCriteria(2);
} catch (e){
  Logger.log("Error logged: "+e)
           }
}


// Captura la informacion capturada por un filtro en un ARRAY.
//Source= https://tanaikech.github.io/2019/07/28/retrieving-values-from-filtered-sheet-in-spreadsheet-using-google-apps-script/
function visibleData(ssID, sheetID, sede) {
  var spreadsheetId = ssID; // Please set Spreadsheet ID.
  var sheetId = sheetID; // Please set Sheet ID.

  var url =
    "https://docs.google.com/spreadsheets/d/" +
    spreadsheetId +
    "/gviz/tq?tqx=out:csv&gid=" +
    sheetId +
    "&access_token=" +
    ScriptApp.getOAuthToken();
  var res = UrlFetchApp.fetch(url);
  var array = Utilities.parseCsv(res.getContentText());
  Logger.log(sede+ "array = "+ array);
  return array;
}


// FUNCION QUE CREAR UNA NUEVA SPREADSHEET CON LOS PARAMETROS ESPECIFICADOS.
function create_New_SS(sheetName, folder, fileName){
    var ss = SpreadsheetApp.create(sheetName+" "+fileName);
    var id = ss.getId();
    var file = DriveApp.getFileById(id);
    folder.addFile(file);
    
    var newSSxSede = SpreadsheetApp.openById(id);
    newSSxSede.setActiveSheet(newSSxSede.getSheets()[0]).setName(sheetName);
    //newSSxSede.setActiveSheet(sheet)
    
   
    return newSSxSede; 
}

function modificaFormatosHoja(spreadsheet) {
  spreadsheet.getRange('A1:D2').activate().merge();
  spreadsheet.getActiveRangeList().setHorizontalAlignment('center').setVerticalAlignment('middle');
  spreadsheet.getRange('C5').activate();
  spreadsheet.setFrozenRows(3);
  spreadsheet.getRange('A3:L3').activate();
  spreadsheet.getActiveRangeList().setBackground('#d9d9d9');
  
  spreadsheet.getActiveRangeList().setHorizontalAlignment('center').setVerticalAlignment('middle');
  
  spreadsheet.getRange('G1').activate().setFormula('=COUNTIF(D4:D;"=st.*")');
  spreadsheet.getRange('G2').activate().setFormula('=COUNTIF(F1:F;"=PRESENT")');
  spreadsheet.getRange('H2').activate().setFormula('=G2/G1');
  spreadsheet.getRange('J1').activate().setFormula('=countUnique(E1:E)');
  spreadsheet.getRange('G3').activate().setValue('PC (mins)');
  spreadsheet.getRange('H3').activate().setValue('Movil (mins)');
  spreadsheet.getRange('I3').activate().setValue('Class Length (min)');
  spreadsheet.getRange('L3').activate().setValue('Fecha');
  spreadsheet.getRange(1, 1, spreadsheet.getMaxRows(), spreadsheet.getMaxColumns()).activate();
  spreadsheet.autoResizeColumns(1, spreadsheet.getMaxColumns());
  spreadsheet.getRange('M1').activate();
  spreadsheet.getCurrentCell().setValue('delete to the right');
  spreadsheet.getRange('M:M').activate();
  var currentCell = spreadsheet.getCurrentCell();
  spreadsheet.getSelection().getNextDataRange(SpreadsheetApp.Direction.NEXT).activate();
  currentCell.activateAsCurrentCell();
  spreadsheet.deleteColumns(spreadsheet.getActiveRange().getColumn(), spreadsheet.getActiveRange().getNumColumns());
  spreadsheet.getRange('A1:D2').activate();
  spreadsheet.getActiveRangeList().setFontWeight('bold').setFontSize(14);
  
  spreadsheet.getRange('F:I').activate();
  spreadsheet.getActiveRangeList().setHorizontalAlignment('center');
  
  
//  spreadsheet.getRange('F3:I').activate();
//  currentCell = spreadsheet.getCurrentCell();
//  currentCell.activateAsCurrentCell();
//  spreadsheet.getActiveRangeList().setHorizontalAlignment('center');

  spreadsheet.getRange('H1').activate();
};




function makeArrayPorSede (thisURL, thisSheet, sedePuntual){
try{
  var ss = SpreadsheetApp.openById(thisURL);
  var sheet = ss.getSheetByName(thisSheet);
//Source: https://stackoverflow.com/questions/58042502/how-to-copy-filtered-spreadsheet-data-with-apps-script
  var data = []

  for (var i = 1; i < sheet.getLastRow(); i++) {
    if(!sheet.isRowHiddenByFilter(i)) {
      var row_data = sheet.getRange(i, 1, 1, sheet.getLastColumn()).getValues()
      data.push(row_data[0])
    }
 }
 Logger.log("DATA X SEDE ="+data);
} catch (e){
  Logger.log("Error logged: "+e)
          }
}




function deleteRowsExcept() {
// function deleteRowsExcept(exception) {
  var sheet = SpreadsheetApp.getActiveSheet();
  var rows = sheet.getDataRange();
  var numRows = rows.getNumRows();
  var values = rows.getValues();

//  Logger.log("The Exception is = "+exception);

  var rowsInColumn = 0;
  for (var i = 4; i <= numRows - 1; i++) { // i= 4 to save the top 3 rows
    var row = values[i]; // 
    // Logger.log(row); // Shows the content of each row
    
    //if (row[1] !== exception) { // This searches all cells in columns B (change to row[3] for columns C and so on) and deletes row if cell is empty or has value 'delete'.
    if (row[1] !== 'Belgrano') { // This searches all cells in columns B (change to row[3] for columns C and so on) and deletes row if cell is empty or has value 'delete'.
      Logger.log("Should delete Row ="+i);
      sheet.deleteRow((parseInt(i)+1) - rowsInColumn);
      rowsInColumn++;
    }
  }
}