
function alle2Daten(){
var sheet1=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Zusammenfassung VA November")
sheet1.clear();
 
// Funktioniert nur wenn Daten in der Tabelle enthalten sind.
var sheet2=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("04.11.19");
var sheet3=SpreadsheetApp.getActiveSpreadsheet().getSheetByName("05.11.19");


  
  
sheet2.getRange(1,1,sheet2.getLastRow(),sheet2.getLastColumn()).copyTo(sheet1.getRange(1,1));
sheet3.getRange(1,1,sheet3.getLastRow(),sheet3.getLastColumn()).copyTo(sheet3.getRange(sheet1.getLastRow()+2,1));


  
  
  
  
  
  
  
  
 
//sheet1.sort(1);
 
}
