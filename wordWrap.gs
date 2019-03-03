//word wrap
// may need to update range once formula changed



function wordWrapOn(){
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName('MERGE');
var range = sheet.getRange("A1:N1000");
range.setWrap(false);

var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName('MEGALIST');
var range = sheet.getRange("A1:N1000");
range.setWrap(false);

}

function wordWrapOff(){
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName('MERGE');
var range = sheet.getRange("A1:N1000");
range.setWrap(true);
}

function wordWrapClip(){
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName('MERGE');
var range = sheet.getRange("A:N");
range.setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
 }
