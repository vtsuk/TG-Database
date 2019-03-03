function AppLast(){
  var sheetname = 'TGPS2APP'
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetname);
  sheet.getRange('A1').activate();
  sheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.NEXT).activate();
  sheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();

};

function AppLastLog(){
  var sheetname = 'ScriptLOG'
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetname);
  sheet.getRange('A6').activate();
  sheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.NEXT).activate();
  sheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();

};
