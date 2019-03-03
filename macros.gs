function AppLastdata() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A4').activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.NEXT).activate();
};

function Untitledmacro() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A2').activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
  spreadsheet.getRange('J143').activate();
};

function Untitledmacro1() {
  var spreadsheet = SpreadsheetApp.getActive();
  spreadsheet.getRange('A1').activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.NEXT).activate();
  spreadsheet.getCurrentCell().getNextDataCell(SpreadsheetApp.Direction.DOWN).activate();
  spreadsheet.getRange('J143').activate();
  spreadsheet.getCurrentCell().offset(1, 0).activate();

};

