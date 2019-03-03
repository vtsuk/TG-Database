//Show Sheets

function showLOG(){
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName('ScriptLOG');
sheet.activate();
}

function showMerge(){
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName('MERGE');
sheet.activate();
}

function showRecruits(){
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName('Recruits');
sheet.activate();
}

function showAppList(){
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName('TGPS2APP');
sheet.activate();
}

function showMegaList(){
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName('MEGALIST');
sheet.activate();
}

function showDataDump(){
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName('DATADUMP');
sheet.activate();
}

function showAppData(){
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName('APPDATA');
sheet.activate();
}

function showInstructions(){
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName('Instructions');
sheet.activate();
}

function showToDo(){
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName('To Do');
sheet.activate();
}
