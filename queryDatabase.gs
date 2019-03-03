//Building query script


//sheet NewPlayers
function NewPlayersDBQueryUpdate(){
var sheetname = 'NewPlayers'
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName(sheetname);
var cell = sheet.getRange('A1')
cell.setValue('MERGE!A:AD');
}




//Sheet QForumsDB
function QForumsDBQueryUpdate(){
var sheetname = 'QForumsDB'
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName(sheetname);
var cell = sheet.getRange('A1')
cell.setValue('MERGE!A:AD');
}

//Sheet Roster
function rosterQueryUpdate(){
var sheetname = 'TG Roster'
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName(sheetname);
var cell = sheet.getRange('A1')
cell.setValue('MERGE!A:AD');
}

function qRecruitsQueryUpdate(){
var sheetname = 'QRecruits'
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName(sheetname);
var cell = sheet.getRange('A1')
cell.setValue('MERGE!A:AD');
}

function QAFK7DQueryUpdate(){
var sheetname = 'QAFK7D'
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName(sheetname);
var cell = sheet.getRange('A1')
cell.setValue('MERGE!A:AD');
}


function QCorpRdyAFK30DQueryUpdate(){
var sheetname = 'QCorpRdyAFK30D'
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName(sheetname);
var cell = sheet.getRange('A1')
cell.setValue('MERGE!A:AD');
}



function QAFK2YQueryUpdate(){
var sheetname = 'QAFK2Y'
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName(sheetname);
var cell = sheet.getRange('A1')
cell.setValue('MERGE!A:AD');
}


function TestLookUpQueryUpdate(){
var sheetname = 'TestLookUp'
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName(sheetname);
var cell = sheet.getRange('A1')
cell.setValue('MERGE!A:AD');
}

function queryUpdadeAll(){
rosterQueryUpdate();
qRecruitsQueryUpdate();
QAFK7DQueryUpdate();
QCorpRdyAFK30DQueryUpdate();
QAFK2YQueryUpdate();
TestLookUpQueryUpdate();
QForumsDBQueryUpdate();
NewPlayersDBQueryUpdate();
}
//Update current queries
