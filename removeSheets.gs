////////////////////REMOVE SHEETS

//function cleanStart(){
//removeAppSheet();
//removeMergeSheet();
//removeDumpSheet();
//}

function removeAllDataSheets(){

//=======================================================================================//
///////////////////////////////////////////////START ERROR LOG - BLOCK START/////////////// E
var sheetName = 'ScriptLOG';                                                             // R
var ss = SpreadsheetApp.getActiveSpreadsheet();                                          // R
var sheet = ss.getSheetByName(sheetName);                                                // O
var eeUserName = Session.getEffectiveUser();                                             // R
var eeinfo ="Remove Sheets";                                                             // 
var eedetails ="Remove APPDATA, DATADUMP Sheets";                                        // L
var eeStatus="Started";                                                                  // O
var fName = arguments.callee.toString().match(/function ([^\(]+)/)[1]                    // G
var eeSheetName = sheetName;                                                             //
var values = [                                                                           // S
[eeUserName,eeStatus, fName, eeinfo,eedetails, eeSheetName]                              // T
];                                                                                       // A
var range = sheet.getRange("B2:G2");                                                     // R
range.setValues(values);                                                                 // T
errorLogStart();                                                                         //
/////////////////////////////////////////////END ERROR LOG BLOCK - BEFORE CODE/////////////
//---------------------------------------------------------------------------------------//


//CODE BLOCK START
removeDumpSheet();
//removeMergeSheet();
removeAppSheet();
//CODE BLOCK END

//---------------------------------------------------------------------------------------//
////////////////////////////////////////////////////END ERROR LOG - AFTER CODE///////////// E
var sheetName = 'ScriptLOG';                                                             // R
var ss = SpreadsheetApp.getActiveSpreadsheet();                                          // R
var sheet = ss.getSheetByName(sheetName);                                                // O
var eeUserName = Session.getEffectiveUser();                                             // R
var eeinfo ="Remove Sheets";                                                             // 
var eedetails ="Remove APPDATA, DATADUMP Sheets";                                        // L
var eeStatus="Ended";                                                                    // O
var fName = arguments.callee.toString().match(/function ([^\(]+)/)[1]                    // G
var eeSheetName = sheetName;                                                             // 
var values = [                                                                           // E
[eeUserName,eeStatus, fName, eeinfo,eedetails, eeSheetName]                              // N
];                                                                                       // D
var range = sheet.getRange("B2:G2");                                                     //
range.setValues(values);                                                                 //
errorLogStart();                                                                         //
/////////////////////////////////////////////END ERROR LOG - BEFORE END////////////////////
//=======================================================================================//

}

function removeTestSheet(){
//////////////////////TEST101 SHEET/////////////////////REMOVE//
var ss= SpreadsheetApp.getActiveSpreadsheet()
var sheetname = 'TEST101';
sheetToRemove= ss.getSheetByName(sheetname);
sheetToRemove.activate();
var removeText = sheetname+" Sheet Removed";
ss.deleteActiveSheet();


Browser.msgBox(removeText);
}


function removeDumpSheet(){
//////////////////////DUMPDATA SHEET/////////////////REMOVE//
var ss = SpreadsheetApp.getActiveSpreadsheet(),
sheetToRemove = ss.getSheetByName("DATADUMP");
sheetToRemove.activate();
ss.deleteActiveSheet();
Browser.msgBox("DATADUMP Sheet Removed");
}


//////////////////////MERGE SHEET/////////////////////REMOVE//
function removeMergeSheet(){
var ss = SpreadsheetApp.getActiveSpreadsheet(),
sheetToRemove = ss.getSheetByName("MERGE");
sheetToRemove.activate();
ss.deleteActiveSheet();
Browser.msgBox("MERGE Sheet Removed");
}


//////////////////////APPDATA SHEET/////////////////////REMOVE//
function removeAppSheet(){
var ss = SpreadsheetApp.getActiveSpreadsheet(),
sheetToRemove = ss.getSheetByName("APPDATA");
sheetToRemove.activate();
ss.deleteActiveSheet();
Browser.msgBox("APPDATA Sheet Removed");
}
