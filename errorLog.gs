//  DATA STORAGE ERROR CODES on ScriptLOG


function errorLogStart(){
var sheetName = 'ScriptLOG';
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName(sheetName);
var uName = sheet.getRange('B2').getValue();
var sName = uName.slice(0, 4);
var status = sheet.getRange('C2').getValue();
var fName = sheet.getRange('D2').getValue();
var info = sheet.getRange('E2').getValue();
var details = sheet.getRange('F2').getValue();
var eeSheetName = sheet.getRange('G2').getValue();
var nowDate = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd hh:mm:ss");


sheet.appendRow([nowDate, sName, status, fName, info, details, eeSheetName]);


}


function eeCode(){
//=======================================================================================//
///////////////////////////////////////////////START ERROR LOG - BLOCK START/////////////// E
var sheetName = 'ScriptLOG';                                                             // R
var ss = SpreadsheetApp.getActiveSpreadsheet();                                          // R
var sheet = ss.getSheetByName(sheetName);                                                // O
var eeUserName = Session.getEffectiveUser();                                             // R
var eeinfo ="Error Step";                                                                // 
var eedetails ="Testing Error Code";                                                     // L
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
//CODE
//CODE BLOCK END

//---------------------------------------------------------------------------------------//
////////////////////////////////////////////////////END ERROR LOG - AFTER CODE///////////// E
var sheetName = 'ScriptLOG';                                                             // R
var ss = SpreadsheetApp.getActiveSpreadsheet();                                          // R
var sheet = ss.getSheetByName(sheetName);                                                // O
var eeUserName = Session.getEffectiveUser();                                             // R
var eeinfo ="Error Step";                                                                // 
var eedetails ="Testing Error Code";                                                     // L
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
