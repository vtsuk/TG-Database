//BUILD MERGE SHEET BUILDMERGE
//First Task is to Move all work relating to MERGE to this code base.


/////////////////////////////////////////////////////
/////////Building MERGE SHEET - Step 5
/////////////////////////////////////////////////////

function makeMerge(){

//=======================================================================================//
///////////////////////////////////////////////START ERROR LOG - BLOCK START/////////////// E
var sheetName = 'ScriptLOG';                                                             // R
var ss = SpreadsheetApp.getActiveSpreadsheet();                                          // R
var sheet = ss.getSheetByName(sheetName);                                                // O
var eeUserName = Session.getEffectiveUser();                                             // R
var eeinfo ="Step 5";                                                                    // 
var eedetails ="Building Headers MERGE SHEET";                                           // L
var eeStatus="Started";                                                                  // O
var fName = arguments.callee.toString().match(/function ([^\(]+)/)[1]                    // G
var eeSheetName = "MERGE";                                                               //
var values = [                                                                           // S
[eeUserName,eeStatus, fName, eeinfo,eedetails, eeSheetName]                              // T
];                                                                                       // A
var range = sheet.getRange("B2:G2");                                                     // R
range.setValues(values);                                                                 // T
errorLogStart();                                                                         //
/////////////////////////////////////////////END ERROR LOG BLOCK - BEFORE CODE/////////////
//---------------------------------------------------------------------------------------//

//CODE BLOCK START

createHeaders();


//CODE BLOCK END

//---------------------------------------------------------------------------------------//
////////////////////////////////////////////////////END ERROR LOG - AFTER CODE///////////// E
var sheetName = 'ScriptLOG';                                                             // R
var ss = SpreadsheetApp.getActiveSpreadsheet();                                          // R
var sheet = ss.getSheetByName(sheetName);                                                // O
var eeUserName = Session.getEffectiveUser();                                             // R
var eeinfo ="Step 5";                                                                    // 
var eedetails ="Building Headers MERGE SHEET";                                           // L
var eeStatus="Ended";                                                                    // O
var fName = arguments.callee.toString().match(/function ([^\(]+)/)[1]                    // G
var eeSheetName = "MERGE";                                                               //
var values = [                                                                           // S
[eeUserName,eeStatus, fName, eeinfo,eedetails, eeSheetName]                              // N
];                                                                                       // D
var range = sheet.getRange("B2:G2");                                                     //
range.setValues(values);                                                                 //
errorLogStart();                                                                         //
/////////////////////////////////////////////END ERROR LOG - BEFORE END////////////////////
//=======================================================================================//

//Browser.msgBox("Step 5 - makeMerge Finnished");
}

/////////////////////////////////////////////////////
/////////CREATE HEADERS START
/////////////////////////////////////////////////////

function createHeaders() {

//=======================================================================================//
///////////////////////////////////////////////START ERROR LOG - BLOCK START/////////////// E
var sheetName = 'ScriptLOG';                                                             // R
var ss = SpreadsheetApp.getActiveSpreadsheet();                                          // R
var sheet = ss.getSheetByName(sheetName);                                                // O
var eeUserName = Session.getEffectiveUser();                                             // R
var eeinfo ="Headers";                                                                   // 
var eedetails ="Create Headers, resize cells";                                           // L
var eeStatus="Started";                                                                  // O
var fName = arguments.callee.toString().match(/function ([^\(]+)/)[1]                    // G
var eeSheetName = "MERGE";                                                               //
var values = [                                                                           // S
[eeUserName,eeStatus, fName, eeinfo,eedetails, eeSheetName]                              // T
];                                                                                       // A
var range = sheet.getRange("B2:G2");                                                     // R
range.setValues(values);                                                                 // T
errorLogStart();                                                                         //
/////////////////////////////////////////////END ERROR LOG BLOCK - BEFORE CODE/////////////
//---------------------------------------------------------------------------------------//

//CODE BLOCK START
//Creates Headers and resizes Cells to auto

var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName('MERGE');
sheet.setFrozenRows(2);
sheet.setFrozenColumns(2);
var values = [
  ["Members Character Id", "Members Name First", "Forum Name", "Recruited by", "Extra Information", "Role", "SL", "Leadership Abilities / Intentions", "Additional Information / Background", "Web App Date", "Rank", "Members Rank Ordinal", "Join Date", "Last Login Date", "X","X", "Corp Offer", "Days", "Till Corp", "X", "EXPIRE DATE", "STATUS", "Days Left",	"Days", "X", "6Ms", "12Ms", "2Ys",	"Active", "AFK", "X"]
];
var range = sheet.getRange("A2:AE2");
range.setValues(values);

var nowDate = Utilities.formatDate(new Date(), "GMT", "yyyy-MM-dd HH:MM.SS");
var buildDate = "Build Date: " + nowDate

var values = [
  [buildDate, "", "", "", "", "", "", "", "", "", "", "", "", "", "X","X", "", "", "", "X", "", "", "","", "X", "=SUM(Z3:Z1001)", "=SUM(AA3:AA1001)", "=SUM(AB3:AB1001)",	"=SUM(AC3:AC1001)", "Days", "X"]
];
var range = sheet.getRange("A1:AE1");
range.setValues(values);

//CODE BLOCK END

//---------------------------------------------------------------------------------------//
////////////////////////////////////////////////////END ERROR LOG - AFTER CODE///////////// E
var sheetName = 'ScriptLOG';                                                             // R
var ss = SpreadsheetApp.getActiveSpreadsheet();                                          // R
var sheet = ss.getSheetByName(sheetName);                                                // O
var eeUserName = Session.getEffectiveUser();                                             // R
var eeinfo ="Headers";                                                                   // 
var eedetails ="Create Headers, resize cells";                                           // L
var eeStatus="Ended";                                                                    // O
var fName = arguments.callee.toString().match(/function ([^\(]+)/)[1]                    // G
var eeSheetName = "MERGE";                                                               // 
var values = [                                                                           // E
[eeUserName,eeStatus, fName, eeinfo,eedetails, eeSheetName]                              // N
];                                                                                       // D
var range = sheet.getRange("B2:G2");                                                     //
range.setValues(values);                                                                 //
errorLogStart();                                                                         //
/////////////////////////////////////////////END ERROR LOG - BEFORE END////////////////////
//=======================================================================================//


//MAIN ERROR BLOCK END
}

/////////////////////////////////////////////////////
/////////CREATE HEADERS END
/////////////////////////////////////////////////////




function makeMergeB(){

//=======================================================================================//
///////////////////////////////////////////////START ERROR LOG - BLOCK START/////////////// E
var sheetName = 'ScriptLOG';                                                             // R
var ss = SpreadsheetApp.getActiveSpreadsheet();                                          // R
var sheet = ss.getSheetByName(sheetName);                                                // O
var eeUserName = Session.getEffectiveUser();                                             // R
var eeinfo ="Step 5B";                                                                   // 
var eedetails ="Adding Data to MERGE Sheet";                                             // L
var eeStatus="Started";                                                                  // O
var fName = arguments.callee.toString().match(/function ([^\(]+)/)[1]                    // G
var eeSheetName = "MERGE";                                                             //
var values = [                                                                           // S
[eeUserName,eeStatus, fName, eeinfo,eedetails, eeSheetName]                              // T
];                                                                                       // A
var range = sheet.getRange("B2:G2");                                                     // R
range.setValues(values);                                                                 // T
errorLogStart();                                                                         //
/////////////////////////////////////////////END ERROR LOG BLOCK - BEFORE CODE/////////////
//---------------------------------------------------------------------------------------//


sortDataDumpByRank();
pasteValues();
//---------------------------------------------------------------------------------------//
////////////////////////////////////////////////////END ERROR LOG - AFTER CODE///////////// E
var sheetName = 'ScriptLOG';                                                             // R
var ss = SpreadsheetApp.getActiveSpreadsheet();                                          // R
var sheet = ss.getSheetByName(sheetName);                                                // O
var eeUserName = Session.getEffectiveUser();                                             // R
var eeinfo ="Step 5B";                                                                   // 
var eedetails ="Adding Data to MERGE Sheet";                                             // L
var eeStatus="Ended";                                                                    // O
var fName = arguments.callee.toString().match(/function ([^\(]+)/)[1]                    // G
var eeSheetName = "MERGE";                                                             // 
var values = [                                                                           // E
[eeUserName,eeStatus, fName, eeinfo,eedetails, eeSheetName]                              // N
];                                                                                       // D
var range = sheet.getRange("B2:G2");                                                     //
range.setValues(values);                                                                 //
errorLogStart();                                                                         //
/////////////////////////////////////////////END ERROR LOG - BEFORE END////////////////////
//=======================================================================================//

//Browser.msgBox("Step 5B - makeMerge Finnished");
}
/////////////////////////////////////////////////////
/////////Building MERGE SHEET END
/////////////////////////////////////////////////////


/////////////////////////////////////////////////////
/////////PASTE VALUES START
/////////////////////////////////////////////////////

function pasteValues() {
//convert formula to values

//=======================================================================================//
///////////////////////////////////////////////START ERROR LOG - BLOCK START/////////////// E
var sheetName = 'ScriptLOG';                                                             // R
var ss = SpreadsheetApp.getActiveSpreadsheet();                                          // R
var sheet = ss.getSheetByName(sheetName);                                                // O
var eeUserName = Session.getEffectiveUser();                                             // R
var eeinfo ="Value";                                                                     // 
var eedetails ="Fix formula to Value";                                                   // L
var eeStatus="Started";                                                                  // O
var fName = arguments.callee.toString().match(/function ([^\(]+)/)[1]                    // G
var eeSheetName = "DATADUMP";                                                            //
var values = [                                                                           // S
[eeUserName,eeStatus, fName, eeinfo,eedetails, eeSheetName]                              // T
];                                                                                       // A
var range = sheet.getRange("B2:G2");                                                     // R
range.setValues(values);                                                                 // T
errorLogStart();                                                                         //
/////////////////////////////////////////////END ERROR LOG BLOCK - BEFORE CODE/////////////
//---------------------------------------------------------------------------------------//

//CODE BLOCK START

//showDataDump();
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName('DATADUMP');
sheet.getRange("A:V").copyTo(sheet.getRange("A:V"), {contentsOnly:true});

//CODE BLOCK END

//---------------------------------------------------------------------------------------//
////////////////////////////////////////////////////END ERROR LOG - AFTER CODE///////////// E
var sheetName = 'ScriptLOG';                                                             // R
var ss = SpreadsheetApp.getActiveSpreadsheet();                                          // R
var sheet = ss.getSheetByName(sheetName);                                                // O
var eeUserName = Session.getEffectiveUser();                                             // R
var eeinfo ="Value";                                                                     // 
var eedetails ="Fix formula to Value";                                                   // L
var eeStatus="Ended";                                                                    // O
var fName = arguments.callee.toString().match(/function ([^\(]+)/)[1]                    // G
var eeSheetName = "DATADUMP";                                                            // 
var values = [                                                                           // E
[eeUserName,eeStatus, fName, eeinfo,eedetails, eeSheetName]                              // N
];                                                                                       // D
var range = sheet.getRange("B2:G2");                                                     //
range.setValues(values);                                                                 //
errorLogStart();                                                                         //
/////////////////////////////////////////////END ERROR LOG - BEFORE END////////////////////
//=======================================================================================//

//MAIN ERROR BLOCK END

}

/////////////////////////////////////////////////////
/////////PASTE VALUES END
/////////////////////////////////////////////////////



/////////////////////////////////////////////////////
/////////Build Merge sheet - Step 6
/////////////////////////////////////////////////////


function callbuildMergeSheet(){

//=======================================================================================//
///////////////////////////////////////////////START ERROR LOG - BLOCK START/////////////// E
var sheetName = 'ScriptLOG';                                                             // R
var ss = SpreadsheetApp.getActiveSpreadsheet();                                          // R
var sheet = ss.getSheetByName(sheetName);                                                // O
var eeUserName = Session.getEffectiveUser();                                             // R
var eeinfo ="Step 6";                                                                    // 
var eedetails ="Build Merge sheet";                                                      // L
var eeStatus="Started";                                                                  // O
var fName = arguments.callee.toString().match(/function ([^\(]+)/)[1]                    // G
var eeSheetName = "MERGE";                                                               //
var values = [                                                                           // S
[eeUserName,eeStatus, fName, eeinfo,eedetails, eeSheetName]                              // T
];                                                                                       // A
var range = sheet.getRange("B2:G2");                                                     // R
range.setValues(values);                                                                 // T
errorLogStart();                                                                         //
/////////////////////////////////////////////END ERROR LOG BLOCK - BEFORE CODE/////////////
//---------------------------------------------------------------------------------------//


//CODE BLOCK START

buildMergeSheet();
/////////////////
paintColor();
/////////////////

//CODE BLOCK END

//---------------------------------------------------------------------------------------//
////////////////////////////////////////////////////END ERROR LOG - AFTER CODE///////////// E
var sheetName = 'ScriptLOG';                                                             // R
var ss = SpreadsheetApp.getActiveSpreadsheet();                                          // R
var sheet = ss.getSheetByName(sheetName);                                                // O
var eeUserName = Session.getEffectiveUser();                                             // R
var eeinfo ="Step 6";                                                                    // 
var eedetails ="Build Merge sheet";                                                      // L
var eeStatus="Ended";                                                                    // O
var fName = arguments.callee.toString().match(/function ([^\(]+)/)[1]                    // G
var eeSheetName = "MERGE";                                                               // 
var values = [                                                                           // E
[eeUserName,eeStatus, fName, eeinfo,eedetails, eeSheetName]                              // N
];                                                                                       // D
var range = sheet.getRange("B2:G2");                                                     //
range.setValues(values);                                                                 //
errorLogStart();                                                                         //
/////////////////////////////////////////////END ERROR LOG - BEFORE END////////////////////
//=======================================================================================//

//MAIN ERROR BLOCK END

//Browser.msgBox("Step 6 - callbuildMergeSheet Finnished");
}


/////////////////////////////////////////////////////
/////////BUILD MERGE SHEET START
/////////////////////////////////////////////////////

function buildMergeSheet(){

//=======================================================================================//
///////////////////////////////////////////////START ERROR LOG - BLOCK START/////////////// E
var sheetName = 'ScriptLOG';                                                             // R
var ss = SpreadsheetApp.getActiveSpreadsheet();                                          // R
var sheet = ss.getSheetByName(sheetName);                                                // O
var eeUserName = Session.getEffectiveUser();                                             // R
var eeinfo ="Build Merge";                                                               // 
var eedetails ="Copy data from DATADUMP to MERGE adding formula";                        // L
var eeStatus="Started";                                                                  // O
var fName = arguments.callee.toString().match(/function ([^\(]+)/)[1]                    // G
var eeSheetName = "MERGE";                                                               //
var values = [                                                                           // S
[eeUserName,eeStatus, fName, eeinfo,eedetails, eeSheetName]                              // T
];                                                                                       // A
var range = sheet.getRange("B2:G2");                                                     // R
range.setValues(values);                                                                 // T
errorLogStart();                                                                         //
/////////////////////////////////////////////END ERROR LOG BLOCK - BEFORE CODE/////////////
//---------------------------------------------------------------------------------------//

//CODE BLOCK START
//start builing Merge Sheet from Datadump
//showMerge();
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sourceSheet = ss.getSheetByName('DATADUMP');
var endSheet = ss.getSheetByName('MERGE');

sourceSheet.getRange("A2:A").copyTo(endSheet.getRange("A3"), {contentsOnly:true});
sourceSheet.getRange("L2:L").copyTo(endSheet.getRange("B3"), {contentsOnly:true});
sourceSheet.getRange("D2:D").copyTo(endSheet.getRange("K3"), {contentsOnly:true});
sourceSheet.getRange("E2:E").copyTo(endSheet.getRange("L3"), {contentsOnly:true});
sourceSheet.getRange("C2:C").copyTo(endSheet.getRange("M3"), {contentsOnly:true});
sourceSheet.getRange("V2:V").copyTo(endSheet.getRange("N3"), {contentsOnly:true});
///////////////////////////////////////
//add formula
///////////////////////////////////////
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName('MERGE');

///////////////Corp Status
var values = [
//['=if("K"&row()="Recruit","NO APP",edate($M3,6))','=if(Q3="NO APP","NO APP",if(edate($M3,6)>edate(now(),0),Days($Q3,now()),"READY"))','=if (R3="NO APP","NO APP",if(R3="READY",+K3,"Days Left"))']
//];

['=if(K3="Recruit","NO APP",if(K3="","-",edate($M3,6)))','=if(Q3="NO APP","NO APP",if(edate($M3,6)>edate(now(),0),Days($Q3,now()),if(K3="","-","")))','=if (R3="NO APP","NO APP",if(R3="READY",+K3,if(K3="","-","Days Left")))']
];


var range = sheet.getRange("Q3:S3");
range.setValues(values);

///////////////Recruit Status
var values = [
['=if(M3="","-",edate(M3,1))','=if($K3="Recruit",if(J3="",if(edate($M3,1)>edate(Now(),0),"TRIAL 30 days","Time UP"),"* * UPDATE Member in Game"),+K3)','=if($V3="TRIAL 30 days",Days($U3,now()),)','=if($V3="Time UP",Days($U3,now()),)']
];

var range = sheet.getRange("U3:X3");
range.setValues(values);

///////////////AFK Status
var values = [
//['=if(N3="","",if(AA3="twelve","",if(AB3="2 Years","",if(edate($N3,6)>edate(Now(),0),"","Six"))))','=if(N3="","",if(AB3="2 Years","",if(edate($N3,12)>edate(Now(),0),"","twelve")))','=if(N3="","",if(edate($N3,24)>edate(Now(),0),"","2 Years"))','=if(AB3="2 Years", 1,"")','=if(N3="","",DAYS($N3-now(),))']
//];
['=if(N3="","",if(AA3=1,"",if(AB3=1,"",if(edate($N3,6)>edate(Now(),0),"",1))))','=if(N3="","",if(AB3=1,"",if(edate($N3,12)>edate(Now(),0),"",1)))','=if(N3="","",if(edate($N3,24)>edate(Now(),0),"",1))','=if(AD3="","",if(AD3<-31,"",1))','=if(N3="","",DAYS($N3-now(),))']
];
var range = sheet.getRange("Z3:AD3");
range.setValues(values);

//CODE BLOCK END

//---------------------------------------------------------------------------------------//
////////////////////////////////////////////////////END ERROR LOG - AFTER CODE///////////// E
var sheetName = 'ScriptLOG';                                                             // R
var ss = SpreadsheetApp.getActiveSpreadsheet();                                          // R
var sheet = ss.getSheetByName(sheetName);                                                // O
var eeUserName = Session.getEffectiveUser();                                             // R
var eeinfo ="Build Merge";                                                               // 
var eedetails ="Copy data from DATADUMP to MERGE adding formula";                        // L
var eeStatus="Ended";                                                                    // O
var fName = arguments.callee.toString().match(/function ([^\(]+)/)[1]                    // G
var eeSheetName = "MERGE";                                                               // 
var values = [                                                                           // E
[eeUserName,eeStatus, fName, eeinfo,eedetails, eeSheetName]                              // N
];                                                                                       // D
var range = sheet.getRange("B2:G2");                                                     //
range.setValues(values);                                                                 //
errorLogStart();                                                                         //
/////////////////////////////////////////////END ERROR LOG - BEFORE END////////////////////
//=======================================================================================//

//MAIN ERROR BLOCK END

}

function paintColor(){
//////////////////
//Colour Columns//
//////////////////

//=======================================================================================//
///////////////////////////////////////////////START ERROR LOG - BLOCK START/////////////// E
var sheetName = 'ScriptLOG';                                                             // R
var ss = SpreadsheetApp.getActiveSpreadsheet();                                          // R
var sheet = ss.getSheetByName(sheetName);                                                // O
var eeUserName = Session.getEffectiveUser();                                             // R
var eeinfo ="Paint";                                                                     // 
var eedetails ="Colours and borders";                                                    // L
var eeStatus="Started";                                                                  // O
var fName = arguments.callee.toString().match(/function ([^\(]+)/)[1]                    // G
var eeSheetName = "MERGE";                                                               //
var values = [                                                                           // S
[eeUserName,eeStatus, fName, eeinfo,eedetails, eeSheetName]                              // T
];                                                                                       // A
var range = sheet.getRange("B2:G2");                                                     // R
range.setValues(values);                                                                 // T
errorLogStart();                                                                         //
/////////////////////////////////////////////END ERROR LOG BLOCK - BEFORE CODE/////////////
//---------------------------------------------------------------------------------------//


//CODE BLOCK START
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName('MERGE');
//////////////////
/////////A, B
var cell = sheet.getRange("A:B");
cell.setFontColor("white");
cell.setBackground("black");
///////////F, I
var cell = sheet.getRange("F:I");
cell.setFontColor("black");
cell.setBackground("#f7ffb9");
///////////K, O
var cell = sheet.getRange("K:O");
cell.setFontColor("white");
cell.setBackground("black");
///////////Q, S
var cell = sheet.getRange("Q:S");
cell.setFontColor("#0000ff");
cell.setBackground("#c9daf8");
///////////U, X
var cell = sheet.getRange("U:X");
cell.setFontColor("#a61c00");
cell.setBackground("#d9ead3");
///////////Z
var cell = sheet.getRange("Z:Z");
cell.setFontColor("blue");
cell.setBackground("green");
///////////AA
var cell = sheet.getRange("AA:AA");
cell.setFontColor("blue");
cell.setBackground("#f1c232");
///////////AB AD
var cell = sheet.getRange("AB:AD");
cell.setFontColor("blue");
cell.setBackground("red");

var cell = sheet.getRange("AC:AC");
cell.setFontColor("#b299ff");
cell.setBackground("blue");
/////////////////
////Borders//////
/////////////////

//F-I - Black
var cell = sheet.getRange("F:I");
cell.setBorder(true,true,false,true,false,false,"black",SpreadsheetApp.BorderStyle.SOLID);
//Q-S - Blue
var cell = sheet.getRange("Q:S");
cell.setBorder(true,true,false,true,false,false,"blue",SpreadsheetApp.BorderStyle.SOLID);
//U-X - Green
var cell = sheet.getRange("U:X");
cell.setBorder(true,true,false,true,false,false,"green",SpreadsheetApp.BorderStyle.SOLID);
//Z-AD - Black dotted infilll
var cell = sheet.getRange("Z:AD");
cell.setBorder(true,true,true,false,true,false,"black",SpreadsheetApp.BorderStyle.DASHED);

//CODE BLOCK END

//---------------------------------------------------------------------------------------//
////////////////////////////////////////////////////END ERROR LOG - AFTER CODE///////////// E
var sheetName = 'ScriptLOG';                                                             // R
var ss = SpreadsheetApp.getActiveSpreadsheet();                                          // R
var sheet = ss.getSheetByName(sheetName);                                                // O
var eeUserName = Session.getEffectiveUser();                                             // R
var eeinfo ="Paint";                                                                     // 
var eedetails ="Colours and borders";                                                    // L
var eeStatus="Ended";                                                                    // O
var fName = arguments.callee.toString().match(/function ([^\(]+)/)[1]                    // G
var eeSheetName = "MERGE";                                                               //
var values = [                                                                           // E
[eeUserName,eeStatus, fName, eeinfo,eedetails, eeSheetName]                              // N
];                                                                                       // D
var range = sheet.getRange("B2:G2");                                                     //
range.setValues(values);                                                                 //
errorLogStart();                                                                         //
/////////////////////////////////////////////END ERROR LOG - BEFORE END////////////////////
//=======================================================================================//



//MAIN ERROR BLOCK END

///////////////////
//Colour Columns // END
///////////////////
}

/////////////////////////////////////////////////////
/////////Add mega List to MERGE - Step 7
/////////////////////////////////////////////////////

function addMegaList(){

//=======================================================================================//
///////////////////////////////////////////////START ERROR LOG - BLOCK START/////////////// E
var sheetName = 'ScriptLOG';                                                             // R
var ss = SpreadsheetApp.getActiveSpreadsheet();                                          // R
var sheet = ss.getSheetByName(sheetName);                                                // O
var eeUserName = Session.getEffectiveUser();                                             // R
var eeinfo ="Step 7";                                                                // 
var eedetails ="Add mega List to MERGE";                                                     // L
var eeStatus="Started";                                                                  // O
var fName = arguments.callee.toString().match(/function ([^\(]+)/)[1]                    // G
var eeSheetName = "MERGE";                                                               //
var values = [                                                                           // S
[eeUserName,eeStatus, fName, eeinfo,eedetails, eeSheetName]                              // T
];                                                                                       // A
var range = sheet.getRange("B2:G2");                                                     // R
range.setValues(values);                                                                 // T
errorLogStart();                                                                         //
/////////////////////////////////////////////END ERROR LOG BLOCK - BEFORE CODE/////////////
//---------------------------------------------------------------------------------------//

//CODE BLOCK START
megaListvlookup();
copyDataMerge();
//CODE BLOCK END

//---------------------------------------------------------------------------------------//
////////////////////////////////////////////////////END ERROR LOG - AFTER CODE///////////// E
var sheetName = 'ScriptLOG';                                                             // R
var ss = SpreadsheetApp.getActiveSpreadsheet();                                          // R
var sheet = ss.getSheetByName(sheetName);                                                // O
var eeUserName = Session.getEffectiveUser();                                             // R
var eeinfo ="Step 7";                                                                // 
var eedetails ="Add mega List to MERGE";                                                     // L
var eeStatus="Ended";                                                                    // O
var fName = arguments.callee.toString().match(/function ([^\(]+)/)[1]                    // G
var eeSheetName = "MERGE";                                                               //
var values = [                                                                           // E
[eeUserName,eeStatus, fName, eeinfo,eedetails, eeSheetName]                              // N
];                                                                                       // D
var range = sheet.getRange("B2:G2");                                                     //
range.setValues(values);                                                                 //
errorLogStart();                                                                         //
/////////////////////////////////////////////END ERROR LOG - BEFORE END////////////////////
//=======================================================================================//


//MAIN ERROR BLOCK END


//Browser.msgBox("Step 7 - addMegaList Finnished");
}

/////////////////////////////////////////////////////
/////////Add mega List to MERGE END
/////////////////////////////////////////////////////

function megaListvlookup(){

//=======================================================================================//
///////////////////////////////////////////////START ERROR LOG - BLOCK START/////////////// E
var sheetName = 'ScriptLOG';                                                             // R
var ss = SpreadsheetApp.getActiveSpreadsheet();                                          // R
var sheet = ss.getSheetByName(sheetName);                                                // O
var eeUserName = Session.getEffectiveUser();                                             // R
var eeinfo ="MEGA Vlookup";                                                              // 
var eedetails ="Copying Data from MEGALIST to MERGE";                                    // L
var eeStatus="Started";                                                                  // O
var fName = arguments.callee.toString().match(/function ([^\(]+)/)[1]                    // G
var eeSheetName = "MERGE";                                                               //
var values = [                                                                           // S
[eeUserName,eeStatus, fName, eeinfo,eedetails, eeSheetName]                              // T
];                                                                                       // A
var range = sheet.getRange("B2:G2");                                                     // R
range.setValues(values);                                                                 // T
errorLogStart();                                                                         //
/////////////////////////////////////////////END ERROR LOG BLOCK - BEFORE CODE/////////////
//---------------------------------------------------------------------------------------//

//CODE BLOCK START
//showMerge();
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName('MERGE');
////vlookup + error code

sheet.insertColumnBefore(3);
sheet.insertColumnBefore(5);
sheet.insertColumnBefore(7);


//var values = [
//["=vlookup($A2,'MEGALIST'!$A:$E,3,false)"]
//];

//var values = [
//  ["=vlookup($A3,'MEGALIST'!$A:$E,3,false)",'=iferror(C3,"")', "=vlookup($A3,'MEGALIST'!$A:$E,4,false)", '=iferror(E3,"")', "=vlookup($A3,'MEGALIST'!$A:$G,5,false)",'=iferror(G3,"")']
//];
//var range = sheet.getRange("C3:H3");
//range.setValues(values);

var values = [
  ["=vlookup($A3,'MEGALIST'!$A:$E,3,false)",'=iferror(C3,"")', "=vlookup($A3,'MEGALIST'!$A:$E,4,false)", '=iferror(E3,"")', "=vlookup($A3,'MEGALIST'!$A:$G,5,false)",'=iferror(G3,"")']
];
var range = sheet.getRange("C3:h3");
range.setValues(values);



//var values = [
//["=vlookup($A3,'MEGALIST'!$A:$E,3,false)"]
//];
//var range = sheet.getRange("C3");
//range.setValues(values);
//
//var values = [
//['=iferror(C3,"")']
//];
//var range = sheet.getRange("D3");
//range.setValues(values);
//////////////////////////////////////////////////////////
//var values = [
//["=vlookup($A3,'MEGALIST'!$A:$E,4,false)"]
//];
//var range = sheet.getRange("E3");
//range.setValues(values);
//var values = [
//['=iferror(E3,"")']
//];
//var range = sheet.getRange("F3");
//range.setValues(values);
////////////////////////////////////////////////
//var values = [
//["=vlookup($A3,'MEGALIST'!$A:$G,5,false)"]
//];
//var range = sheet.getRange("G3");
//range.setValues(values);
//var values = [
//['=iferror(G3,"")']
//];
//var range = sheet.getRange("H3");
//range.setValues(values);
/////////////////////////////////////////////////





sheet.getRange("C3:H3").copyTo(sheet.getRange("C3:H1000"), {contentsOnly:false});
//copy cade bar down
SpreadsheetApp.flush()
sheet.getRange("D3:D1000").copyTo(sheet.getRange("D3"), {contentsOnly:true});

sheet.getRange("F3:F1000").copyTo(sheet.getRange("F3"), {contentsOnly:true});

sheet.getRange("H3:H1000").copyTo(sheet.getRange("H3"), {contentsOnly:true});




//copy code values only

//sheet.getRange("d:d").copyTo(sheet.getRange("d:d"), {contentsOnly:true});


//sheet.getRange("F3:F1000").copyTo(sheet.getRange("F3:F1000"), {contentsOnly:true});

//sheet.getRange("H3:H1000").copyTo(sheet.getRange("H3:H1000"), {contentsOnly:true});

//sheet.getRange("d2:f2").copyTo(sheet.getRange("d3:f1000"), {contentsOnly:false});
//sheet.getRange("d2:f1000").copyTo(sheet.getRange("d2:f1000"), {contentsOnly:true});
//sheet.getRange("d2:d1000").copyTo(sheet.getRange("g2:g1000"), {contentsOnly:true});



//Convert code to values

//Browser.msgBox("AFTER ADD");
//////////////////////////////////////////////////////break;
sheet.deleteColumn(7);
sheet.deleteColumn(5);
sheet.deleteColumn(3);

//var range = sheet.getRange("E1:E1000");
//range.setWrap(true);


//CODE BLOCK END

//---------------------------------------------------------------------------------------//
////////////////////////////////////////////////////END ERROR LOG - AFTER CODE///////////// E
var sheetName = 'ScriptLOG';                                                             // R
var ss = SpreadsheetApp.getActiveSpreadsheet();                                          // R
var sheet = ss.getSheetByName(sheetName);                                                // O
var eeUserName = Session.getEffectiveUser();                                             // R
var eeinfo ="MEGA Vlookup";                                                              // 
var eedetails ="Copying Data from MEGALIST to MERGE";                                    // L
var eeStatus="Ended";                                                                    // O
var fName = arguments.callee.toString().match(/function ([^\(]+)/)[1]                    // G
var eeSheetName = "MERGE";                                                               //
var values = [                                                                           // E
[eeUserName,eeStatus, fName, eeinfo,eedetails, eeSheetName]                              // N
];                                                                                       // D
var range = sheet.getRange("B2:G2");                                                     //
range.setValues(values);                                                                 //
errorLogStart();                                                                         //
/////////////////////////////////////////////END ERROR LOG - BEFORE END////////////////////
//=======================================================================================//


//MAIN ERROR BLOCK END


}
/////////////////////////////////////////////////////
/////////BUILD MERGE SHEET END
/////////////////////////////////////////////////////

//copy DATA in MERGE
function copyDataMerge(){

//=======================================================================================//
///////////////////////////////////////////////START ERROR LOG - BLOCK START/////////////// E
var sheetName = 'ScriptLOG';                                                             // R
var ss = SpreadsheetApp.getActiveSpreadsheet();                                          // R
var sheet = ss.getSheetByName(sheetName);                                                // O
var eeUserName = Session.getEffectiveUser();                                             // R
var eeinfo ="MERGE";                                                                     // 
var eedetails ="Copy cells to populate sheet";                                           // L
var eeStatus="Started";                                                                  // O
var fName = arguments.callee.toString().match(/function ([^\(]+)/)[1]                    // G
var eeSheetName = "MERGE";                                                               //
var values = [                                                                           // S
[eeUserName,eeStatus, fName, eeinfo,eedetails, eeSheetName]                              // T
];                                                                                       // A
var range = sheet.getRange("B2:G2");                                                     // R
range.setValues(values);                                                                 // T
errorLogStart();                                                                         //
/////////////////////////////////////////////END ERROR LOG BLOCK - BEFORE CODE/////////////
//---------------------------------------------------------------------------------------//


//CODE BLOCK START
//showMerge();
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName('MERGE');
sheet.getRange("Q3:AD3").copyTo(sheet.getRange("Q3:Q1000"), {contentsOnly:false});
//CODE BLOCK END

//---------------------------------------------------------------------------------------//
////////////////////////////////////////////////////END ERROR LOG - AFTER CODE///////////// E
var sheetName = 'ScriptLOG';                                                             // R
var ss = SpreadsheetApp.getActiveSpreadsheet();                                          // R
var sheet = ss.getSheetByName(sheetName);                                                // O
var eeUserName = Session.getEffectiveUser();                                             // R
var eeinfo ="MERGE";                                                                     // 
var eedetails ="Copy cells to populate sheet";                                           // L
var eeStatus="Ended";                                                                    // O
var fName = arguments.callee.toString().match(/function ([^\(]+)/)[1]                    // G
var eeSheetName = "MERGE";                                                               // 
var values = [                                                                           // E
[eeUserName,eeStatus, fName, eeinfo,eedetails, eeSheetName]                              // N
];                                                                                       // D
var range = sheet.getRange("B2:G2");                                                     //
range.setValues(values);                                                                 //
errorLogStart();                                                                         //
/////////////////////////////////////////////END ERROR LOG - BEFORE END////////////////////
//=======================================================================================//


//MAIN ERROR BLOCK END
}
