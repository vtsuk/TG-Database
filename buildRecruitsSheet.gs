  
//function buildRecruitsSheet(){
//addCodeToRecruitSheet();
//}

//Test function to add 'value to current cell
function addValueToActiveCell(){
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getActiveSheet();
var cell = sheet.getActiveCell();
cell.setValue('=if(I5="","..",if($G5="Recruit",if(edate($I5,1)>edate(Now(),0),"TRIAL 30 days","Time UP"),+$G5))');

}

/////////////////////////////////////////////////////
/////////MAKE makeRecruitSheet START
/////////////////////////////////////////////////////
//change sheet name when ready to deploy
/////////////////////////////////////////////////////

function makeRecruitSheet(){
var sheetname = 'Recruits'
var ss = SpreadsheetApp.getActiveSpreadsheet(), newSheet;
newSheet = ss.insertSheet();
newSheet.setName(sheetname);
}


function buildRecruitSheet(){

//text
//recruitText
recruitText();

//headers
//recruitHeader
recruitHeader();

//formula
//recruitFormula
recruitFormula();

//colour
//recruitColour
recruitColour();
//size
//recruitSize
recruitSize();

//date
//recruitDate
recruitDate();
//
}

//text
//recruitText
function recruitText(){

var sheetname = 'Recruits'
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName(sheetname);

//var cell = sheet.getRange("B2");
//cell.setValue('Press F5 to REFRESH SCREEN');

//var cell = sheet.getRange("C2");
//cell.setValue('Tactical Gamer Officer New recruits input Sheet');
//var cell = sheet.getRange("C3");
//cell.setValue('Enter details below for new recruits');
//var cell = sheet.getRange("C4");
//cell.setValue('Data admin will then merge these details into the mega list.');
}



//headers
//recruitHeader
function recruitHeader(){
var sheetname = 'Recruits'
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName(sheetname);

sheet.setFrozenRows(4);
sheet.setFrozenColumns(1);

//------------------------TOP LINE
var values = [
['=Concatenate("MEGALIST:    ",F1)', "", "", "", "", '=MEGALIST!A1']
];

var range = sheet.getRange("A1:F1");
range.setValues(values);

//--------------------------ROW2------
var values = [
['=Concatenate("MERGELIST: ",F2)', "", "", "", "", '=MERGE!A1']
];

var range = sheet.getRange("A2:F2");
range.setValues(values);
//--------------------------ROW3-------

var values = [
["", "", "", "", "X", "", "ON", "ON", "", "Action", "Trial", "Days", "Date", "X","ERROR TEST","ERROR TEST","ERROR TEST","ERROR TEST","ERROR TEST"]
];

var range = sheet.getRange("A3:S3");
range.setValues(values);

//--------------------------ROW4-------
//var values = [
//["In Game Name", "Forum Name", "Recruited by", "Extra Information", "", "Members Character Id", "MEGA LIST", "TGPS2APP", "Rank","Needed", "Status","Passed","Joined","","PS2APP","NAME","RANK","DATE","CELL A"]
//];
var values = [
["In Game Name", "Forum Name", "Recruited by", "Extra Information", "", "Members Character Id", "MERGE LIST", "TGPS2APP", "Rank","Needed", "Status","Passed","Joined","","PS2APP","NAME","RANK","DATE","CELL A"]
];

var range = sheet.getRange("A4:S4");
range.setValues(values);

}

//formula
//recruitFormula
function recruitFormula(){

//
var sheetname = 'Recruits'
//

var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName(sheetname);
//---------------------------row5
var values = [
//[".", '=A5','=if(S5="No DATA","..",if(R5="Not found","X","MERGE"))','=if(S5="No DATA","..",if(O5="Not found","X","APPS"))','=if(S5="No DATA","..",if(Q5="Not found","",Q5))','=if(S5="No DATA","",if(H5="APPS",if(G5="X","Invite",""),if(G5="MERGE",if(H5="X","WebAPP","L"),"No INFO")))','=if(S5="No DATA","",if($Q5="Recruit",if(edate($M5,1)>edate(Now(),0),"TRIAL 30 day","Time UP"),""))','=if(S5="No DATA","",if(I5="Recruit",-Days($M5,now()),""))','=if(S5="No DATA","",if(R5="Not found","",+R5))',".",'=if(S5="No DATA","...",iferror(vlookup($F5,TGPS2APP!D:J,1,false),"Not found"))','=if(S5="No DATA","...",iferror(vlookup($F5,MERGE!B:J,1,false),"Not found"))','=if(S5="No DATA","...",iferror(vlookup($F5,MERGE!B:K,10,false),"Not found"))','=if(S5="No DATA","...",iferror(vlookup($F5,MERGE!B:M,12,false),"Not found"))','=if($A5="","No DATA","READY")']
//];
[".", '=A5','=if(S5="No DATA","..",if(R5="Not found","NOT","TRUE"))','=if(S5="No DATA","..",if(O5="Not found","NOT","TRUE"))','=if(S5="No DATA","..",if(Q5="Not found","",Q5))','=if(S5="No DATA","",if(H5="TRUE",if(G5="NOT","Invite",""),if(G5="TRUE",if(H5="NOT","WebAPP","L"),"No INFO")))','=if(S5="No DATA","",if($Q5="Recruit",if(edate($M5,1)>edate(Now(),0),"TRIAL 30 day","Time UP"),""))','=if(S5="No DATA","",if(I5="Recruit",-Days($M5,now()),""))','=if(S5="No DATA","",if(R5="Not found","",+R5))',".",'=if(S5="No DATA","...",iferror(vlookup($F5,TGPS2APP!D:J,1,false),"Not found"))','=if(S5="No DATA","...",iferror(vlookup($F5,MERGE!B:J,1,false),"Not found"))','=if(S5="No DATA","...",iferror(vlookup($F5,MERGE!B:K,10,false),"Not found"))','=if(S5="No DATA","...",iferror(vlookup($F5,MERGE!B:M,12,false),"Not found"))','=if($A5="","No DATA","READY")']
];

var range = sheet.getRange("E5:S5");
range.setValues(values);


//var cell = sheet.getRange("S5");
//cell.setValue('=if($A5="","No DATA","READY")');
//var cell = sheet.getRange("R5");
//cell.setValue('=if(S5="No DATA","...",iferror(vlookup($F5,MEGALIST!B:M,12,false),"Not found"))');
//var cell = sheet.getRange("Q5");
//cell.setValue('=if(S5="No DATA","...",iferror(vlookup($F5,MEGALIST!B:K,10,false),"Not found"))');
//var cell = sheet.getRange("P5");
//cell.setValue('=if(S5="No DATA","...",iferror(vlookup($F5,MEGALIST!B:J,1,false),"Not found"))');
//var cell = sheet.getRange("O5");
//cell.setValue('=if(S5="No DATA","...",iferror(vlookup($F5,TGPS2APP!D:J,1,false),"Not found"))');
////var cell = sheet.getRange("N5");
////cell.setValue('');
//var cell = sheet.getRange("M5");
//cell.setValue('=if(S5="No DATA","",if(R5="Not found","",+R5))');
//var cell = sheet.getRange("L5");
//cell.setValue('=if(S5="No DATA","",if(I5="Recruit",-Days($M5,now()),""))');
//var cell = sheet.getRange("K5");
//cell.setValue('=if(S5="No DATA","",if($Q5="Recruit",if(edate($M5,1)>edate(Now(),0),"TRIAL 30 day","Time UP"),""))');
//var cell = sheet.getRange("J5");
//cell.setValue('=if(S5="No DATA","",if(H5="APPS",if(G5="X","Invite",""),if(G5="MEGA",if(H5="X","WebAPP","L"),"No INFO")))');
//var cell = sheet.getRange("I5");
//cell.setValue('=if(S5="No DATA","..",if(Q5="Not found","",Q5))');
//var cell = sheet.getRange("H5");
//cell.setValue('=if(S5="No DATA","..",if(O5="Not found","X","APPS"))');
//var cell = sheet.getRange("G5");
//cell.setValue('=if(S5="No DATA","..",if(R5="Not found","X","MEGA"))');
//var cell = sheet.getRange("F5");
//cell.setValue('=A5');

sheet.getRange("E5:S5").copyTo(sheet.getRange("E6:S200"), {contentsOnly:false});

}


//colour
//recruitColour
function recruitColour(){
var sheetname = 'Recruits'
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName(sheetname);


//Recruit Section

var cell = sheet.getRange("A:D");
cell.setFontColor("black");
cell.setBackground("#d9ead3"); //green
//cell.setFontWeight("bold");

var cell = sheet.getRange("A1:D4");
cell.setFontColor("#d9ead3");
cell.setBackground("black"); 
cell.setFontWeight("bold");

var cell = sheet.getRange("E:E");
cell.setFontColor("black");
cell.setBackground("Black");

//Status

var cell = sheet.getRange("F:M");
cell.setFontColor("#0000ff");
cell.setBackground("#cfe2f3"); 
//cell.setFontWeight("bold");

var cell = sheet.getRange("F1:M4");
cell.setFontColor("#cfe2f3");
cell.setBackground("#0000ff"); 
cell.setFontWeight("bold");


var cell = sheet.getRange("N:N");
cell.setFontColor("black");
cell.setBackground("Black");

//ERROR TEST

var cell = sheet.getRange("O:S");
cell.setFontColor("#ffa500");
cell.setBackground("#ff0000"); 
//cell.setFontWeight("bold");

var cell = sheet.getRange("O1:S4");
cell.setFontColor("#ff0000");
cell.setBackground("#ffa500"); 
//cell.setFontWeight("bold");
}



//size
//recruitSize
function recruitSize(){
var sheetname = 'Recruits'
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName(sheetname);

//sheet.autoResizeColumn(1);
sheet.setColumnWidth(1,150);
sheet.autoResizeColumn(2);
sheet.autoResizeColumn(3);
//sheet.autoResizeColumn(4);
sheet.setColumnWidth(4,110);
sheet.autoResizeColumn(5);
//sheet.autoResizeColumn(6);
sheet.setColumnWidth(6,150);
sheet.autoResizeColumn(7);
sheet.autoResizeColumn(8);
sheet.autoResizeColumn(9);
sheet.autoResizeColumn(10);
sheet.autoResizeColumn(11);
sheet.autoResizeColumn(12);
//sheet.autoResizeColumn(13);
sheet.setColumnWidth(13,80);
sheet.autoResizeColumn(14);
sheet.autoResizeColumn(15);
sheet.autoResizeColumn(16);
sheet.autoResizeColumn(17);
sheet.autoResizeColumn(18);
sheet.autoResizeColumn(19);


}



//date
//recruitDate
function recruitDate(){

  var sheetname = 'Recruits'
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheet = ss.getSheetByName(sheetname);


var range = sheet.getRange("M:M");
range.setNumberFormat("YYYY-MM-DD");
//var range = sheet.getRange("M:M");
//range.setNumberFormat("YYYY-MM-DD");
//var range = sheet.getRange("H:H");
//range.setNumberFormat("YYYY-MM-DD");
}
