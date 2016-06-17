//activesheet is the current spreadhseet (in this case it's the 2016 Joint Finances)
var activesheet = SpreadsheetApp.getActiveSpreadsheet();
var folders = DriveApp.getFoldersByName("Monthly_Statements");

//for every file in the Monthly_Statements folder, run the cleanStatements function
function cleanMonthlyStatements() {
  while (folders.hasNext()) {
    var folder = folders.next();
    files = folder.getFiles();
    while (files.hasNext()) {
      var file = files.next();
      cleanStatements(file);
    }
  }
}

//for each file that has been cleaned, ingest the data into the 2016 Joint Finances file
function ingestCleanStatements() {
  while (folders.hasNext()) {
    var folder = folders.next();
    files = folder.getFiles();
    while (files.hasNext()) {
      var file = files.next();
      ingestStatements(file);
    }
  }
}

function ingestStatements(file) {
  var ss = SpreadsheetApp.open(file);
  var sheetNames = ["becu", "alaska"];
  sheetNames.forEach(deleteDupes);
  sheetNames.forEach(moveData);
  
  function deleteDupes(item, index) {
    //if the name of the sheet matches the name of the relevant file
    if(ss.getSheetByName(item)) {
      Logger.log(ss.getName());
      //select the specific named sheet
      var destinationSheet = activesheet.getSheetByName(item);
      var dheight = destinationSheet.getDataRange().getHeight();
      var dwidth = destinationSheet.getDataRange().getWidth();
      //ss is the current workbook with data to be copied into 2016 Joint Finances
      var sheet = ss.getSheetByName(item);
      var range = sheet.getDataRange();
      
      Logger.log(range.getValues());
      var width = range.getWidth();
      
      //the current sheet's dates
      var dateCheckRange = sheet.getRange(1,1,range.getHeight(), 1);
      // the destination sheet's dates
      var destDateRange = destinationSheet.getRange(1,1,dheight,1);
      //Logger.log(destDateRange.getValues());
      
      //Logger.log(dateCheckRange); 
      //(row, column, numRows, numColumns);
      
      //for each date value in the current sheet's date column
      for (var j = 1; j <= dateCheckRange.getHeight(); j++) {
        Logger.log("j:" + j);
        //sd is the date string for comparision
        var sd = new Date(dateCheckRange.getValues()[j]).toDateString();
        
        //for each date value in the destination sheet's date column, starting with the third row (after header rows)
        for (var i = 1; i <= (destDateRange.getHeight()); i++) { 
          //Logger.log("i:" + i);
          //Logger.log(destDateRange.getValues()[i]);
          //dd is the date string for comparision
          var dd = new Date(destDateRange.getValues()[i]).toDateString();
          //Logger.log(dd, sd);
          
          //the string dates match AND the values of the desination sheet starting in the Number/Reference Column going through the final column before category
          if(dd == sd) {// && destinationSheet.getRange(i,3,1,dwidth-5).getValues().toString() == sheet.getRange(j,3,1,width-4).getValues().toString()) {
            Logger.log("exact match found" + dd + " || " + sd);// and deleted");
            Logger.log(destinationSheet.getRange(i,1,1,dwidth-3).getValues());
            Logger.log(sheet.getRange(j,1,1,width-1).getValues());
            //sheet.deleteRow(j);
          }
        }
      }
    }
  }
  
  function moveData(item, index) {
    if(ss.getSheetByName(item)) {
      Logger.log(ss.getName());
      Logger.log(ss.getSheetName());
      var mvdestinationSheet = activesheet.getSheetByName(item);
      Logger.log(mvdestinationSheet.getName());
      var mvdheight = mvdestinationSheet.getDataRange().getHeight();
      var mvdwidth = mvdestinationSheet.getDataRange().getWidth();
      var mvsheet = ss.getSheetByName(item);
      var mvrange = mvsheet.getDataRange();
      var mvwidth = mvrange.getWidth();
      var mvheight = mvrange.getHeight();
      Logger.log(mvdheight);
      var destCopy = mvdestinationSheet.getRange(mvdheight+1,1,mvheight,mvwidth);
      Logger.log(mvrange.getValues());
      //destCopy.setValues(mvrange.getValues());
      //mvrange.clearContent();
    }
  }
  /*
  var sheet = ss.getSheetByName("becu");
  var range = sheet.getRange(1,1);
  var values = range.getValues();
  //credit card statement
  if(values == "Posted Date") {
    sheet.deleteColumn(2);
    var range = sheet.getDataRange();
    var width = range.getWidth();
  
  
  
  Logger.log(activesheet.getName());
  Logger.log(ss.getName());
  //([date, game1players, game1teams, game2players, game2teams, pay.toString(), locale]);*/
}

function cleanStatements(file) {
  
  var ss = SpreadsheetApp.open(file);
  var sheet = ss.getSheets()[0];
  var range = sheet.getRange(1,1);
  var values = range.getValues();
  
  //credit card statement
  if(values == "Posted Date") {
    sheet.insertColumnAfter(1);
    var range = sheet.getDataRange();
    var width = range.getWidth();
    for (i = 2; i <= range.getHeight(); i++) { 
      var newRange = sheet.getRange(i,width+1);
      var letter = String.fromCharCode(65 + width - 1);
      var equation = '=IF(' + letter + i + '<0,' + letter + i + ',"")';
      newRange.setValue(equation);
      
      var newRange2 = sheet.getRange(i,width+2);
      var letter2 = String.fromCharCode(65 + width);
      var equation2 = '=IF(' + letter + i + '>0,' + letter + i + ',"")';
      newRange2.setValue(equation2);
      
      var newRange3 = sheet.getRange(i,2);
      var equation3 = '=MONTH(A' + i + ')';
      newRange3.setValue(equation3).setNumberFormat("0");
      
      var newValue = sheet.getRange(i,5).getValue().toString().replace(/\s+/g,' ').trim();
      Logger.log(newValue);
      sheet.getRange(i,5).setValue(newValue);
      Logger.log(newValue);
      
      var newValue = sheet.getRange(i,4).getValue().toString().replace(/\s+/g,' ').trim();
      Logger.log(newValue);
      sheet.getRange(i,4).setValue(newValue);
      Logger.log(newValue);
    }
    range.sort({column: 1, ascending: false});
    sheet.deleteRow(1);
    var type = "alaska";
    sheet.setName(type);
    updateFileName(file,type);
  }
  //becu statement
  else if(values == "Date") {
    Logger.log("date only!");
    sheet.insertColumnAfter(3);
    sheet.insertColumnAfter(1);
    var range = sheet.getDataRange();
    
    for (i = 2; i <= range.getHeight(); i++) { 
      var newRange = sheet.getRange(i,5);
      var letter = String.fromCharCode(65 + width - 1);
      var equation = '=F' + i + '+G' + i;
      newRange.setValue(equation);
      
      var newRange3 = sheet.getRange(i,2);
      var equation3 = '=MONTH(A' + i + ')';
      newRange3.setValue(equation3).setNumberFormat("0");
      
      var newValue = sheet.getRange(i,4).getValue().toString().replace("POS Withdrawal - ", "");
      newValue = newValue.replace(/ - Card Ending In ([0-9])\w+/g,"");
      newValue = newValue.replace("External Withdrawal - ","");
      newValue = newValue.replace(/\s+/g,' ').trim();
      Logger.log(newValue);
      sheet.getRange(i,4).setValue(newValue);
      Logger.log(newValue);
      
    }
    range.sort({column: 1, ascending: false});
    sheet.deleteRow(1);
    var type = "becu";
    sheet.setName(type);
    updateFileName(file,type);
  }
}

function updateFileName(file,type) {
  var ss = SpreadsheetApp.open(file);
  var sheet = ss.getSheets()[0];
  var labelDate = sheet.getRange(1,1).getValues();
  labelDate = new Date(labelDate);
  labelDate = ((labelDate.getMonth())+1) + "-" + labelDate.getDate() + "-" + labelDate.getFullYear();
  if(type == "alaska") {
    var sheetName = "alaska_cc_" + labelDate + ".csv";
  } else if(type == "becu") {
    var sheetName = "becu_checking_" + labelDate + ".csv";
  }
  file.setName(sheetName);
}
