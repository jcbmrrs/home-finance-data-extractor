var activesheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

function cleanMonthlyStatements() {
  folders = DriveApp.getFoldersByName("Monthly_Statements");
  while (folders.hasNext()) {
    var folder = folders.next();
    files = folder.getFiles();
    while (files.hasNext()) {
      var file = files.next();
      cleanStatements(file);
    }
  }
}

function ingestCleanStatements() {
  folders = DriveApp.getFoldersByName("Monthly_Statements");
  while (folders.hasNext()) {
    var folder = folders.next();
    files = folder.getFiles();
    while (files.hasNext()) {
      var file = files.next();
      updateFileName(file);
    }
  }
}

function ingestStatements(file) {
  var ss = SpreadsheetApp.open(file);
  Logger.log(activesheet.getName());
  Logger.log(ss.getName());
}

function cleanStatements(file) {
  
  var ss = SpreadsheetApp.open(file);
  var sheet = ss.getSheets()[0];
  var range = sheet.getRange(1,1);
  var values = range.getValues();
  //credit card statement
  if(values == "Posted Date") {
    sheet.deleteColumn(2);
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
    }
    sheet.deleteRow(1);
    var type = "alaska";
    updateFileName(file,type);
    //Logger.log(activesheet.getName());
  }
  else if(values == "Date") {
    Logger.log("date only!"); 
    sheet.insertColumnAfter(3);
    var range = sheet.getDataRange();
    
    for (i = 2; i <= range.getHeight(); i++) { 
      var newRange = sheet.getRange(i,4);
      var letter = String.fromCharCode(65 + width - 1);
      //var equation = "jacob";
      var equation = '=E' + i + '+F' + i;
      
      newRange.setValue(equation);
      
      var newValue = sheet.getRange(i,3).getValue().toString().replace("POS Withdrawal - ", "");
      newValue = newValue.replace(/ - Card Ending In ([0-9])\w+/g,"");
      newValue = newValue.replace("External Withdrawal - ","");
      Logger.log(newValue);
      sheet.getRange(i,3).setValue(newValue);
      Logger.log(newValue);
      
    }
    sheet.deleteRow(1);
    var type = "becu";
    updateFileName(file,type);
  }
  //
  //var range = sheet.getDataRange();
  //
  
  //Logger.log(values);
}

function updateFileName(file,type) {
  var ss = SpreadsheetApp.open(file);
  var sheet = ss.getSheets()[0];
  var labelDate = sheet.getRange(1,1).getValues();
  labelDate = new Date(labelDate);
  Logger.log(labelDate);
  labelDate = ((labelDate.getMonth())+1) + "-" + labelDate.getDate() + "-" + labelDate.getFullYear();
  if(type == "alaska") {
    var sheetName = "alaska_cc_" + labelDate + ".csv";
  } else if(type == "becu") {
    var sheetName = "becu_checking_" + labelDate + ".csv";
  }
  file.setName(sheetName);
}
