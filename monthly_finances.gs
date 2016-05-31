function openMonthlyStatements() {
  folders = DriveApp.getFoldersByName("Monthly_Statements");
  while (folders.hasNext()) {
    var folder = folders.next();
    files = folder.getFiles();
    while (files.hasNext()) {
      var file = files.next();
      extractStatement(file);
    }
  }
}

function extractStatement(file) {
  var activesheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
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
  }
  //
  //var range = sheet.getDataRange();
  //
  
  //Logger.log(values);
}
