var activesheet = SpreadsheetApp.getActiveSpreadsheet();
var folders = DriveApp.getFoldersByName("Monthly_Statements");

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
    if(ss.getSheetByName(item)) {
      Logger.log(ss.getName());
      var destinationSheet = activesheet.getSheetByName(item);
      var dheight = destinationSheet.getDataRange().getHeight();
      var dwidth = destinationSheet.getDataRange().getWidth();
      var sheet = ss.getSheetByName(item);
      var range = sheet.getDataRange();
      
      Logger.log(range.getValues());
      var width = range.getWidth();
      var dateCheckRange = sheet.getRange(1,1,range.getHeight(), 1).getValues();
      var destDateRange = destinationSheet.getRange(3,1,dheight,1);
      //Logger.log(destDateRange.getValues());
      
      //Logger.log(dateCheckRange); (row, column, numRows, numColumns)
      //Logger.log("break");
      //Logger.log(destDateRange.getValues());
      
      for (var j = 0; j <  dateCheckRange.length; j++) {
        var sd = new Date(dateCheckRange[j]).toDateString();
        //Logger.log(dateCheckRange.length)
        
        for (var i = 1; i <= destDateRange.getHeight(); i++) { 
          //Logger.log(destDateRange.getValues()[i]);
          var dd = new Date(destDateRange.getValues()[i]).toDateString();
          //Logger.log(dd, sd);
          
          if(dd == sd && destinationSheet.getRange(i+3,2,1,dwidth-2).getValues().toString() == sheet.getRange(j+1,2,1,width-1).getValues().toString()) {
            Logger.log("exact match found and deleted");
            Logger.log(destinationSheet.getRange(i+3,2,1,dwidth-2).getValues());
            Logger.log(sheet.getRange(j+1,2,1,width-1).getValues());
            sheet.deleteRow(j+1);
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
      destCopy.setValues(mvrange.getValues());
      mvrange.clearContent();
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
  
  /*var activeRange = sheet.getActiveRange();
  // iterate through all cells in the selected range
  for (var cellRow = 1; cellRow <= activeRange.getHeight(); cellRow++) {
    for (var cellColumn = 1; cellColumn <= activeRange.getWidth(); cellColumn++) {
      cell = activeRange.getCell(cellRow, cellColumn);
      cellValue = cell.getValue();
      cell.setValue(String(cellValue).replace(/\s+/g,'').trim());
      Logger.log("string replaced!");
    }
  }*/
  
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
      
      var newValue = sheet.getRange(i,3).getValue().toString().replace(/\s+/g,' ').trim();
      Logger.log(newValue);
      sheet.getRange(i,3).setValue(newValue);
      Logger.log(newValue);
      
      var newValue = sheet.getRange(i,2).getValue().toString().replace(/\s+/g,' ').trim();
      Logger.log(newValue);
      sheet.getRange(i,2).setValue(newValue);
      Logger.log(newValue);
    }
    sheet.deleteRow(1);
    var type = "alaska";
    sheet.setName(type);
    updateFileName(file,type);
  }
  //becu statement
  else if(values == "Date") {
    Logger.log("date only!"); 
    sheet.insertColumnAfter(3);
    var range = sheet.getDataRange();
    
    for (i = 2; i <= range.getHeight(); i++) { 
      var newRange = sheet.getRange(i,4);
      var letter = String.fromCharCode(65 + width - 1);
      var equation = '=E' + i + '+F' + i;
      
      newRange.setValue(equation);
      
      var newValue = sheet.getRange(i,3).getValue().toString().replace("POS Withdrawal - ", "");
      newValue = newValue.replace(/ - Card Ending In ([0-9])\w+/g,"");
      newValue = newValue.replace("External Withdrawal - ","");
      newValue = newValue.replace(/\s+/g,' ').trim();
      Logger.log(newValue);
      sheet.getRange(i,3).setValue(newValue);
      Logger.log(newValue);
      
    }
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
