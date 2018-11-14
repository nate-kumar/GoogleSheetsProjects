var ss = SpreadsheetApp.getActiveSpreadsheet();

var sheetMain = ss.getSheetByName("Main");
var sheetClosedIssues = ss.getSheetByName("Closed Issues");
var sheetNewIssues = ss.getSheetByName("New Issues");
var sheetRAGChanges = ss.getSheetByName("RAG Changes");
var sheetStatusUpdates = ss.getSheetByName("Status Updates");
var sheetCompiledResults = ss.getSheetByName("Compiled Results");

function runIssuesTrackerCompare() {
  
  // Manual upload of both .xls documents is required before running this function  (any folder EXCEPT ROOT in Drive is acceptable)
  
  masterFunctionSequence("2016_11_02","2016_10_19");
  
}


function masterFunctionSequence(dateRecent,date2WeeksAgo) {
  
  // Delete all sheets that are not Main, Closed Issues, New Issues, Status Updates, RAG Changes, Compiled Results
  
  initialiseRoutine();
  
  // Converts file1.xls (Excel) and file2.xls (Excel) to file1.xls (Gsheet) and file2.xls (Gsheet)
  
  testConvertExcel2Sheets('TR JDC Issues Tracker_'+dateRecent+'.xls');
  testConvertExcel2Sheets('TR JDC Issues Tracker_'+date2WeeksAgo+'.xls');
  
  copySheetsToMasterSpreadsheet(dateRecent,date2WeeksAgo);
  findDeltasFromSSCompare(dateRecent,date2WeeksAgo);
  compileResults(dateRecent,date2WeeksAgo);
  
  savePDF("Compiled Results",dateRecent,date2WeeksAgo);
  emailPDF(dateRecent,date2WeeksAgo);
  
  cleanUpDrive();
  
}


function initialiseRoutine() {
  
  // Clear data in Main, New Issues, Closed Issues, RAG Changes and Status Updates
  
  clearSubsheets()
  
  // Clear all .xls files from Root Folder 
  
  cleanUpDrive()
  
  var arrayDeleteSheets = [];
  
  // Identify all values of i in which getSheets()[i] != Main / Compiled Results / Closed Issues / New Issues / RAG Changes / Status Updates
  
  for (var i = 0, len = ss.getSheets().length; i < len; i++) {
    
    if (ss.getSheets()[i].getName() != "Compiled Results" && ss.getSheets()[i].getName() != "Closed Issues" && ss.getSheets()[i].getName() != "New Issues" && ss.getSheets()[i].getName() != "RAG Changes" && ss.getSheets()[i].getName() != "Status Updates" && ss.getSheets()[i].getName() != "Main") {
      arrayDeleteSheets.push(i);
    }
    
  }
  
  // Delete all sheets with index i
  
  Logger.log(arrayDeleteSheets)
  for (var j = 0, len1 = arrayDeleteSheets.length; j < len1; j++) {
    ss.deleteSheet(ss.getSheets()[arrayDeleteSheets[j]-j]);
  }
  
  // Clear data and formatting in sheetCompiledResults 
  
  sheetCompiledResults.getRange(2,9).setValue("a");
  sheetCompiledResults.getRange(1,1,sheetCompiledResults.getLastRow(),sheetCompiledResults.getLastColumn()).breakApart().setValue("").setFontSize(10).setFontWeight("normal").setHorizontalAlignment("center").setBackground("white").setBorder(true,true,true,true,true,true,"white",null);
  
}


function convertExcel2Sheets(excelFile, filename, arrParents) {
  
  var parents  = arrParents || []; // check if optional arrParents argument was provided, default to empty array if not
  if ( !parents.isArray ) parents = []; // make sure parents is an array, reset to empty array if not
  
  // Parameters for Drive API Simple Upload request (see https://developers.google.com/drive/web/manage-uploads#simple)
  var uploadParams = {
    method:'post',
    contentType: 'application/vnd.ms-excel', // works for both .xls and .xlsx files
    contentLength: excelFile.getBytes().length,
    headers: {'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()},
    payload: excelFile.getBytes()
  };
  
  // Upload file to Drive root folder and convert to Sheets
  var uploadResponse = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v2/files/?uploadType=media&convert=true', uploadParams);
  
  // Parse upload&convert response data (need this to be able to get id of converted sheet)
  var fileDataResponse = JSON.parse(uploadResponse.getContentText());
  
  // Create payload (body) data for updating converted file's name and parent folder(s)
  var payloadData = {
    title: filename, 
    parents: []
  };
  if ( parents.length ) { // Add provided parent folder(s) id(s) to payloadData, if any
    for ( var i=0; i<parents.length; i++ ) {
      try {
        var folder = DriveApp.getFolderById(parents[i]); // check that this folder id exists in drive and user can write to it
        payloadData.parents.push({id: parents[i]});
      }
      catch(e){} // fail silently if no such folder id exists in Drive
    }
  }
  // Parameters for Drive API File Update request (see https://developers.google.com/drive/v2/reference/files/update)
  var updateParams = {
    method:'put',
    headers: {'Authorization': 'Bearer ' + ScriptApp.getOAuthToken()},
    contentType: 'application/json',
    payload: JSON.stringify(payloadData)
  };
  
  // Update metadata (filename and parent folder(s)) of converted sheet
  UrlFetchApp.fetch('https://www.googleapis.com/drive/v2/files/'+fileDataResponse.id, updateParams);
  
  return SpreadsheetApp.openById(fileDataResponse.id);
}


function testConvertExcel2Sheets(spreadsheetName) {
  
  var files = DriveApp.getFilesByName(spreadsheetName);
  
  var xlsFile = DriveApp.getFileById(files.next().getId());// get the file object
  var xlsBlob = xlsFile.getBlob(); // Blob source of Excel file for conversion
  var xlsFilename = "xls" + xlsFile.getName(); // File name to give to converted file; defaults to same as source file
  var destFolders = []; // array of IDs of Drive folders to put converted file in; empty array = root folder
  var ss = convertExcel2Sheets(xlsBlob, xlsFilename, destFolders);
  Logger.log(ss.getId());
  
}


function moveFilesToFolder() {
  
  var driveRoot = DriveApp.getRootFolder();
  var driveIssuesTracker = DriveApp.getFoldersByName("Issues Tracker Script")
  
}


function copySheetsToMasterSpreadsheet(dateRecent,date2WeeksAgo) {
  
  var source1 = SpreadsheetApp.openById(DriveApp.getFilesByName('xlsTR JDC Issues Tracker_'+dateRecent+'.xls').next().getId());
  var source2 = SpreadsheetApp.openById(DriveApp.getFilesByName('xlsTR JDC Issues Tracker_'+date2WeeksAgo+'.xls').next().getId());
  
  var sheet1 = source1.getSheets()[0];
  var sheet2 = source2.getSheets()[0];
  
  var destination = SpreadsheetApp.getActiveSpreadsheet();
  
  sheet1.copyTo(destination).setName(dateRecent);
  sheet2.copyTo(destination).setName(date2WeeksAgo);
  
}


function clearSubsheets() {
  
  var lastColumn = sheetMain.getLastColumn() || 9;
  var lastRowNewIssues = sheetNewIssues.getLastRow() || 1;
  var lastRowClosedIssues = sheetClosedIssues.getLastRow() || 1;
  var lastRowRAGChanges = sheetRAGChanges.getLastRow() || 1;
  var lastRowStatusUpdates = sheetStatusUpdates.getLastRow() || 1;
  
  var arraySheets = [sheetNewIssues,lastRowNewIssues,sheetClosedIssues,lastRowClosedIssues,sheetRAGChanges,lastRowRAGChanges,sheetStatusUpdates,lastRowStatusUpdates];
  
  for (var i = 0, len = arraySheets.length/2; i < len; i++) {
    arraySheets[2*i].getRange(1,10).setValue("a");    
    arraySheets[2*i].getRange(1,1,arraySheets[(2*i)+1],lastColumn).setValue("").setFontWeight("normal").setFontSize(10);
  }
  
}


function findDeltasFromSSCompare(dateRecent,date2WeeksAgo) {
  
  var sheetDateRecent = ss.getSheetByName(dateRecent);
  var sheetDate2WeeksAgo = ss.getSheetByName(date2WeeksAgo);
  
  if (sheetDateRecent.getLastRow() >= sheetDate2WeeksAgo.getLastRow()) {
    var rangeLastRow = sheetDateRecent.getLastRow()
    }  
  else if (sheetDateRecent.getLastRow() < sheetDate2WeeksAgo.getLastRow()) {
    var rangeLastRow = sheetDate2WeeksAgo.getLastRow()
    }
  
  if (sheetDateRecent.getLastColumn() >= sheetDate2WeeksAgo.getLastColumn()) {
    var rangeLastColumn = sheetDateRecent.getLastColumn()
    }  
  else if (sheetDateRecent.getLastColumn() < sheetDate2WeeksAgo.getLastColumn()) {
    var rangeLastColumn = sheetDate2WeeksAgo.getLastColumn()
    }
  
  var sheetDRValues = sheetDateRecent.getRange(1,1,rangeLastRow,rangeLastColumn).getValues();
  var sheetD2WValues = sheetDate2WeeksAgo.getRange(1,1,rangeLastRow,rangeLastColumn).getValues();
  
  var countMain = 0;
  var countClosedIssues = 0;
  var countNewIssues = 0;
  var countRAGChanges = 0;
  var countStatusUpdates = 0;
  var valueLasti = 0;
  
  for (var i = 0, len = sheetDRValues.length; i < len; i++) {
    for (var j = 0, len1 = sheetDRValues[0].length; j < len1; j++) {
      if (sheetDRValues[i][j] != sheetD2WValues[i][j]) {
        if (valueLasti < i) {
          valueLasti = i
          countMain += 1;
          sheetMain.getRange(countMain,2,1,rangeLastColumn).setValues(sheetDateRecent.getRange(i+1,1,1,rangeLastColumn).getValues());
          sheetMain.getRange(countMain,1).setValue(i+1);
          sheetMain.getRange(countMain,j+2).setFontWeight("bold").setFontSize(11);
          if (j == 0 && sheetDRValues[i][0] != "" && sheetD2WValues[i][0] == "") {
            countNewIssues += 1;
            sheetNewIssues.getRange(countNewIssues,2,1,rangeLastColumn).setValues(sheetDateRecent.getRange(i+1,1,1,rangeLastColumn).getValues());
            sheetNewIssues.getRange(countNewIssues,1).setValue(i+1);
            sheetNewIssues.getRange(countNewIssues,2).setFontWeight("bold").setFontSize(11);
          }
          else if (j == 2 && sheetMain.getRange(countMain,j+2).getValue() == "G") {
            countClosedIssues += 1;
            sheetClosedIssues.getRange(countClosedIssues,2,1,rangeLastColumn).setValues(sheetDateRecent.getRange(i+1,1,1,rangeLastColumn).getValues());
            sheetClosedIssues.getRange(countClosedIssues,4).setValue(sheetDate2WeeksAgo.getRange(i+1,3).getValue()+" -> "+sheetDateRecent.getRange(i+1,3).getValue()).setFontWeight("bold").setFontSize(11);
            sheetClosedIssues.getRange(countClosedIssues,1).setValue(i+1);
          }
          else if (j == 2 && sheetDRValues[i][2] != "G" && sheetDRValues[i][2] != sheetD2WValues[i][2]) {
            countRAGChanges += 1;
            sheetRAGChanges.getRange(countRAGChanges,2,1,rangeLastColumn).setValues(sheetDateRecent.getRange(i+1,1,1,rangeLastColumn).getValues())
            sheetRAGChanges.getRange(countRAGChanges,4).setValue(sheetDate2WeeksAgo.getRange(i+1,3).getValue()+" -> "+sheetDateRecent.getRange(i+1,3).getValue()).setFontWeight("bold").setFontSize(11);
            sheetRAGChanges.getRange(countRAGChanges,1).setValue(i+1);
          }
          else {
            countStatusUpdates += 1;
            sheetStatusUpdates.getRange(countStatusUpdates,2,1,rangeLastColumn).setValues(sheetDateRecent.getRange(i+1,1,1,rangeLastColumn).getValues());
            sheetStatusUpdates.getRange(countStatusUpdates,1).setValue(i+1);
            sheetStatusUpdates.getRange(countStatusUpdates,j+2).setFontWeight("bold").setFontSize(11);
          }
        }
        else if (valueLasti == i) {
          Logger.log("i: " + (i+1) + " j: " + j)
          sheetMain.getRange(countMain,j+2).setFontWeight("bold");
          if (sheetDRValues[i][0] != "" && sheetD2WValues[i][0] == "") {
            sheetNewIssues.getRange(countNewIssues,j+2).setFontWeight("bold").setFontSize(11);
          }
          else if (sheetMain.getRange(countMain,4).getValue() == "G") {
            sheetClosedIssues.getRange(countClosedIssues,j+2).setFontWeight("bold").setFontSize(11);
          }
          else if (sheetD2WValues[i][0] != "" && (sheetDRValues[i][2] != sheetD2WValues[i][2]) && (sheetDRValues[i][j] != sheetD2WValues[i][j])) {
            sheetRAGChanges.getRange(countRAGChanges,j+2).setFontWeight("bold").setFontSize(11);
          }
          else {
            sheetStatusUpdates.getRange(countStatusUpdates,j+2).setFontWeight("bold").setFontSize(11);
          }
        }
      }
    }
  }
  
}


function testSendEmail() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheetMain = ss.getSheetByName("Main");
  
  var body = '<div style="font-size:10">'
  body += '<H1 style="font-size:8">'+ 'testText' +'</H1>';
  body += '<H2>'
  "test"
  + '</H2>';
  body += '</div>';
  
  var recipient = 'XXXXX@XXXXX.com';  // For debugging, send only to self
  var subject = "testSubject"
  MailApp.sendEmail(recipient, subject, "", {htmlBody:body})
  
  
}

function savePDF(sheetPDF,dateRecent,date2WeeksAgo) {
  SpreadsheetApp.flush();
  
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetPDF);
  
  var url = ss.getUrl();
  
  //remove the trailing 'edit' from the url
  url = url.replace(/edit$/,'');
  
  //additional parameters for exporting the sheet as a pdf
  var url_ext = 'export?exportFormat=pdf&format=pdf' + //export as pdf
    //below parameters are optional...
    '&size=letter' + //paper size
      '&portrait=true' + //orientation, false for landscape
        '&fitw=true' + //fit to width, false for actual size
          '&sheetnames=false&printtitle=false&pagenumbers=false' + //hide optional headers and footers
            '&gridlines=false' + //hide gridlines
              '&fzr=false' + //do not repeat row headers (frozen rows) on each page
                '&gid=' + sheet.getSheetId(); //the sheet's Id
  
  var token = ScriptApp.getOAuthToken();
  
  var response = UrlFetchApp.fetch(url + url_ext, {
                                   headers: {
                                   'Authorization': 'Bearer ' +  token
                                   }
                                   });
  
  var blob = response.getBlob().setName('Driver Modes Issues Meeting'  + ' (' + dateModifier(dateRecent) + ' - ' + dateModifier(date2WeeksAgo) + ').pdf');
  
  
  //OR DocsList.createFile(blob);
  DriveApp.createFile(blob)
  
}

function testSavePDF() {
  
  savePDF("Compiled Results");
  
}

function compileResults(dateRecent,date2WeeksAgo) {
  
  var sheetDateRecent = ss.getSheetByName(dateRecent);
  var sheetDate2WeeksAgo = ss.getSheetByName(date2WeeksAgo);
  
  var lastColumn = sheetMain.getLastColumn() || 9;
  var lastRowNewIssues = sheetNewIssues.getLastRow() || 1;
  var lastRowClosedIssues = sheetClosedIssues.getLastRow() || 1;
  var lastRowRAGChanges = sheetRAGChanges.getLastRow() || 1;
  var lastRowStatusUpdates = sheetStatusUpdates.getLastRow() || 1;
  
  sheetCompiledResults.getRange(2,9).setValue("a");
  sheetCompiledResults.getRange(1,1,sheetCompiledResults.getLastRow(),sheetCompiledResults.getLastColumn()).breakApart().setValue("").setFontSize(10).setFontWeight("normal").setHorizontalAlignment("center").setBackground("white").setBorder(true,true,true,true,true,true,"white",null);
  
  var arraySheets = [sheetNewIssues,lastRowNewIssues,sheetClosedIssues,lastRowClosedIssues,sheetRAGChanges,lastRowRAGChanges,sheetStatusUpdates,lastRowStatusUpdates];
  Logger.log(arraySheets);
  
  var lastRowCompiledResults = 2;
  
  var todaysDate = new Date()
  
  sheetCompiledResults.getRange(1,1,1,5).merge().setFontSize(18).setFontWeight("bold").setValue("Driver Modes Issues Meeting - Change Report").setHorizontalAlignment("left");
  sheetCompiledResults.getRange(2,1,1,5).merge().setFontSize(14).setFontStyle("italic").setValue(dateModifier(date2WeeksAgo) + "  to  " + dateModifier(dateRecent)).setHorizontalAlignment("left");
  sheetCompiledResults.setRowHeight(3,50);
  sheetCompiledResults.getRange(1,9).setFontSize(14).setFontWeight("bold").setValue(todaysDate);
  sheetCompiledResults.getRange(1,8).setFontSize(14).setFontWeight("bold").setValue("Created on: ").setHorizontalAlignment("right");
  
  for (var i = 0, len = (arraySheets.length)/2; i < len; i++) {
    sheetCompiledResults.getRange(lastRowCompiledResults + 2,1,1,lastColumn).merge().setHorizontalAlignment("left").setFontSize(14).setFontWeight("Bold").setBackground("#6d9eeb").setBorder(true, false, true, false, false, false, null, null);
    sheetCompiledResults.getRange(lastRowCompiledResults + 2,1).setValue(arraySheets[2*i].getName())
    sheetCompiledResults.getRange(lastRowCompiledResults + 3,2,1,lastColumn-1).setValues(sheetDateRecent.getRange(2,1,1,lastColumn-1).getValues()).setBackground("#a4c2f4").setVerticalAlignment("middle");
    sheetCompiledResults.getRange(lastRowCompiledResults + 3,1,1,lastColumn).setBorder(true,false, true,false,false,false,null,null);
    sheetCompiledResults.getRange(lastRowCompiledResults + 3,1).setValue("Row ID").setBackground("#a4c2f4");
    sheetCompiledResults.getRange(lastRowCompiledResults + 4,1,arraySheets[(2*i)+1],lastColumn).setValues(arraySheets[2*i].getRange(1,1,arraySheets[(2*i)+1],lastColumn).getValues());
    
    arraySheets[2*i].getRange(1,1,arraySheets[(2*i)+1],lastColumn).copyFormatToRange(sheetCompiledResults,1,lastColumn,lastRowCompiledResults + 4,lastRowCompiledResults + arraySheets[(2*i)+1] + 3);
    
    sheetCompiledResults.getRange(lastRowCompiledResults + 4,1,arraySheets[(2*i)+1],lastColumn).setBorder(false,false,false,false,false,true,"#cccccc",null)
    
    lastRowCompiledResults = sheetCompiledResults.getLastRow();
    
    sheetCompiledResults.getRange(lastRowCompiledResults+1,1,1,lastColumn).setBorder(true,false,true,false,false,false,null,null)
    sheetCompiledResults.setRowHeight(lastRowCompiledResults+1,50);
    
  } 
  Logger.log(lastRowCompiledResults);
  sheetCompiledResults.getRange(lastRowCompiledResults+2,1,1,lastColumn).setBorder(true,false,false,false,false,false,"white",null);
  
}


function cleanUpDrive() {
  
  var driveRoot = DriveApp.getRootFolder().getFiles();
  while (driveRoot.hasNext()) {
    var file = driveRoot.next();
    if (file.getName().substring(0,3) == "xls") {
      Logger.log(file.setTrashed(true));
    }
  } 
  
}


function emailPDF(dateRecent,date2WeeksAgo) {
  
  var driveRoot = DriveApp.getRootFolder().getFiles();
  
  var emailAddress = "XXXXX@XXXXX.com";
  
  var subject = "Driver Modes Issues Meeting Minutes - " + dateModifier(dateRecent);
  var message = "testMessage to Nathan";
  
  var dateRecentMod = "";
  
  var fileCount = 0;
  var arrayFileId = [];
  var arrayFileDate = [];
  
  while (driveRoot.hasNext()) {
    
    var file = driveRoot.next()
    
    if (file.getName() == 'Driver Modes Issues Meeting'  + ' (' + dateModifier(dateRecent) + ' - ' + dateModifier(date2WeeksAgo) + ').pdf') {
      fileCount += 1;
      arrayFileId.push(file.getId());
      arrayFileDate.push(file.getDateCreated().getTime());
    }
    
    if (fileCount > 1) {
      for (var i = 0, len = arrayFileDate.length; i < len; i++) {
        if (arrayFileDate[i] == Math.max.apply(null, arrayFileDate)) {
          var dateLatestPDF = arrayFileDate[i];
          var fileCorrectPDFId = arrayFileId[i];
        }
      }
    }
    else if (fileCount == 0) {
      return;
    }
    else {
      var fileCorrectPDFId = arrayFileId[0];
    }
    
  } 
  
  var fileCorrectPDF = DriveApp.getFileById(fileCorrectPDFId);
  
  var files = DriveApp.getFilesByName('TR JDC Issues Tracker_'+dateRecent+'.xls');
  
  var xlsFile = DriveApp.getFileById(files.next().getId());// get the file object
  
  //MailApp.sendEmail(emailAddress, subject, message, {attachments: fileCorrectPDF});
  MailApp.sendEmail({
    to: emailAddress,
    subject: subject,
    htmlBody: "Hi all,<br><br>"
    + "Please find attached the change report from the latest Driver Modes Issues Meeting (<b>" 
    + dateModifier(dateRecent) + "</b>).<br><br>"
    + "<i>Note: Deltas relative to the previous meeting (<b>" + dateModifier(date2WeeksAgo)
    + "</b>) are highlighted in <b>bold</b>. This affects the whole cell, not just the text that has changed.</i><br><br>"
    + "---- This report contains actions for both 19/10/16 and 05/10/16 ----<br><br>"
    + "Kind Regards,",
    attachments: [fileCorrectPDF, xlsFile]
  });
  
}


function test() {
  
  emailPDF("2016_03_09","2016_02_24");
  
}

function dateModifier(x) {
  
  return (x.substring(8,10) + "/" + x.substring(5,7) + "/" + x.substring(2,4));
  
}
/*===========TO DO====================
@Format copying to Compiled Results is not currently working
@Set subtitles under heading
@Need to automate input of copySheetsToMasterSpreadsheet(dateRecent,date2WeeksAgo)
@add borders around cells prior to being PDFd
@Remove files from Drive as necessary
@Add title to PDF
Test emailing to self
Email trigger

@sheetDateRecent and sheetDate2WeeksAgo are currently hardcoded

*/
