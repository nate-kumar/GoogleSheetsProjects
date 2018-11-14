// Global Variables
var ss = SpreadsheetApp.getActiveSpreadsheet();
var sheetIssues = ss.getSheetByName("Issues Tracker");
var calSystemIssues = CalendarApp.getCalendarsByName("TR Systems Engineering Weekly")[0];
var masterData = sheetIssues.getRange(7,2,sheetIssues.getLastRow()-7,14).getValues();

/*================================================================================================================================================*/

function weeklyActions() { 
  // MASTER FUNCTION - Generate actions calendar event for both engineers (Nathan Kumar + H___ G___)
  
  createCalendarEvent("NK")
  createCalendarEvent("HG")
  
}

function createActions(engineer) {
  // Identify actions (and corresponding data) per specific 'engineer' argument. Return as an array that is a subset of masterData
  
  var lenData = masterData.length
  
  var arrayActionsNK = [];  
  var arrayActionsHG = []; 
  
  // Separate actions into:
  // - NK Actions
  // - HG Actions
  // - Headings
  
  for (var a = 0; a < lenData; a++) {
    
    if (masterData[a][5] == "Action") {   
      if (masterData[a][7] == "NK" && (masterData[a][8] == "A" || masterData[a][8] == "R")) {
        arrayActionsNK.push([masterData[a][0],masterData[a][1],masterData[a][2],masterData[a][3],masterData[a][4]]);
      }
      else if (masterData[a][7] == "HG" && (masterData[a][8] == "A" || masterData[a][8] == "R")) {
        arrayActionsHG.push([masterData[a][0],masterData[a][1],masterData[a][2],masterData[a][3],masterData[a][4]])
      }
    }
    
  }
  
  // Return respective list of actions depending on selected 'engineer'
  
  if (engineer == "NK") {
    return actionsTable(arrayActionsNK);
  }
  else if (engineer == "HG") {
    return actionsTable(arrayActionsHG);
  }
  else {
  }
  
}

function actionsTable(arrayActions) {
  // Generate heading tree for each action, identify corresponding data for each heading ID by cross-comparing vs masterData, then remove duplicates. End result = all actions and all respective headings upto root
  
  var lenData = masterData.length
  
  var arrayHeadingTree = deriveHeadingTree(arrayActions);
  var headersFinal = [];
  Logger.log(arrayHeadingTree)
  
  for (var b = 0; b < lenData; b++) {
    for (var c = 0; c < arrayHeadingTree.length; c++) {
      
      var dataID = String(masterData[b][0]) + String(masterData[b][1]) + String(masterData[b][2]) + String(masterData[b][3]) + String(masterData[b][4]) 
      var headersID = String(arrayHeadingTree[c][0]) + String(arrayHeadingTree[c][1]) + String(arrayHeadingTree[c][2]) + String(arrayHeadingTree[c][3]) + String(arrayHeadingTree[c][4])
      if (dataID == headersID) {
        headersFinal.push(masterData[b])
      }
      
    }
  }
  var arrayFinal = removeDuplicates(headersFinal);
  return arrayFinal;
  
}

function deriveHeadingTree(data) { 
  // Calculates all heading IDs, but does not remove duplicates (e.g. [1,1,2,1,] -> [1,1,2,1,],[1,2,1,,],[2,1,,,],[1,,,,])
  
  var arrayHeadings = [];
  var lenHeader = data.length;
  
  for (var b = 0; b < lenHeader; b++) {
    
    for (var c = 4; c >= 0; c--) {
      
      if (data[b][c] == "") {
      }
      else {
        arrayHeadings.push([data[b][c-4]||"",data[b][c-3]||"",data[b][c-2]||"",data[b][c-1]||"",data[b][c]||""]);
      }
      
    }
    
  }
  
  var arrayHeadingsLength = arrayHeadings.length;
  var arrayHeadingsWidth = arrayHeadings[0].length;
  var arrayHeadingsOrder = []
  Logger.log(arrayHeadingsLength)

  for (var d = 0; d < arrayHeadingsLength; d++) {
    
    var arrayTemp = [];
    var counter = 0;
    
    for (var e = 0; e < arrayHeadingsWidth; e++) {

      if (arrayHeadings[d][e] == ""){
        counter += 1;
      }
      else {
        arrayTemp.push(arrayHeadings[d][e]);
      }
      
    }
    
    while (counter--){
      arrayTemp.push(""); 
    }
    arrayHeadingsOrder.push(arrayTemp);
    
  }
  return arrayHeadingsOrder;
  
}

function removeDuplicates(data) {
  // Scan array and remove any duplicates
  
  var filteredArray = data.filter(function (a) {
    if (!this[a]) {
      this[a] = true;
      return true;
    }
  }, Object.create(null));
  
  return filteredArray;
  
}

function createCalendarEvent(engineer) {
  // Generate 'description' from totalData by importing and formatting (input from createActions functions) and create Calendar event for relevant engineer
  
  var totalData = createActions(engineer)
  Logger.log(totalData)
  var dataLength = totalData.length;
  var dataWidth = totalData[0].length;
  var description = engineer + " Actions\n";
  
  for (var a = 0; a < dataLength; a++) {
    // Add content and format description, using ASCII \n for new-line
    
    if (totalData[a][1] == "") {
      description += "\n";
    }
    
    var b = 5
    while (b--) {
      if (totalData[a][b] != "") {
        for (var c = 0; c < b; c++) {
        // Add "- " per child to convey hierarchy
          
          description += "- "
          
        }
        
        if (totalData[a][6] != "") {
        // Add due date for all actions with a due date
          description += "[DUE: " + String(totalData[a][6]).slice(0,10) + "] "
        }
        
        description += String(totalData[a][9]) + String(totalData[a][10]) + String(totalData[a][11]) + String(totalData[a][12]) + String(totalData[a][13]);
        // Add heading or action data
        
        description += "\n"
        break;
      }
    }
    
  }
  Logger.log(description)
  
  var location = ""
  var today = new Date();  
  var startDate = addDays(today,(3 + 7 - today.getDay()) % 7)
  startDate.setHours(0);
  startDate.setMinutes(0);
  startDate.setSeconds(0);
  
  calSystemIssues.createEvent(
    // Create calendar event in calSystemIssues calendar
    engineer + " Actions", startDate, addDays(startDate,1), 
    {description: description,
      location: location, 
        
        sendInvites: false}
  )
  
}

function addDays(date, days) {
  return new Date(date.getTime() + (days * 24 * 60 * 60 * 1000));
}

/*================================================================================================================================================*/

function test() {
  
  Logger.log(createActions("HG"));
  
}
