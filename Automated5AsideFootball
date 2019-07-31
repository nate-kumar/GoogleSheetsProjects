/*~~~~~~~~~~~~~~~~~~~~~~~~~~~~STRUCTURE~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~
- GLOBAL VARIABLES
- TEST FUNCTION

1 - [Chase Matt Email]
2 - EMAIL: REMIND UNANSWERED - Wed 10:00

3 - DOWNLOAD PLAYERS FROM CALENDAR - Wed 14:00
4 - COPY THIS WEEKS PLAYERS - Wed 14:00
5 - TEAM BALANCER - Wed 14:00
6 - PASTE BALANCED TEAMS INTO ATTENDANCE SHEET - Wed 14:00
7 - UPDATE PLAYER STATUS IN CALENDARRESPONSES - Wed 14:00

8A - EMAIL: THIS WEEKS TEAMS - Wed 14:00
8B - CANCEL GAME AND SPLIT PAYMENT - Wed 14:00

9 - PARTNERSHIPS - Thu - 00:00

10 - EMAIL: PAYMENT CHASER - Thu 10:00
11 - CALENDAR INVITE - Thu 10:00
~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~~*/

//==GLOBAL VARIABLES====================================================================================================================//

var ss = SpreadsheetApp.getActiveSpreadsheet()

var sheetBalance = ss.getSheetByName("Player List");
var sheetPlayers = ss.getSheetByName("Player List");
var season = 2;

var sheetAttendanceAllTime = ss.getSheetByName("Results (All-time)");
var sheetPartnershipsAllTime = ss.getSheetByName("Partnerships (All-time)");
if (season == 1) {
  var sheetAttendance = ss.getSheetByName("Results (Season 1)");
  var sheetPartnerships = ss.getSheetByName("Partnerships (Season 1)");
  var sheetBalancer = ss.getSheetByName("Team Balancer (OLD)");
}
else if (season == 2) {
  var sheetAttendance = ss.getSheetByName("Results (Season 2)");
  var sheetPartnerships = ss.getSheetByName("Partnerships (Season 2)");
  var sheetBalancer = ss.getSheetByName("Team Balancer");
}

var sheetCalendar = ss.getSheetByName("Calendar Responses");

var calFootball = CalendarApp.getCalendarsByName("Wednesday Football")[0];

var playerList = sheetPlayers.getRange(5,3,sheetPlayers.getLastRow()-5,1).getValues();
var playerListEmails = sheetPlayers.getRange(5,4,sheetPlayers.getLastRow()-5,1).getValues();
var playerListStatus = sheetCalendar.getRange(4,1,sheetCalendar.getLastRow()-4,3).getValues();
function getCalendarGuestList() {
  var week = sheetBalancer.getRange(4,3).getValue();                
  var today = new Date();
  var day = addDays(today,(3 + 7 - today.getDay()) % 7)
  var event = calFootball.getEventsForDay(day, {search: 'Social Football'})[0];
  if (event != null) {
    return event.getGuestList();
  }
}
var guestlist = getCalendarGuestList();

var weekNumberString = sheetBalancer.getRange(2,2).getValue();      
var weekNumber = parseInt(weekNumberString.slice(15,weekNumberString.length));

if (season == 1) {
  var weekNumberGlobal = weekNumber;
}
else if (season == 2) {
  var weekNumberGlobal = weekNumber+26;
}

//==TEST FUNCTION====================================================================================================================//

function test() {
  
  for (var i = 0; i < playerList.length; i++) {
//    playerStatus(playerList[i],30);
    Logger.log("abc");
  } 
  
}


//==EMAIL: REMIND UNANSWERED====================================================================================================================//

function remindUnanswered() {
  
  var guestlist = getCalendarGuestList();
  var unansweredEmail = [];
  for (var i = 0; i < guestlist.length; i++) {
    if (guestlist[i].getGuestStatus() == "MAYBE" || guestlist[i].getGuestStatus() == "INVITED") {
      unansweredEmail.push(guestlist[i].getEmail());
    }
  }
  Logger.log(unansweredEmail)
  var emailAddress = ""   // EMAIL ADDRESS BLANKED
  for (var j = 0; j < unansweredEmail.length; j++) {
    emailAddress = emailAddress + "," + unansweredEmail[j];
  }
  
  var subject = "Social Football - are you playing tonight?";
  
  MailApp.sendEmail({
    to: emailAddress,
    subject: subject,
    htmlBody: "Hi all,<br><br>"
    + "Just a friendly reminder to please update your response for tonight's Social Football game by 2pm today. The teams will be auto-generated and sent out at this time."
    + "<br><br>"
    + "Any responses left at Maybe/Unanswered will be treated as Not Attending as from 2pm."
    + "<br><br>"
    + "<i>Note: If you are <b>not</b> planning on playing regularly and want to avoid these emails let me know and I can change your rank to stand-in</i>"
    + "<br><br>"
    + "Thanks"
  }); 
  
}

//==WED 14:00 WRAPPER====================================================================================================================//

function wrapperWed2pm() {
  
  sheetBalancer.getRange(4,3).setValue(weekNumber)
  var numPlayers = downloadPlayersFromCalendar()
  if (numPlayers >= 10) {
    copyThisWeeksPlayers();
    teamBalancerMaster();
    pasteBalancedTeams();
    thisWeekTeamsEmail();
  }
  else {
    emailNotEnoughPlayers();
  }
  updateAllStatus();
  
}

//==DOWNLOAD PLAYERS FROM CALENDAR====================================================================================================================//

function downloadPlayersFromCalendar() {
  
  var week = sheetBalancer.getRange(4,3).getValue();         
  sheetAttendance.getRange(14,16+week,sheetAttendance.getLastRow()-14,1).setValue("")
  var attendeesEmail = [];
  var attendeesName = [];
  var counter = 0;
  for (var i = 0; i < guestlist.length; i++) {
    if (guestlist[i].getGuestStatus() == "YES") {
      attendeesEmail.push(guestlist[i].getEmail());
    }
  }
  for (var j = 0; j < attendeesEmail.length; j++) {
    for (var k = 0; k < playerListEmails.length; k++) {
      if (playerListEmails[k][0] == attendeesEmail[j]) {
        counter+=1;
        sheetAttendance.getRange(14+k,16+week).setValue("t");
      }
    }
  }
  return counter;
}

//==COPY THIS WEEKS PLAYERS====================================================================================================================//

function copyThisWeeksPlayers() {
  
  sheetBalancer.getRange(5,5,20,2).setValue("");         
  
  var arrayPlayers = [];
  var arrayPlayersRatings = [];
  var arrayPlayersUnfilt = sheetBalancer.getRange(5,2,20,1).getValues();    
  var arrayPlayers = filterArray(arrayPlayersUnfilt);
  
  var numPlayers = arrayPlayers.length;
  for (var i = 0; i < playerList.length; i++) {
    for (var j = 0; j < numPlayers; j++) {
      if (arrayPlayers[j][0] == playerList[i][0]) {
        arrayPlayersRatings.push([playerList[i][0],playerListStatus[i][0]])
      }   
    }
  }
  
  if (numPlayers < 10) {
    Logger.log("email cancel")
    return
  }
  
  shuffleArray(arrayPlayersRatings);
  
  arrayPlayersRatings.sort(function(a, b) {
    return a[1] - b[1];
  });
  
  var arrayCount = [0,0,0,0,0];
  for (var k = 0; k < arrayPlayersRatings.length; k++) {
    //Logger.log(arrayPlayersRatings[k][1])
    if (arrayPlayersRatings[k][1] == 1) { arrayCount[0]+=1; }
    else if (arrayPlayersRatings[k][1] == 2) { arrayCount[1]+=1; }
    else if (arrayPlayersRatings[k][1] == 3) { arrayCount[2]+=1; }
    else if (arrayPlayersRatings[k][1] == 4) { arrayCount[3]+=1; }
    else if (arrayPlayersRatings[k][1] == 5) { arrayCount[4]+=1; }
  }
  Logger.log(arrayCount)
  
  var arraySplitPlayersRatings = arrayPlayersRatings;
  var arrayPriorityBin = playerListPriority(arrayCount);
  //Logger.log(arrayPriorityBin);
  var arrayDefoPlaying = arraySplitPlayersRatings.splice(0,arrayPriorityBin[0]);
  var arrayMaybePlaying = shuffleArray(arraySplitPlayersRatings);
  
  var arrayFinalPlayers = [];
  var arrayNotPlaying = [];
  
  for (var l = 0; l < arrayPriorityBin[0]; l++) {
    arrayFinalPlayers.push(arrayDefoPlaying[l])
  }
  for (var m = 0; m < arrayMaybePlaying.length; m++) {
    if (m < arrayPriorityBin[1]) {
      arrayFinalPlayers.push(arrayMaybePlaying[m])
    }
    else {
      arrayNotPlaying.push(arrayMaybePlaying[m])
    }
  }
  
  for (var n = 0; n < arrayFinalPlayers.length; n++) {
    sheetBalancer.getRange(5+n,5).setValue(arrayFinalPlayers[n][0]);
    sheetBalancer.getRange(5+n,6).setValue(arrayFinalPlayers[n][1]);  
  }
  for (var o = 0; o < arrayNotPlaying.length; o++) {
    sheetBalancer.getRange(18+o,5).setValue(arrayNotPlaying[o][0]);
    sheetBalancer.getRange(18+o,6).setValue(arrayNotPlaying[o][1]);
  }
  
}

function playerListPriority(arrayCount) {
  
  var numPlayers = 0;
  for (var a = 0; a < arrayCount.length; a++) { numPlayers+=arrayCount[a] }
  var gameNumPlayers = 12;
  if (numPlayers >= 10) {
    if (numPlayers > 12) {
      gameNumPlayers = 12;
    }
    else if (numPlayers == 12) {
      gameNumPlayers = 12;
    }
    else if (numPlayers == 11) {
      gameNumPlayers = 10;
    }
    else if (numPlayers == 10) {
      gameNumPlayers = 10;
    }
  }
  else {
    Logger.log("Cancel game");
  }
  Logger.log(gameNumPlayers)
  Logger.log(arrayCount.length);
  
  var numPlayersRemain = numPlayers;
  for (var i = 0; i < arrayCount.length; i++) {
    if (arrayCount[i] < numPlayersRemain) {
      Logger.log([arrayCount[i],numPlayersRemain])
      numPlayersRemain-=arrayCount[i]
      Logger.log(numPlayersRemain)
    }
    else if (arrayCount[i] == numPlayersRemain) {
      numPlayersRemain-=arrayCount[i]
      Logger.log([gameNumPlayers-numPlayersRemain,numPlayersRemain,i])
      return [gameNumPlayers-numPlayersRemain,numPlayersRemain,i]
    }
    else {
      Logger.log([gameNumPlayers-numPlayersRemain,numPlayersRemain,i])
      return [gameNumPlayers-numPlayersRemain,numPlayersRemain,i]
    }
  }
  
}


//==TEAM BALANCER====================================================================================================================//

function teamBalancerMaster() {
  var gameFormat = sheetBalancer.getRange("H2").getValue()
  if (gameFormat == "4-aside") {
    balanceTeams(arrayPlayerLetters4,arrayComb4);
  }
  else if (gameFormat == "5-aside") {
    balanceTeams(arrayPlayerLetters5,arrayComb5);
  }
  else if (gameFormat == "6-aside") {
    balanceTeams(arrayPlayerLetters6,arrayComb6);
  }
}

function balanceTeams(arrayPlayerLetters, arrayComb) {
  
  Logger.log(arrayPlayerLetters)
  Logger.log(arrayComb)
  //var ssPerformance = SpreadsheetApp.openById("1oidOEJRx1-Bh2q94QqlHOPELBwOSDl3M6sp6vJdHvSE").getSheetByName("Temp Balancer").getRange(2,2).getValue()
  var portPlayerList = sheetBalancer.getRange(5,5,12,1).getValues();
  SpreadsheetApp.openById("1oidOEJRx1-Bh2q94QqlHOPELBwOSDl3M6sp6vJdHvSE").getSheetByName("Temp Balancer").getRange(2,2,12,1).setValues(portPlayerList);
  
  if (season == 1) {
    var dataPlayers = sheetBalancer.getRange(5,5,sheetBalancer.getLastRow(),2).getValues();
  }
  else if (season == 2) {
    var dataPlayers = SpreadsheetApp.openById("1oidOEJRx1-Bh2q94QqlHOPELBwOSDl3M6sp6vJdHvSE").getSheetByName("Temp Balancer").getRange(2,2,12,2).getValues();
  }
  var dataPlayersClean = [];
  
  sheetBalancer.getRange(5,8,6,1).setValue("");
  if (season == 1) {
    sheetBalancer.getRange(5,11,sheetBalancer.getLastRow()-5,1).setValue("");
  }
  else if (season == 2) {
    sheetBalancer.getRange(5,10,6,1).setValue("");
  }
  
  for (var y = 0; y < dataPlayers.length; y++) {
    if (dataPlayers[y][0] != "") {
      dataPlayersClean.push(dataPlayers[y]);
    }
  }
  dataPlayersClean = shuffleArray(dataPlayersClean);
  
  var arrayPlayers = [];
  var sumPlayers = 0;
  
  for (var z = 0; z < arrayPlayerLetters.length; z++) {
    arrayPlayers.push([arrayPlayerLetters[z],dataPlayersClean[z][0],dataPlayersClean[z][1]]);
  }
  
  for (var a = 0; a < arrayPlayers.length; a++) {
    sumPlayers = sumPlayers + arrayPlayers[a][2];
  }
  var averagePlayers = sumPlayers/2
  
  var arrayScores = [];
  var sum = 0;
  
  for (var b = 0; b < arrayComb.length; b++) {
    sum = 0
    for (var c = 0; c < arrayComb[0].length; c++) {
      for (var d = 0; d < arrayPlayers.length; d++) {
        if (arrayPlayers[d][0] == arrayComb[b][c])
          sum = sum + arrayPlayers[d][2];
      }
    }
    arrayScores.push([Math.abs(sum-averagePlayers),arrayComb[b]])
  }
  
  arrayScores.sort(function(a, b) {
    return a[0] - b[0];
  });
  
  Logger.log(arrayScores)
  var teamALetters = arrayScores[0][1];
  var teamBLetters = arrayPlayerLetters.filter( function( el ) {
    return teamALetters.indexOf( el ) < 0;
  });
  
  var teamA = []
  var teamB = []
  var counterA = 0
  var counterB = 0
  
  for (var e = 0; e < teamALetters.length; e++) {
    for (var f = 0; f < arrayPlayers.length; f++) { 
      if (teamALetters[e] == arrayPlayers[f][0]) {
        teamA.push(arrayPlayers[f][1]);
        counterA += 1;
        sheetBalancer.getRange(4+counterA,8,1,1).setValue(arrayPlayers[f][1]);
      }
      else if (teamBLetters[e] == arrayPlayers[f][0]) {
        teamB.push(arrayPlayers[f][1]);
        counterB += 1;
        if (season == 1) {
          sheetBalancer.getRange(4+counterB,11,1,1).setValue(arrayPlayers[f][1]);
        }
        else if (season == 2) {
          sheetBalancer.getRange(4+counterB,10,1,1).setValue(arrayPlayers[f][1]);
        }
      }
    }
  }
  if (sheetBalancer.getLastRow()-17>0) {
  var benchedPlayers = sheetBalancer.getRange(18,5,sheetBalancer.getLastRow()-17,2).getValues();
  }
  else {
  benchedPlayers = 0
  }
    var arrayPlayingPriority = sheetBalancer.getRange(5,6,12,1).getValues();
  var lowestPriority = 0
  for (var g = 0; g < arrayPlayingPriority.length; g++) {
    if (arrayPlayingPriority[g][0] > lowestPriority) {
      lowestPriority = arrayPlayingPriority[g][0]
    }
  }
  Logger.log(benchedPlayers);
  for (var f = 0; f < benchedPlayers.length; f++) {
    if (benchedPlayers[f][1] == lowestPriority) {
      Logger.log("Randomised")
      sheetBalancer.getRange(13+f,8).setValue(benchedPlayers[f][0]);
      sheetBalancer.getRange(13+f,10).setValue("Equal Priority (Removed at random)");
    }
    else if (benchedPlayers[f][1] == 2) {
      sheetBalancer.getRange(13+f,8).setValue(benchedPlayers[f][0]);
      sheetBalancer.getRange(13+f,10).setValue("Lower Priority (Back-up)");
      Logger.log("Backup (Lower Priority)")
    }
    else {
      Logger.log("Payment")
      sheetBalancer.getRange(13+f,8).setValue(benchedPlayers[f][0]);
      sheetBalancer.getRange(13+f,10).setValue("Lower Priority (Payment Outstanding)");
    }
  }
  
}

//==PASTE BALANCED TEAMS INTO ATTENDANCE SHEET====================================================================================================================//

function pasteBalancedTeams() {
  var length = parseInt(sheetBalancer.getRange(2,8).getValue().charAt(0));
  var week = sheetBalancer.getRange(4,3).getValue();
  var team1 = getBalancedTeams(1,length);
  var team2 = getBalancedTeams(2,length);
  var masterList = playerList;
  
  sheetAttendance.getRange(14,16+week,sheetAttendance.getLastColumn()-14,1).setValue("")
  for (var i = 0; i < masterList.length; i++) {
    for (var j = 0; j < length; j++) {
      if (masterList[i][0] == team1[j][0]) {
        sheetAttendance.getRange(14+i,16+week).setValue("T1");
      }
      if (masterList[i][0] == team2[j][0]) {
        sheetAttendance.getRange(14+i,16+week).setValue("T2");
      }
    }
  }
  
}

function getBalancedTeams(team,size) {
  if (season == 1) {
    return sheetBalancer.getRange(5,8+(team-1)*3,size,1).getValues() ;
  }
  else if (season == 2) {
    return sheetBalancer.getRange(5,8+(team-1)*2,size,1).getValues() ;
  }
}

//==UPDATE PLAYER STATUS IN CALENDARRESPONSES====================================================================================================================//

function updateAllStatus() {
  
  //var weekNumberString = sheetBalancerOld.getRange(2,2).getValue();
  //var weekNumber = parseInt(weekNumberString.slice(15,weekNumberString.length));
  //var weekNumber = 20;
  respondedCalendar(weekNumber);
  for (var i = 0; i < playerList.length; i++) {
    playerStatus(playerList[i],weekNumberGlobal);
  } 
  
}

function playerStatus(playerName,weekNumberGlobal) {
  
  for (var i = 0; i < playerList.length; i++) {
    if (playerName == playerList[i][0]) {
      var playerID = i+1;
    }
  }
  var attendance = sheetAttendanceAllTime.getRange(13+playerID,7).getValue();
  var currentPlayerStatus = sheetCalendar.getRange(3+playerID,2).getValue();
  var played5wks = playedLast5Weeks(playerID,weekNumberGlobal);
  var calendar3wks = respondedLast3Calendar(playerID,weekNumberGlobal);
  if (currentPlayerStatus != "S") {
    Logger.log([played5wks,attendance,calendar3wks])
    if ((played5wks >= 2 || attendance >= 0.4) && (calendar3wks >= 2)) {
      Logger.log(playerName + " Regular")
      sheetCalendar.getRange(3+playerID,2).setValue("R")
    }
    else if ((calendar3wks >= 1)) {
      Logger.log(playerName + " Backup")
      sheetCalendar.getRange(3+playerID,2).setValue("B")
    }
    else {
      Logger.log(playerName + " Standin")
      sheetCalendar.getRange(3+playerID,2).setValue("S")
    }
  }
  else {
    Logger.log(playerName + " Standin")
  }
  
}

function playedLast5Weeks(playerID, weekNumber) {
  
  var playerRecord = sheetAttendanceAllTime.getRange(13+playerID,17+weekNumber-5,1,5).getValues();
  //Logger.log(playerRecord);
  var counter = 0;
  for (var j = 0; j < playerRecord[0].length; j++) {
    //Logger.log(playerRecord[0][j])
    if (playerRecord[0][j] == "T1" || playerRecord[0][j] == "T2") {
      counter+=1;
    }
    else {}
  }
  return counter;
  //Logger.log(playerList)
  
}

function respondedLast3Calendar(playerID,weekNumber) {
  
  var responses = sheetCalendar.getRange(3+playerID,5+weekNumber-3,1,3).getValues()
  //Logger.log(responses)
  var counter = 0;
  for (var i = 0; i < responses[0].length; i++) {
    //Logger.log(responses[0][i])
    if (responses[0][i] != "M" && responses[0][i] != "U") {
      counter+=1;
    }
  }
  return counter;
  
}

function respondedCalendar(weekNumber) {
  
  var startDate = sheetAttendance.getRange(3,17).getValue();
  var startYear = startDate.getFullYear();
  var startMonth = startDate.getMonth();
  var startDay = startDate.getDate();
  var newStartDate = new Date(Date.UTC(startYear, startMonth, startDay, 0, 0, 0))
  Logger.log(startDate)
  var day = addDays(newStartDate,((weekNumber-1)*7))
  
  Logger.log(day)
  var event = calFootball.getEventsForDay(day, {search: 'Social Football'})[0];
  var guestlist = event.getGuestList();
  var todaysStatusList = sheetCalendar.getRange(4,4+weekNumberGlobal,sheetCalendar.getLastRow()-4,1).getValues()
  
  
  for (var i = 0; i < guestlist.length; i++) {
    for (var j = 0; j < playerListEmails.length; j++) {
      
      if (guestlist[i].getEmail() == playerListEmails[j][0] && (todaysStatusList[j] != "New" && todaysStatusList[j] != "H")) {
        if (guestlist[i].getGuestStatus() == "YES") {
          sheetCalendar.getRange(4+j,4+weekNumberGlobal).setValue("A")
        }
        else if (guestlist[i].getGuestStatus() == "NO") {
          sheetCalendar.getRange(4+j,4+weekNumberGlobal).setValue("N")
        }
        else if (guestlist[i].getGuestStatus() == "MAYBE") {
          sheetCalendar.getRange(4+j,4+weekNumberGlobal).setValue("M")
        }
        else if (guestlist[i].getGuestStatus() == "INVITED") {
          sheetCalendar.getRange(4+j,4+weekNumberGlobal).setValue("U")
        }
      }
    }
  }
  
}

//==EMAIL: THIS WEEKS TEAMS=====================================================================================================//

function thisWeekTeamsEmail() {
  
  var today = Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yy");
  var weekNumberString = sheetBalancer.getRange(2,2).getValue();
  var weekNumber = parseInt(weekNumberString.slice(15,weekNumberString.length));
  //var weekNumber = 19
  var subject = "Wednesday Football Teams - " + today + " (Season " + season + " - Week " + weekNumber + ")"
  
  var guestlist = getCalendarGuestList();
  var arrayAttending = [];
  
  for (var i = 0; i < guestlist.length; i++) {
    for (var j = 0; j < playerListEmails.length; j++) {
      if (guestlist[i].getEmail() == playerListEmails[j][0]) {
        if (guestlist[i].getGuestStatus() == "YES") {
          arrayAttending.push(guestlist[i].getEmail())
        }
      }
    }
  }
  
  var emailAddress = "nkumar6@jaguarlandrover.com";
  
  for (var k = 0; k < arrayAttending.length; k++) {
    if (arrayAttending[k] != "") {   // EMAIL ADDRESS BLANKED
      emailAddress = emailAddress + "," + arrayAttending[k];
    }
  }
  Logger.log(emailAddress)
  
  if (season == 1) {
    var emailTableColumns = 5;
  }
  else if (season == 2) {
    var emailTableColumns = 3;
  }
  
  var numBenched = sheetBalancer.getLastRow()-17;
  Logger.log(numBenched);
  
  MailApp.sendEmail({
    to: emailAddress,
    subject: subject,
    htmlBody: getHtmlTable(sheetBalancer.getRange(4,8,9+numBenched,emailTableColumns))
  }); 

}

//==EMAIL: THIS WEEKS TEAMS (VERSION TWO) ================================================================================================//

function thisWeekTeamsEmailNew() {
  
  var today = Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yy");
  var weekNumberString = sheetBalancer.getRange(2,2).getValue();
  var weekNumber = parseInt(weekNumberString.slice(15,weekNumberString.length));
  //var weekNumber = 19
  var subject = "Wednesday Football Teams - " + today + " (Season " + season + " - Week " + weekNumber + ")"
  
  var guestlist = getCalendarGuestList();
  var arrayAttending = [];
  
  for (var i = 0; i < guestlist.length; i++) {
    for (var j = 0; j < playerListEmails.length; j++) {
      if (guestlist[i].getEmail() == playerListEmails[j][0]) {
        if (guestlist[i].getGuestStatus() == "YES") {
          arrayAttending.push(guestlist[i].getEmail())
        }
      }
    }
  }
  
  var emailAddress = "nkumar6@jaguarlandrover.com";
  
  /*for (var k = 0; k < arrayAttending.length; k++) {
  if (arrayAttending[k] != "") {  // EMAIL ADDRESS BLANKED
  emailAddress = emailAddress + "," + arrayAttending[k];
  }
  }*/
  Logger.log(emailAddress)
  
  MailApp.sendEmail({
    to: emailAddress,
    subject: subject,
    htmlBody: getHtmlTable(sheetBalancer.getRange(4,8,7,3))
  }); 
  getHtmlTable(sheetBalancer.getRange(4,8,7,3))
}

//==EMAIL: NOT ENOUGH PLAYERS====================================================================================================================//

function emailNotEnoughPlayers() {
  
  var emailAddress = ""; // EMAIL ADDRESS BLANKED
  var today = Utilities.formatDate(new Date(), "GMT+1", "dd/MM/yy");
  var weekNumberString = sheetBalancer.getRange(2,2).getValue();
  var weekNumber = parseInt(weekNumberString.slice(15,weekNumberString.length));
  var subject = "Not Enough Players - " + today + " (Week " + weekNumber + ")"
  
  var arrayY = [];
  var arrayN = [];
  var arrayM = [];
  var arrayI = [];
  
  for (var i = 0; i < guestlist.length; i++) {
    for (var j = 0; j < playerListEmails.length; j++) {
      if (guestlist[i].getEmail() == playerListEmails[j][0]) {
        if (guestlist[i].getGuestStatus() == "YES") {
          arrayY.push(playerList[j])
        }
        else if (guestlist[i].getGuestStatus() == "NO") {
          arrayN.push(playerList[j])
        }
        else if (guestlist[i].getGuestStatus() == "MAYBE") {
          arrayM.push(playerList[j])
        }
        else if (guestlist[i].getGuestStatus() == "INVITED") {
          arrayI.push(playerList[j])
        }
      }
    }
  }
  
  MailApp.sendEmail({
    to: emailAddress,
    subject: subject,
    htmlBody: 
    "<b><u>Attending: " + arrayY.length + "</u></b><br>" + arrayY
    + "<br><br>"
    + "<b><u>Not Attending: " + arrayN.length + "</u></b><br>" + arrayN
    + "<br><br>"
    + "<b><u>Maybe: " + arrayM.length + "</u></b><br>" + arrayM
    + "<br><br>"
    + "<b><u>Not Responded: " + arrayI.length + "</u></b><br>" + arrayI
    + "<br><br>"
  }); 
  
}

//==CANCEL GAME AND SPLIT PAYMENT====================================================================================================================//

function cancelSplitPayment() {
  
  var weekNumber = 19
  for (var i = 0; i < playerList.length; i++) {
    if (playerListStatus[i][1] == "R" || playerListStatus[i][1] == "B") {
      sheetAttendance.getRange(14+i,17+weekNumber-1).setValue("x")
    }
  }
  sheetAttendance.getRange(8,17+weekNumber-1).setValue("x");
  sheetAttendance.getRange(9,17+weekNumber-1).setValue("x");
}

//==PARTNERSHIPS====================================================================================================================//

function masterPartnerships() {
  
  partnerships("currentSeason")
  partnerships("allTime")
  
}

function partnerships(seasonScope) {
  
  if (seasonScope == "currentSeason") {
    var sheetTempAttendance = sheetAttendance;
    var sheetTempPartnerships = sheetPartnerships;
  }
  else if (seasonScope == "allTime") {
    var sheetTempAttendance = sheetAttendanceAllTime;
    var sheetTempPartnerships = sheetPartnershipsAllTime;
  }
  
  var gamesPlayed = 0;
  
  var numPlayers = getLastPlayers(seasonScope);
  var numGames = getLastGame(seasonScope);
  Logger.log(numPlayers);
  Logger.log(numGames);
  
  var data = sheetTempAttendance.getRange(14,4,numPlayers,13+numGames).getValues();
  var widthData = data[0].length;
  var heightData = data.length;
  
  var tableWins = zero2D(numPlayers,numPlayers);
  var tableTotalGames = zero2D(numPlayers,numPlayers);
  
  var counter = 0
  var gameResults = sheetTempAttendance.getRange(11,17,1,sheetTempAttendance.getLastColumn()-17).getValues();
  
  for (var j = 0; j < widthData-13; j++) { // j is Game Week
    for (var i = 0; i < heightData; i++) { // i is Host Player
      for (var k = 0; k < heightData; k++) { // k is Partner Player
        if (data[k][j+13] == "" || data[k][j+13] == "x" || data[k][j+13] == "t" ) {
        }
        else {
          if (data[k][j+13] == data[i][j+13]) {
            var specResult = [data[k][0],data[i][0],data[k][j+13],gameResults[0][j],j]
            if (specResult[0] != specResult[1]) {
              if (specResult[2] == specResult[3]) {
                tableWins[i][k] = tableWins[i][k]+1;
                tableTotalGames[i][k] = tableTotalGames[i][k]+1;
              }
              else {
                tableTotalGames[i][k] = tableTotalGames[i][k]+1;
              }
            }
            
          }
        }
      }
    }
  }
  
  for (var a = 0; a < tableWins.length; a++) {
    for (var b = 0; b < tableWins[0].length; b++) {
      sheetTempPartnerships.getRange(5+b,3+a).setValue(tableWins[a][b] + " / " + tableTotalGames[a][b]);
    }
  }
}


function getLastGame(seasonScope){
  if (seasonScope == "currentSeason") {
    var sheetTempAttendance = sheetAttendance;
    var sheetTempPartnerships = sheetPartnerships;
  }
  else if (seasonScope == "allTime") {
    var sheetTempAttendance = sheetAttendanceAllTime;
    var sheetTempPartnerships = sheetPartnershipsAllTime;
  }
  var arrayGamesPlayed = sheetTempAttendance.getRange(6,17,1,sheetTempAttendance.getLastColumn()-17).getValues();
  for (var i = 0; i <= arrayGamesPlayed[0].length; i++) {
    Logger.log(arrayGamesPlayed[0].length)
    if (arrayGamesPlayed[0][i] != 0) {
    }
    else {
      return i
    }
  }
  return i
}

function getLastPlayers(seasonScope){
  if (seasonScope == "currentSeason") {
    var sheetTempAttendance = sheetAttendance;
    var sheetTempPartnerships = sheetPartnerships;
  }
  else if (seasonScope == "allTime") {
    var sheetTempAttendance = sheetAttendanceAllTime;
    var sheetTempPartnerships = sheetPartnershipsAllTime;
  }
  var arrayPlayers = sheetTempAttendance.getRange(14,4,sheetTempAttendance.getLastRow(),1).getValues();
  for (var i = 0; i <= arrayPlayers.length; i++) {
    if (arrayPlayers[i] != "") {
    }
    else {
      return i
    }
  }
  return i
}

//==EMAIL: PAYMENT CHASER====================================================================================================================//

function emailOutstandingPayments() {
  
  var data = sheetBalance.getRange("C5:H50").getValues();
  
  var heightData = data.length;
  var widthData = data[0].length;
  
  var arrayNotPaidNames = [];
  var arrayNotPaidEmails = [];
  var arrayNotPaidAmount = [];
  
  Logger.log(data)
  
  for (var i = 0; i < heightData; i++) {
    
    if (Math.round(100*data[i][5])/100 < 0) {
      
      arrayNotPaidNames.push(data[i][0]);
      arrayNotPaidAmount.push(data[i][5].toFixed(2));
      arrayNotPaidEmails.push(data[i][1]);
      
    }
    
  }
  
  Logger.log(arrayNotPaidEmails);
  
  var lenArrayNPN = arrayNotPaidNames.length;
  
  var arrayEmailsTest = [] // BLANKED
  
  Logger.log(arrayEmailsTest)
  var emailAddress = "";  // BLANKED
  emailAddress = emailAddress + "," + "";  // BLANKED
  
  for (var j = 0; j < arrayNotPaidEmails.length; j++) {
    emailAddress = emailAddress + "," + arrayNotPaidEmails[j];
  }
  
  //var emailAddress = "";   // BLANKED
  // var emailAddress = arrayNotPaidEmails;
  var subject = "Wednesday Football - Payment Tracker";
  
  var bullets = ""
  
  for (var j = 0; j < lenArrayNPN; j++) {
    bullets += 
      "<tr>"
    + "<td style=\"width:50%\"><font face=\"Trebuchet MS\">£" + arrayNotPaidAmount[j] + "</font></td>"
    + "<td><font face=\"Trebuchet MS\">" + arrayNotPaidNames[j] + "</font></td>"
    + "</tr>"
  }
  
  MailApp.sendEmail({
    to: emailAddress,
    subject: subject,
    htmlBody: "Hi all,<br><br>"
    + "Please find included the list of currently outstanding payments from the <a href=\"https://docs.google.com/spreadsheets/d/1Xd_gRNNkSJNAqWE4uARditENxQ391ZOklMs9SgoqRRM/edit?ts=58dc0029#gid=369654410\">Wednesday Football Payments</a> spreadsheet."
    + "<br><br>"
    + "<table>"
    + bullets
    + "</table>"
    + "<br>"
    + "Could you please top up your funds prior to next Wednesday's game using the following bank details:"
    + "<br><br>"
    + "<b>HSBC</b>"
    + "<br>"
    + ""  // NAME BLANKED
    + "<br>"
    + " <br> " // PAYMENT DETAILS BLANKED
    + "<br><br>"
    + "<i>Note: Can you please ensure Matt is cc'd on any response regarding payments</i>"
    + "<br><br>"
    + "Thanks"
  }); 
  
}

//==CALENDAR INVITE====================================================================================================================//

function calendarInvite() {
  
  var today = new Date();
  
  var dateCalEventStart = addDays(today,(3 + 7 - today.getDay()) % 7)
  dateCalEventStart.setHours(19);
  dateCalEventStart.setMinutes(0);
  dateCalEventStart.setSeconds(0);
  
  var dateCalEventEnd = addDays(today,(3 + 7 - today.getDay()) % 7);
  dateCalEventEnd.setHours(dateCalEventStart.getHours()+1);
  dateCalEventEnd.setMinutes(0);
  dateCalEventEnd.setSeconds(0);
  
  var title = "FINAL Social Football"
  var location = "" // BLANKED
  
  var description = "Weekly invite for Social Football\n"
  + "\n"
  + "Final Wednesday night football game (under current management at least!). Thanks all for your attendance over the years :)\n"
  + "\n"
  + "Rules:\n"
  + "---------\n"
  + "- 5/6-aside\n"
  + "- Indoor wooden floor\n"
  + "- Game kicks off at exactly 7pm - first team to get 5 players doesn't wear the bibs\n"
  + "- Cost of the game is £50 split by the number of players in attendance\n"
  + "- Priority will be given to those who are in the green on the payment sheet, followed by regulars taking priority over backups\n"
  + "- Payment status is on the 'Player List' tab of the spreadsheet below\n"
  + "\n"
  + "Scoring:\n"
  + "-----------\n"
  + "- W/D/L system, 4 points per win, 2 point per draw, 0 points per loss\n"
  + "- Bonus point on offer for:\n"
  + "  ~ Winning by 4+ goals\n"
  + "  ~ Losing by 1 goal\n"
  + "- Ranking in the table is based on your points AND attendance\n"
  + "\n"
  + "Current League Table:\n"
  + "https://docs.google.com/spreadsheets/d/1Xd_gRNNkSJNAqWE4uARditENxQ391ZOklMs9SgoqRRM/edit?ts=58dc0029#gid=322144541\n"
  //+ "\n"
  //+ "Matt"
  
  /*
  var arrayGuests = [
  
  ]
  
  BLANKED
  */
  var arrayGuests = []
  var arrayStandins = []
  for (var i = 0; i < playerListStatus.length; i++) {
    if (playerListStatus[i][1] != "") {
      if (playerListStatus[i][1] == "R" || playerListStatus[i][1] == "B") {
        arrayGuests.push(playerListEmails[i][0])
      }
      else {
        arrayStandins.push(playerListEmails[i][0])
      }
    }
  }
  Logger.log(arrayStandins)
  
  var guests = "nkumar6@jaguarlandrover.com";
  
  for (var i = 0; i < arrayGuests.length; i++) {
    guests = guests + "," + arrayGuests[i];
  }
  Logger.log(guests)
  calFootball.createEvent(title, dateCalEventStart, dateCalEventEnd, {description: description, location: location, guests:guests, sendInvites: true})
}

//=======================================================================================================//
