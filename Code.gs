// Director Schedule Spreadsheet to Google Calendar script written by Benjamin Bauer
// January 2017 / Version 0.6
// web app url - https://script.google.com/a/macros/weathergroup.com/s/AKfycbzT8Db6Z0EWjf-stHsyxojJ0dLq8zInUEnujP3bA5xrArSejnbD/exec
// 


function myFunction() {
  var ln = getLastName();
  var sheets = getSheets();
  
  pushWeekToCalendar(sheets.ThisWeek, ln);
  pushWeekToCalendar(sheets.NextWeek, ln);
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index');
}

//
function pushWeekToCalendar(ss, lastname) {
  
    //get sheet for each day of week
  var spreadsheet = SpreadsheetApp.open(ss);
  var DaysOfWeek = ["MONDAY","TUESDAY","WEDNESDAY","THURSDAY","FRIDAY","SATURDAY","SUNDAY"];
  
  function EventFromSchedule(start, end, title) { //object containing event info to be created for each day
      this.start = start; 
      this.end = end; 
      this.title = title;
  }
  
  //Assign current work day
  for (var x = 0; x < 7; x++) {
    var sheet = spreadsheet.getSheetByName(DaysOfWeek[x]); //grab sheet for that day              
  
    //spreadsheet variables
    var lastRow = 56;    //sheet.getLastRow(); TOO SLOW 
    var lastColumn = 30; //sheet.getLastColumn(); TOO SLOW 
    var range = sheet.getRange(1,1,56,30);
    var values = range.getValues();
    var calendar = CalendarApp.getDefaultCalendar(); 
    var Worker = 0; //init to zero for every day
    var Schedule = [];

    //find work column using last name
    for (var k = 3; k < lastColumn; k++) {
      if (values[2][k]) { //is field blank?
        var lower = values[2][k].toLowerCase();
        if (lower.search(lastname) != -1) { 
          Worker = k;
          k=lastColumn;
        }  
      }
    }  
    
    Logger.log(DaysOfWeek[x]);
    Logger.log('Searched '+ k + ' columns.');
    if (!Worker) { Logger.log('No worker found on this day.')}
    else { Logger.log('Worker name found in column ' + Worker);}
    
    
    if (Worker) { //if no worker then don't do anything, go to next day
    
      for (var i = 5; i < lastRow; i++) {
        var ECount = 1; 
                
        //Logger.log('Row: ' + i);
        
        //if job value found this create/modify events, if not delete calendar events if found, ignore white space // check if PTO or OPH event
        if (values[i][Worker].trim() && (values[i][Worker].search('PTO') == -1) && (values[i][Worker].search('OPH') == -1)) {  
          Logger.log('Job found ' + values[i][Worker]);
                    
          for (var u = i; u < lastRow; u++) { //find show names       
            var Job = values[u][Worker].trim();
            var NextJob = values[u+1][Worker].trim();        
          
            //find show name
            if (values[u][1]) {         
              var Show = values[u][1];
            }
            else {
              for (var j = u-1; j > 4; j--) {
                if (values[j][1]) {
                  var Show = values[j][1];
                  j = 4;
                }
              }
            }  
            
            //find NEXT show name
            var NextShow = values[u+1][1] ? values[u+1][1] : Show;
            
            //decide whether to combine into one event or start a new one
            if ((Job==NextJob) && (Show==NextShow)) {
              ECount += 1;}
            else {
              var CurrentRow = u;
              u = lastRow;
            }  
          }
          
          
          var newEventTitle = "On-Air: " + Show + " (" + Job + ")";
      
          var newEventStart = FindShowStartTime(values, i);
          
          var newEventEnd = new Date(newEventStart);
          newEventEnd.setMinutes(newEventEnd.getMinutes()+(ECount*30)); //multiply by the amount of extra blocks of :30
          
          
          
          //put into object array
          Schedule.push(new EventFromSchedule(newEventStart, newEventEnd, newEventTitle)); 
          
          i = CurrentRow; //update to latest row
          Logger.log(i);
        }
      } 
      
      //grab events from calendar
      var Today = new Date(values[1][5]);
      var OnAirEvent = calendar.getEventsForDay(Today, {search: 'On Air:'});
      
      //cross check calendar events with schedule spreadsheet events THEN delete non-exact matches
      for (var c = 0; c < OnAirEvent.length; c++) { //each calendar event found
        var delCalEvent = true; //init as no match and to be deleted
        
        for (var d = 0; d < Schedule.length; d++) { //check with every schedule event
          if ((OnAirEvent[c].getStartTime().valueOf() == Schedule[d].start.valueOf()) && (OnAirEvent[c].getEndTime().valueOf() == Schedule[d].end.valueOf()) && (OnAirEvent[c].getTitle() == Schedule[d].title)) {
            Logger.log("Found matching event: " + OnAirEvent[c].getTitle() + " starting at " + OnAirEvent[c].getStartTime() + " ending at " + OnAirEvent[c].getEndTime());
            delCalEvent = false; //match found, don't delete
          }
        }
        
        if (delCalEvent) {
          OnAirEvent[c].deleteEvent();
          Logger.log("Event deleted: "  + OnAirEvent[c].getTitle() + " starting at " + OnAirEvent[c].getStartTime() + " ending at " + OnAirEvent[c].getEndTime());
        }
      }
      
      //clean up duplicates
      for (var e = 0; e < OnAirEvent.length; e++) { //each calendar event found
        var DuplicateEventCheck = calendar.getEvents(OnAirEvent[e].getStartTime(), OnAirEvent[e].getEndTime(), {search: OnAirEvent[e].title});
          
        if (DuplicateEventCheck[1]) {
          for (var a = 1; a < DuplicateEventCheck.length; a++) {
            DuplicateEventCheck[a].deleteEvent();
            Logger.log('Duplicate event deleted.');
          }
        }
      }
      
      //write new events
      for (var y = 0; y < Schedule.length; y++) {
        var makeEvent = true;
        
        for (var z = 0; z < OnAirEvent.length; z++) { //each calendar event found
          if ((OnAirEvent[z].getStartTime().valueOf() == Schedule[y].start.valueOf()) && (OnAirEvent[z].getEndTime().valueOf() == Schedule[y].end.valueOf()) && (OnAirEvent[z].getTitle() == Schedule[y].title)) {
            Logger.log("Found matching event: " + OnAirEvent[z].getTitle() + " starting at " + OnAirEvent[z].getStartTime() + " ending at " + OnAirEvent[z].getEndTime());
            makeEvent = false; //match found, don't make
          }
        }
        
        if (makeEvent) {
          var newEvent = calendar.createEvent(Schedule[y].title, Schedule[y].start, Schedule[y].end);
          var newEventId = newEvent.getId();
          Logger.log("Event created; " + Schedule[y].title + " from " + Schedule[y].start + " to " + Schedule[y].end)
        }
      }  
    }
  }
}



//finds show time from a spreadsheet range and given row  
function FindShowStartTime(val, i) {

  //var Today = ; //from same place in every sheet
  var Start = new Date(val[1][5]);
  var JobTime = val[i][0]; //value for Time
  var res = JobTime.split(':'); //split into two different strings at the ":"
  var Hrs = parseInt(res[0], 10); //+3 to make spreadsheet correct, take out +3 to make calender correct
  var Mins = parseInt(res[1].slice(0,res[1].length-1), 10); //strip "A" or "P" off end AND parse into integer
  
  //check to see if AM/PM, change hours accordingly 
  if ((JobTime.search('A') == -1) && (Hrs < 12)) { //set PM hours to 24
    Hrs += 12; //PM offset +3 for some reason.. GMT to EST?
  } 
  
  if ((JobTime.search('A') != -1) && (i > 45)) { //set 12A onward down sheet to next day
    Start.setDate(Start.getDate()+1);
    Hrs = Hrs == 12 ? 0 : Hrs; 
  }
  
  Start.setHours(Hrs);
  Start.setMinutes(Mins);

  return Start;
}



//gets last of name of active user
function getLastName() {
  var email = Session.getActiveUser().getEmail();
  
  //find name of user
  var r = email.split('@'); //split into two different strings at the "@"
  var name = r[0].split('.'); //split into two different strings at the "."
  return name[1]; //assign second string to last name and return value
}



//gets this and next week's director's schedule spreadsheets
function getSheets() {
  // Find parent 2017 Schedule Folder
//  var AllFolders = DriveApp.searchFolders('title contains "2017 Production Control Schedules"');
//  while (AllFolders.hasNext()) {
//    var PCSfolder = AllFolders.next();
//    Logger.log(PCSfolder.getName());
//    Logger.log(PCSfolder.getId()); // 0B6BXSE7lNq7FNXdxaGRVQzZ1VjQ
//  }
//  
//  
//  //Find child Director Schedule Folder
//  var DirFolder = PCSfolder.searchFolders('title contains "DIRECTOR SCHEDULES"');
//  while (DirFolder.hasNext()) {
//    var Dfolder = DirFolder.next();
//    Logger.log(Dfolder.getName());
//    Logger.log(Dfolder.getId()); // 0B6BXSE7lNq7FX1dsM3RLRHBaRTQ
//  }
  
  //skip all that use exact id for Director Schedules folder
  var Dfolder = DriveApp.getFolderById('0B6BXSE7lNq7FX1dsM3RLRHBaRTQ');
  
  //Find child month folder
  var Today = new Date();
  
  //Get current month in string form
  var MonthOfYear = ["JANUARY","FEBRUARY","MARCH","APRIL","MAY","JUNE","JULY","AUGUST","SEPTEMBER","OCTOBER","NOVEMBER","DECEMBER"];
  var Month = MonthOfYear[Today.getMonth()]; //Get string form of month
  var LastMonth = (Month == 0) ? 11 : MonthOfYear[Today.getMonth()-1];
  var NextMonth = MonthOfYear[Today.getMonth()+1];
  Logger.log('This month: ' + Month)
  Logger.log('Last month: ' + LastMonth);
  Logger.log('Next month: ' + NextMonth);

  var MonthFolder = Dfolder.searchFolders('title contains "'+ Month +'"');
  while (MonthFolder.hasNext()) {
    var folder = MonthFolder.next();
    Logger.log(folder.getName());
  }
  
    //Find this and next week's sheets
  var Day = Today.getDate(); //returns day of month int betw 1 - 31
  var folderSheets = folder.getFiles();
  var gotoLastMonth = false;
  
  while (folderSheets.hasNext()) {
    var Sheets = folderSheets.next();
    var WeekOf = Sheets.getName().toUpperCase();
    var W = WeekOf.search(Month);
    var WeekOfNum = WeekOf.slice(W+Month.length);
    
    Logger.log('Week of Number: ' + WeekOfNum);
    
    //if (Day < WeekOfNum) {var gotoLastMonth == true}; //check if we need to go find this week in last month folder
    gotoLastMonth = (Day < WeekOfNum) ? true : false;    
        
    if (((Day-WeekOfNum) < 7) && ((Day-WeekOfNum) >= 0)) {
      var ThisWeekSheet = Sheets;    
      Logger.log('This weeks calender: ' + ThisWeekSheet);
    }
     
    if ((Math.abs(Day-WeekOfNum) <= 7) && ((Day-WeekOfNum) < 0)) {
      var NextWeekSheet = Sheets; 
      Logger.log('Next weeks calender: ' + NextWeekSheet);
    }
  }  

  //find this week's sheet if we need to go into last month
  if (gotoLastMonth) {
    var DirFolder = Dfolder.searchFolders('title contains "'+ LastMonth +'"');
    while (DirFolder.hasNext()) {
      var folder = DirFolder.next();
      Logger.log(folder.getName());
    }
    
    var folderSheets = folder.getFiles();
    var WeekOfNumCheck;
    
    while (folderSheets.hasNext()) {
      var Sheets = folderSheets.next();
      var WeekOf = Sheets.getName().toUpperCase();
      var W = WeekOf.search(LastMonth);
      var WeekOfNum = WeekOf.slice(W+LastMonth.length);
      
      if (!WeekOfNumCheck) {WeekOfNumCheck = parseInt(WeekOfNum, 10)}; //grabs first WeekOf number 
      
      if (WeekOfNumCheck <= WeekOfNum) {ThisWeekSheet = Sheets}; //checks against all others and returns largest
    }
    
    Logger.log('This is the last week of last month: ' + ThisWeekSheet);      
  }
  
  //find next week's sheet if we need to go into next month 
  if (!NextWeekSheet) { 
    var DirFolder = Dfolder.searchFolders('title contains "'+ NextMonth +'"');
    while (DirFolder.hasNext()) {
      var folder = DirFolder.next();
      Logger.log(folder.getName());
    }  

    var folderSheets = folder.getFiles();
    
    while (folderSheets.hasNext()) {
      var Sheets = folderSheets.next();
      var WeekOf = Sheets.getName().toUpperCase();
      var W = WeekOf.search(NextMonth);
      var WeekOfNum = WeekOf.slice(W+NextMonth.length);
      
      if (WeekOfNum < 7) {
        var NextWeekSheet = Sheets;    
        Logger.log('This is the first week of next month: ' + NextWeekSheet);
      }
    }      
  }  
  
  return {ThisWeek: ThisWeekSheet, NextWeek: NextWeekSheet};
}