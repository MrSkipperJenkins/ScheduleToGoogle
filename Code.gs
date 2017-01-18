function myFunction() {
  getFolders();
}

function pushWeekToCalendar(spreadsheet) {
  
  //get sheet for each day of week
  var DaysOfWeek = ["MONDAY","TUESDAY","WEDNESDAY","THURSDAY","FRIDAY","SATURDAY","SUNDAY"];
  var email = Session.getActiveUser().getEmail();
  
  //find name of user
  var r = email.split('@'); //split into two different strings at the "@"
  var name = r[0].split('.'); //split into two different strings at the "."
  var lastname = name[1]; //assign second string to last name
  
  //Assign current work day
  for (var d = 0; d < 7; d++) {
    var sheet = spreadsheet.getSheetByName(DaysOfWeek[d]); //grab sheet for that day              
  
    //spreadsheet variables
    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn(); 
    var range = sheet.getRange(1,1,lastRow,lastColumn);
    var values = range.getValues();
    var calendar = CalendarApp.getDefaultCalendar();
    var Worker = 0; //init to zero for every day
    
    //find work col with last name
    for (var k = 3; k < lastColumn; k++) {
      if (values[2][k]) { //is field blank?
        var lower = values[2][k].toLowerCase();
        if (lower.search(lastname) != -1) { 
          Worker = k;
          k=lastColumn;
        }  
      }
    }  
    
    Logger.log(DaysOfWeek[d]);
    Logger.log('Searched '+ k + ' columns.');
    if (!Worker) { Logger.log('No worker found on this day.')}
    else { Logger.log('Worker name found in column ' + Worker);}
    
    
    if (Worker) { //if no worker then don't do anything else
    
      for (var i = 5; i < lastRow; i++) {
        var ECount = 1; 
        //Logger.log('Row: ' + i);
        
        if (values[i][Worker].trim()) {  //if job value found this create/modify events, if not delete calendar events if found, ignore white space
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
          
          //var SpecialCase = Job.search('CNBC');
          //Logger.log('CNBC? ' + SpecialCase);
          
          //combine variables into event title here   
          //if (Job.search('CNBC') !== -1) {                                               // WHY WONT THIS WORK? it works fine it tuns out, it was just getting checked over and not replaced for new event
          //  var newEventTitle = "On-Air: (" + Job + ")";
          //  Logger.log('CNBC Title changed...');
          //}
          //else {
            var newEventTitle = "On-Air: " + Show + " (" + Job + ")";
          //  Logger.log('Title not changed...');
          //}
          
          //find show time
          var newEventStart = new Date(values[1][5]); //from same place in every sheet
          var WTime = values[i][0]; //val for Time
          var res = WTime.split(':'); //split into two different strings at the ":"
          var Hours = parseInt(res[0], 10); //+3 to make spreadsheet correct, take out +3 to make calender correct
          
          
          //check to see if AM/PM, change hours accordingly 
          if ((WTime.search('A') == -1) && (Hours < 12)) { //set PM hours to 24
            Hours += 12; //PM offset +3 for some reason.. GMT to EST?
          } 
          
          if ((WTime.search('A') != -1) && (i > 45)) { //set 12A onward down sheet to next day
            newEventStart.setDate(newEventStart.getDate()+1);
            Hours = Hours == 12 ? 0 : Hours; 
          }
          
          newEventStart.setHours(Hours);
          
          
          //grab minutes and assign
          var Mins = parseInt(res[1].slice(0,res[1].length-1), 10); //strip "A" or "P" off end AND parse into integer
          newEventStart.setMinutes(Mins); 
          
          //create end time set for a half hour later
          var newEventEnd = new Date(newEventStart);
          newEventEnd.setMinutes(Mins+(ECount*30)); //multiply by the amount of extra blocks of :30
          
          //check if duplicate & create calender event
          //var CheckEvent = calendar.getEvents(newEventStart, newEventEnd, {search: newEventTitle});
          
          //if (CheckEvent=='') {
          //  var newEvent = calendar.createEvent(newEventTitle, newEventStart, newEventEnd);
          //  var newEventId = newEvent.getId();
            
          //  i = CurrentRow; //update to latest row
          //}
          
          var OnAirEvent = calendar.getEvents(newEventStart, newEventEnd, {search: 'On-Air:'});
          
          //delete any duplicates before starting
          if (OnAirEvent[1]) {
            for (var a = 1; a < OnAirEvent.length; a++) {
              OnAirEvent[a].deleteEvent();
              Logger.log('Duplicate event deleted.');
            }
          }  
          
          
            
          
          if (OnAirEvent[0]) {
            Logger.log('On-Air event found');
            
            //var CheckIfExactEvent = calendar.getEvents(newEventStart, newEventEnd, {search: newEventTitle});
            var CheckTitle = OnAirEvent[0].getTitle();    
            Logger.log('Title: ' + CheckTitle);
            
            if ((CheckTitle.search(Show) == -1) || (CheckTitle.search(Job) == -1)) { 
              Logger.log('Show or Job did not match at time slot. Delete events and write new.');
              //delete old and make new
              //for (var a = 0; a < OnAirEvent.length; a++) {
                OnAirEvent[0].deleteEvent();
                Logger.log('Event deleted.');
              //}
                            
              var newEvent = calendar.createEvent(newEventTitle, newEventStart, newEventEnd);
              var newEventId = newEvent.getId();
              
              i = CurrentRow; //update to latest row
              Logger.log(i);
            } 
            else {
              Logger.log('Show and Job matched at time slot.');
              
              //Delete matched Job/Work events inside time slot
              var TestEndTime = OnAirEvent[0].getEndTime();                 // CHANGE TO STRING MAYBE? to avoid delcaring unnecessary vars
              TestEndTime = TestEndTime.getHours() + (TestEndTime.getMinutes() / 10);
              var newEventEndTimeCheck = newEventEnd.getHours() + (newEventEnd.getMinutes() / 10);
              
              if (TestEndTime != newEventEndTimeCheck) {
                Logger.log('End time "' + TestEndTime + '" did not match "' + newEventEndTimeCheck + '" (deleting event and creating new).');
                OnAirEvent[0].deleteEvent();
                
                var newEvent = calendar.createEvent(newEventTitle, newEventStart, newEventEnd);
                var newEventId = newEvent.getId();
              }
              
              //if show and job match but there is another shorter duplicate event within time slot, delete one within
              //if (OnAirEvent[1]) {
              //  for (var a = 1; a < OnAirEvent.length; a++) {
              //    OnAirEvent[a].deleteEvent();
              //    Logger.log('Duplicate event deleted.');
              //  }
              //}  
              
              i = CurrentRow; //update to latest row
              Logger.log(i);
            }
          }
          
          else {
            Logger.log('No On-Air event found, creating new.');
            //no event go ahead and make it
            var newEvent = calendar.createEvent(newEventTitle, newEventStart, newEventEnd);
            var newEventId = newEvent.getId();
            
            i = CurrentRow; //update to latest row
            Logger.log(i);
          }
        } 
        else { //nothing in schedule so check time and delete calendar event if exists SO SLOOOOOOW
          
          //find show time
          var newEventStart = new Date(values[1][5]); //from same place in every sheet
          var WTime = values[i][0]; //val for Time
          var res = WTime.split(':'); //split into two different strings at the ":"
          var Hours = parseInt(res[0], 10); //+3 to make spreadsheet correct, take out +3 to make calender correct
          
          
          //check to see if AM/PM, change hours accordingly 
          if ((WTime.search('A') == -1) && (Hours < 12)) { //set PM hours to 24
            Hours += 12; //PM offset +3 for some reason.. GMT to EST?
          } 
          
          if ((WTime.search('A') != -1) && (i > 45)) { //set 12A onward down sheet to next day
            newEventStart.setDate(newEventStart.getDate()+1);
            Hours = Hours == 12 ? 0 : Hours; 
          }
          
          newEventStart.setHours(Hours);
          
          
          //grab minutes and assign
          var Mins = parseInt(res[1].slice(0,res[1].length-1), 10); //strip "A" or "P" off end AND parse into integer
          newEventStart.setMinutes(Mins); 
          
          //create end time set for a half hour later
          var newEventEnd = new Date(newEventStart);
          newEventEnd.setMinutes(Mins+(ECount*30)); //multiply by the amount of extra blocks of :30
          
          var OnAirEvent = calendar.getEvents(newEventStart, newEventEnd, {search: 'On-Air:'});
          
          //delete any duplicates before starting
          if (OnAirEvent) {
            for (var a = 0; a < OnAirEvent.length; a++) {
              OnAirEvent[a].deleteEvent();
              Logger.log('Calendar event found without schedule entry. ' + OnAirEvent[a].getTitle() + ' deleted.');
            }
          }  
        }
      }    
    }      
  }
}










function getFolders() {
  // Find parent 2017 Schedule Folder
  var AllFolders = DriveApp.searchFolders('title contains "2017 Production Control Schedules"');
  while (AllFolders.hasNext()) {
    var PCSfolder = AllFolders.next();
    Logger.log(PCSfolder.getName());
  }
  
  //Find child Director Schedule Folder
  var DirFolder = PCSfolder.searchFolders('title contains "DIRECTOR SCHEDULES"');
  while (DirFolder.hasNext()) {
    var Dfolder = DirFolder.next();
    Logger.log(Dfolder.getName());
  }
  
  //Find child month folder
  var Today = new Date();
  var M = Today.getMonth();

  //Get current month in string form
  var MonthOfYear = ["JANUARY","FEBRUARY","MARCH","APRIL","MAY","JUNE","JULY","AUGUST","SEPTEMBER","OCTOBER","NOVEMBER","DECEMBER"];
  var Month = MonthOfYear[M]; //Get string form of month
  var NextMonth = MonthOfYear[(M+1)];
  Logger.log(Month+NextMonth);
  
  var MonthFolder = Dfolder.searchFolders('title contains "'+ Month +'"');
  while (MonthFolder.hasNext()) {
    var folder = MonthFolder.next();
    Logger.log(folder.getName());
  }
  
  //Find this and next week's sheets
  var Day = Today.getDate(); //returns day of month int betw 1 - 31
  var folderSheets = folder.getFiles();
 
  while (folderSheets.hasNext()) {
    var Sheets = folderSheets.next();
    var SheetNames = Sheets.getName();
    var WeekOf = SheetNames.toUpperCase();
    var W = WeekOf.search(Month);
    var WeekOfNum = WeekOf.slice(W+Month.length);
    
    Logger.log('Week of Number: ' + WeekOfNum);
    
    if (((Day-WeekOfNum) < 7) && ((Day-WeekOfNum) >= 0)) {
      var ThisWeekSheet = Sheets;    
      Logger.log('This weeks calender: ' + ThisWeekSheet);
    }
     
    if ((Math.abs(Day-WeekOfNum) <= 7) && ((Day-WeekOfNum) < 0)) {
      var NextWeekSheet = Sheets; 
      Logger.log('Next weeks calender: ' + NextWeekSheet);
    }
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
      var SheetNames = Sheets.getName();
      var WeekOf = SheetNames.toUpperCase();
      var W = WeekOf.search(NextMonth);
      var WeekOfNum = WeekOf.slice(W+NextMonth.length);
      
      if (WeekOfNum < 7) {
        var NextWeekSheet = Sheets;    
        Logger.log('This is the first week of next month: ' + NextWeekSheet);
      }
    }      
  }  
  
  pushWeekToCalendar(SpreadsheetApp.open(ThisWeekSheet));
  pushWeekToCalendar(SpreadsheetApp.open(NextWeekSheet));
}  


//deletes all events starting from today to two weeks ahead 
function DeleteAllCalendarEvents() {
  var calendar = CalendarApp.getCalendarsByName('On-Air Schedule');
  if (calendar) {
    Logger.log('Calender Found!');}
  else {
    Logger.log('Calender Not Found!');}
   
  //need to delete current and next week's events (14 days starting today)
  for (var w = 0; w < 14; w++) { //start at 0 for today
    var Today = new Date();
    var Day = Today.getDate(); //returns day of month int betw 1-31
    
    Today.setDate(Day+w); //adj to current working date
    Logger.log(Today);
    
    var events = calendar[0].getEventsForDay(Today);
  
    for (var e = 0; e < events.length; e++) {
      Logger.log('Event Deleted: ' + events[e]);
      events[e].deleteEvent();
    }  
  }  
} 


function TestBed() {
  
  //Logger.log();
  
  
}