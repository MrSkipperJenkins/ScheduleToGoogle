function myFunction() {
  
  var ln = getLastName();
  var sheets = getSheets(); 
  
  pushWeekToCalendar(sheets.ThisWeek, ln);
  pushWeekToCalendar(sheets.NextWeek, ln);

}

function pushWeekToCalendar(ss, lastname) {
  
  //get sheet for each day of week
  var spreadsheet = SpreadsheetApp.open(ss);
  var DaysOfWeek = ["MONDAY","TUESDAY","WEDNESDAY","THURSDAY","FRIDAY","SATURDAY","SUNDAY"];
  
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
    var EventCount = 0; //init to zero for each day
    var DaysEvents = { //object containing event info to be created for each day
      start: "", 
      end: "", 
      title: "",
      createEvent: function() {return calendar.createEvent(this.Title, this.Start, this.End)}
    }; 
    
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
          //if (Job.search('CNBC') !== -1) {                // WHY WONT THIS WORK? it works fine it tuns out, it was just getting checked over and not replaced for new event
          //  var newEventTitle = "On-Air: (" + Job + ")";
          //  Logger.log('CNBC Title changed...');
          //}
          //else {
            var newEventTitle = "On-Air: " + Show + " (" + Job + ")";
          //  Logger.log('Title not changed...');
          //}
          
          //find show time
          var Today = new Date(values[1][5]); //from same place in every sheet
          var newEventStart = Today;
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
              
              DaysEvents[EventCount].start = newEventStart;
              DaysEvents[EventCount].end = newEventEnd;
              DaysEvents[EventCount].title = newEventTitle;
              EventCount += 1;
              
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
                
                DaysEvents[EventCount].start = newEventStart;
                DaysEvents[EventCount].end = newEventEnd;
                DaysEvents[EventCount].title = newEventTitle;
                EventCount += 1;
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
            
            DaysEvents[EventCount].start = newEventStart;
            DaysEvents[EventCount].end = newEventEnd;
            DaysEvents[EventCount].title = newEventTitle;
            EventCount += 1;
            
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
    
    try { 
      Logger.log(EventCount + DaysEvents[EventCount].title);
    }
    
    catch(err) {
      Logger.log(err);
    }
    
  }
}










function getSheets() {
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
  
  //Get current month in string form
  var MonthOfYear = ["JANUARY","FEBRUARY","MARCH","APRIL","MAY","JUNE","JULY","AUGUST","SEPTEMBER","OCTOBER","NOVEMBER","DECEMBER"];
  var Month = MonthOfYear[Today.getMonth()]; //Get string form of month
  var NextMonth = MonthOfYear[Today.getMonth()+1];
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
    var WeekOf = Sheets.getName().toUpperCase();
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









function getLastName() {
  var email = Session.getActiveUser().getEmail();
  
  //find name of user
  var r = email.split('@'); //split into two different strings at the "@"
  var name = r[0].split('.'); //split into two different strings at the "."
  return name[1]; //assign second string to last name and return value
}










//find show time CREATE FUNCTION

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





function myTESTBED() {
  var ln = getLastName();
  var sheets = getSheets();
  
  TESTBED(sheets.ThisWeek, ln);
  TESTBED(sheets.NextWeek, ln);
}


function TESTBED(ss, lastname) {
  
    //get sheet for each day of week
  var spreadsheet = SpreadsheetApp.open(ss);
  var DaysOfWeek = ["MONDAY","TUESDAY","WEDNESDAY","THURSDAY","FRIDAY","SATURDAY","SUNDAY"];
  
  //Assign current work day
  for (var d = 0; d < 7; d++) {
    var sheet = spreadsheet.getSheetByName(DaysOfWeek[d]); //grab sheet for that day              
  
    //spreadsheet variables
    var lastRow = sheet.getLastRow();
    var lastColumn = sheet.getLastColumn(); 
    var range = sheet.getRange(1,1,lastRow,lastColumn);
    var values = range.getValues();
    var calendar = CalendarApp.getDefaultCalendar(); //maybe don't need this
    var Worker = 0; //init to zero for every day
    var Schedule = [];
    var test = 1;
    
    function EventFromSchedule(start, end, title) { //object containing event info to be created for each day
      this.start = start; 
      this.end = end; 
      this.title = title;
      this.createEvent = function() {return calendar.createEvent(this.title, this.start, this.end);}
    }
    
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
    
    Logger.log(DaysOfWeek[d]);
    Logger.log('Searched '+ k + ' columns.');
    if (!Worker) { Logger.log('No worker found on this day.')}
    else { Logger.log('Worker name found in column ' + Worker);}
    
    
    if (Worker) { //if no worker then don't do anything, go to next day
    
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
          
          
          var newEventTitle = "On-Air: " + Show + " (" + Job + ")";
      
          var newEventStart = FindShowStartTime(values, i);
          
          var newEventEnd = new Date(newEventStart);
          newEventEnd.setMinutes(newEventEnd.getMinutes()+(ECount*30)); //multiply by the amount of extra blocks of :30
          
          
          
          //put into our object array
          Schedule.push(new EventFromSchedule(newEventStart, newEventEnd, newEventTitle)); 
          
          i = CurrentRow; //update to latest row
          Logger.log(i);
        }
      } 
      
      //out all info to logger
      for (var b = 0; b < Schedule.length; b++) {
        for (var prop in Schedule[b]) {
          Logger.log("Schedule array[" + b + "]:" + prop + "=" + Schedule[b][prop]);
        }
      }
      
      //grab events from calendar
      var Today = new Date(values[1][5]);
      
      var OnAirEvent = calendar.getEventsForDay(Today, {search: 'On Air:'});
      
//      var CalEventStartTime = OnAirEvent[0].getStartTime();
//      Logger.log(CalEventStartTime);
//      Logger.log(CalEventStartTime.valueOf());
//      Logger.log(Schedule[0].start);
//      Logger.log(Schedule[0].start.valueOf());
//      if (CalEventStartTime == Schedule[0].start) {
//        Logger.log("DATE MATCH");
//      }
      
      
      
      //cross check calendar events with schedule spreadsheet events THEN delete non-exact matches
      for (var c = 0; c < OnAirEvent.length; c++) {
        //var CalEventStartTime = OnAirEvent[c].getStartTime().valueOf();
        //var CalEventEndTime = OnAirEvent[c].getEndTime().valueOf(); 
               
        for (var d = 0; d < Schedule.length; d++) {
                    
          if ((OnAirEvent[c].getStartTime().valueOf() != Schedule[d].start.valueOf()) || (OnAirEvent[c].getEndTime().valueOf() != Schedule[d].end.valueOf()) || (OnAirEvent[c].getTitle() != Schedule[d].title)) {
            Logger.log("Found non-matching event: " + OnAirEvent[c].getTitle() + " starting at " + OnAirEvent[c].getStartTime());
          }
        }
      }
      
      
      
          //delete any duplicates before starting
//          if (OnAirEvent[1]) {
//            for (var a = 1; a < OnAirEvent.length; a++) {
//              OnAirEvent[a].deleteEvent();
//              Logger.log('Duplicate event deleted.');
//            }
//          }  
          
          

//          
//      NEW CODE SECTION    
//          
//      
//          var DaysEvents = calendar.getEventsForDay(Today, {search: 'On Air:'});
//          
//          if (DaysEvents) {
//            for (var a = 0; a < DaysEvents.length; a++) {
//              DaysEvents[a].deleteEvent();
//              Logger.log('Calendar event found without schedule entry. ' + OnAirEvent[a].getTitle() + ' deleted.');
//            
//
//      NEW CODE SECTION
//
//
         
//          if (OnAirEvent[0]) {
//            Logger.log('On-Air event found');
//            
//            //var CheckIfExactEvent = calendar.getEvents(newEventStart, newEventEnd, {search: newEventTitle});
//            var CheckTitle = OnAirEvent[0].getTitle();    
//            Logger.log('Title: ' + CheckTitle);
//            
//            if ((CheckTitle.search(Show) == -1) || (CheckTitle.search(Job) == -1)) { 
//              Logger.log('Show or Job did not match at time slot. Delete events and write new.');
//              //delete old and make new
//              //for (var a = 0; a < OnAirEvent.length; a++) {
//                OnAirEvent[0].deleteEvent();
//                Logger.log('Event deleted.');
//              //}
//                            
//              var newEvent = calendar.createEvent(newEventTitle, newEventStart, newEventEnd);
//              var newEventId = newEvent.getId();
//              
//              ///////////////////              DaysEvents[EventCount] = {start:newEventStart, end:newEventEnd, title:newEventTitle};
//              
//              i = CurrentRow; //update to latest row
//              Logger.log(i);
//            } 
//            else {
//              Logger.log('Show and Job matched at time slot.');
//              
//              //Delete matched Job/Work events inside time slot
//              var TestEndTime = OnAirEvent[0].getEndTime();                 // CHANGE TO STRING MAYBE? to avoid delcaring unnecessary vars
//              TestEndTime = TestEndTime.getHours() + (TestEndTime.getMinutes() / 10);
//              var newEventEndTimeCheck = newEventEnd.getHours() + (newEventEnd.getMinutes() / 10);
//              
//              if (TestEndTime != newEventEndTimeCheck) {
//                Logger.log('End time "' + TestEndTime + '" did not match "' + newEventEndTimeCheck + '" (deleting event and creating new).');
//                OnAirEvent[0].deleteEvent();
//                
//                var newEvent = calendar.createEvent(newEventTitle, newEventStart, newEventEnd);
//                var newEventId = newEvent.getId();
//              }
//              
//              //if show and job match but there is another shorter duplicate event within time slot, delete one within
//              //if (OnAirEvent[1]) {
//              //  for (var a = 1; a < OnAirEvent.length; a++) {
//              //    OnAirEvent[a].deleteEvent();
//              //    Logger.log('Duplicate event deleted.');
//              //  }
//              //}  
//              
//              i = CurrentRow; //update to latest row
//              Logger.log(i);
//            }
//          }
//          
//          else {
//            Logger.log('No On-Air event found, creating new.');
//            //no event go ahead and make it
//            var newEvent = calendar.createEvent(newEventTitle, newEventStart, newEventEnd);
//            var newEventId = newEvent.getId();
//            
//            i = CurrentRow; //update to latest row
//            Logger.log(i);
//          }
         
      
    }      
  }
}