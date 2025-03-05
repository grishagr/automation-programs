// Determines the hour past which the script will fetch next day's events (default 11 am)
// So if this value is set to 11 and you run the script past 11 am, then the doc will be updated to tomorrow's date
// Otherwise, it will update to today's date
const TIME_TO_UPDATE_TO_TMR = 11;

function updateScheduleDoc() {
  var docId = '<google doc template id>'; // Google Doc ID
  var firstFloorCalendarId = '<first floor calendar id>'; // first floor calendar ID
  var secondFloorCalendarId = '<second floor calendar id>'; // second floor calendar ID

  var doc = DocumentApp.openById(docId);
  var body = doc.getBody();
  
  body.clear();

  insertPage(body, "First Floor Conference Room", firstFloorCalendarId);
  body.appendPageBreak();

  insertPage(body, "First Floor Conference Room", firstFloorCalendarId);
  body.appendPageBreak();

  var formattedDate = insertPage(body, "Second Floor Conference Room", secondFloorCalendarId);

  doc.saveAndClose();

  Logger.log("Fetched events for "+ formattedDate)
  Logger.log("Check updated document at " + doc.getUrl())
}

function insertPage(body, roomName, calendarId) {

  var date = new Date();
  
  if (isAfterTime()){ // use tomorrow's date and fetch events
    var events = getTomorrowEvents(calendarId);
    date.setDate(date.getDate() + 1);
    var formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), "EEEE, MMMM d, yyyy");
  }
  else { // use todays date and fetch events
    var events = getTodayEvents(calendarId);
    var formattedDate = Utilities.formatDate(date, Session.getScriptTimeZone(), "EEEE, MMMM d, yyyy");
  }

  body.appendParagraph("Elihu Root House").setBold(true).setAlignment(DocumentApp.HorizontalAlignment.CENTER)
      .setFontFamily("Times New Roman").setFontSize(17).setUnderline(false);

  body.appendParagraph(roomName).setBold(true).setAlignment(DocumentApp.HorizontalAlignment.CENTER)
      .setFontFamily("Times New Roman").setFontSize(17);
  body.appendParagraph("");

  body.appendParagraph(formattedDate + "\n\n").setBold(true).setUnderline(true).setAlignment(DocumentApp.HorizontalAlignment.CENTER)
      .setFontFamily("Times New Roman").setFontSize(17);
  body.appendParagraph("");

    if (events.length > 0) {
      events.forEach(function(event) {
        body.appendParagraph(formatEvent(event)).setFontFamily("Times New Roman").setFontSize(15).setBold(false).setUnderline(false);
      });
    } else {
      body.appendParagraph("\n\nNo meetings scheduled today")
          .setAlignment(DocumentApp.HorizontalAlignment.CENTER)
          .setFontFamily("Times New Roman").setFontSize(16).setBold(false).setUnderline(false);
    }
  return formattedDate;
}

function isAfterTime() {
  var now = new Date();
  
  var currentHour = now.getHours();
  return currentHour >= TIME_TO_UPDATE_TO_TMR;
}


function getTomorrowEvents(calendarId) {
  var tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1); // Set to tomorrow's date
  
  var startOfDay = new Date(tomorrow);
  startOfDay.setHours(0, 0, 0); // Set to midnight
  var endOfDay = new Date(tomorrow);
  endOfDay.setHours(23, 59, 59); // Set to 11:59 PM

  var calendar = CalendarApp.getCalendarById(calendarId);
  var events = calendar.getEvents(startOfDay, endOfDay);
  return events;
}

function getTodayEvents(calendarId) {
  var today = new Date();
  
  var startOfDay = new Date(today);
  startOfDay.setHours(0, 0, 0);
  var endOfDay = new Date(today);
  endOfDay.setHours(23, 59, 59);

  var calendar = CalendarApp.getCalendarById(calendarId);
  var events = calendar.getEvents(startOfDay, endOfDay);
  return events;
}

function formatEvent(event) {
  var startTime = Utilities.formatDate(event.getStartTime(), Session.getScriptTimeZone(), "hh:mm a");
  var endTime = Utilities.formatDate(event.getEndTime(), Session.getScriptTimeZone(), "hh:mm a");
  var eventTitle = event.getTitle();
  return startTime + " - " + endTime + "\t\t" + eventTitle + "\n";
}
