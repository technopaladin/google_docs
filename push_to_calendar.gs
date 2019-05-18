//push new events to calendar
function pushToCalendar() {
  //spreadsheet variables
  var sheet = SpreadsheetApp.getActiveSheet();
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();
  var range = sheet.getRange(1,1,lastRow,lastColumn);
  var values = range.getValues();
  var start_date = 0;
  var offer_status = 1;
  var name = 3;
  var location = 4;
  var company = 5;
  var position = 6;
  var calendar_col = 22;
  var remote_on = false;
  var yes = "y";

  //calendar variables
  var CHI = 'CHI@group.calendar.google.com';
  var SF = 'SF@group.calendar.google.com';
  var NY = 'NY@group.calendar.google.com';
  var REMOTE = 'Remote@group.calendar.google.com';

  //Calendar Control Column
  for (var i = 0; i <= lastColumn; i++) {
    if (values[0][i] == "In Calendar?")
      calendar_col = i;
  }
  //Start Date
  for (var i = 0; i <= lastColumn; i++) {
    if (values[0][i] == "Start Date")
      start_date = i;
  }
  //Offer Status
  for (var i = 0; i <= lastColumn; i++) {
    if (values[0][i] == "Offer Status")
      offer_status = i;
  }
  //Name
  for (var i = 0; i <= lastColumn; i++) {
    if (values[0][i] == "Name")
      name = i;
  }
  //Location
  for (var i = 0; i <= lastColumn; i++) {
    if (values[0][i] == "Location")
      location = i;
  }
  //Company
  for (var i = 0; i <= lastColumn; i++) {
    if (values[0][i] == "Company")
      company = i;
  }
  //Position
  for (var i = 0; i <= lastColumn; i++) {
    if (values[0][i] == "Position")
      position = i;
  }

  //Choose the right calendar for the right City
  for (var i = 0; i < values.length; i++) {
    if ((values[i][calendar_col] != yes)  && !(isNaN(values[i][start_date]))) {
      if ((values[i][name].length > 0) && (values[i][offer_status] == "Accepted") && (values[i][start_date] != "")) {
        switch (values[i][location])  {
          case "Chicago":
            var calendar = CalendarApp.getCalendarById(CHI);
            break;
          case "New York":
            var calendar = CalendarApp.getCalendarById(NY);
            break;
          case "San Francisco":
            var calendar = CalendarApp.getCalendarById(SF);
            break;
          default:
            var calendar = CalendarApp.getCalendarById(REMOTE)
            remote_on = true;
            break;
        }
      if (values[i][calendar_col] != yes) {
        //create event https://developers.google.com/apps-script/class_calendarapp#createEvent
        if (remote_on) {
            var newEventTitle = values[i][name] + ' - ' + values[i][company] + " " + values[i][position] + ' - ' + values[i][location];
        } else {
            var newEventTitle = values[i][name] + ' - ' + values[i][company] + " " + values[i][position];
            }
        }
        var newEvent = calendar.createAllDayEvent(newEventTitle, values[i][start_date],{location: values[i][location]});
        //calendar.setTimeZone("America/New_York")
        //mark as entered
        cell = sheet.getRange(i+1,calendar_col+1);
        cell.setValue(yes);
      }
    }
  }
  Browser.msgBox('Calendars Updated');
}

//add a menu when the spreadsheet is opened
function onOpen() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet();
  var menuEntries = [];
  menuEntries.push({name: "Update Calendar", functionName: "pushToCalendar"});
  sheet.addMenu("Calendar", menuEntries);
}
