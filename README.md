# google_docs
push_to_calendar.gs
Creates calendar events based on a Google Sheet.
1. Create a Google Sheet with these headers: "Offer Status	Name	Start Date	Position	Company	Location	In Calendar?"
2. The section in the code called "//calendar variables" needs to be updated with the calendar's addresses in which are to be updated.
3. The section in the code called "//Choose the right calendar for the right City" needs to be updated with the expected text in the Location column of the spreadsheet. Example:
case "San Francisco":
  var calendar = CalendarApp.getCalendarById(SF);
  break;
case "Emerald City":
  var calendar = CalendarApp.getCalendarById(EC); where EC = the variable in "//calendar variables" that corresponds to the Emerald City calendar
  break;
default:
  var calendar = CalendarApp.getCalendarById(REMOTE)
  remote_on = true;
  break;
4. Save the script and refresh the Google Sheet.  A new menu will appear next to the Help menu called "Calendar"
5. The first time you run it, Google will ask for permission.
