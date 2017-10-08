function populateCalendar() {
  var data = getCalenderData();
  addDataToSheet(data);
  var h = "h"
}

var getCalenderData = function(){
  var calendars = CalendarApp.getAllCalendars();
  var events = [];
  var now = new Date();
  var oneYearAgo = new Date(now.getTime() + (60 * 60 * 24 * 365 * 1000));

  calendars.forEach(function(calendar, i){
    cEvents = calendar.getEvents(now, oneYearAgo);
    events.push(cEvents);
  });

  var rows = [["Active", "Start Time", "End Time", "Title", "Description", "Location"]];
  events.forEach(function(event, i){
    event.forEach(function(e, i){
      var row = [];
      var title = e.getTitle();
      row.push("NO")
      row.push(Utilities.formatDate(e.getStartTime(), 'America/Denver', 'MMMM dd, yyyy HH:mm'));
      row.push(Utilities.formatDate(e.getEndTime(), 'America/Denver', 'MMMM dd, yyyy HH:mm'));
      row.push(title);
      row.push(e.getDescription());
      row.push(e.getLocation());

      if(shouldAdd(title)) rows.push(row);
    })
  });

  return rows
}

var shouldAdd = function(title){
  var badWords =  [
    "PEC",
    "Ward Council",
    "Blom @ Mutual",
    "BYC",
    "Church ",
    "Church",
    "Scout Roundtable Mtg.",
    "Priest Presidency Meeting",
    "Key Scouter Meeting",
    "EQ/HP PPI's",
    "R2 Men's Basketball",
    "Stake High Priest Meeting",
    "Womens BBall",
    "Stake Priesthood Leadership",
    "Stake Council Meeting",
    "Stake Baptisms",
    "High Council Sunday",
    "Stake PEC"
  ];

  if (badWords.indexOf(title) > -1 || title.indexOf("ishop") !== -1) {
    return false;
  } else {
    return true;
  }
}

var addDataToSheet = function(data){
  var ss = SpreadsheetApp.openById('1K_hK9Q8AOSfHxZn5Uu-rKDRwBR3QboKRT1mTUtp2o_I');
  var as = ss.getSheetByName('Calendar');

  as.clearContents();

  data.forEach(function(d){
    as.appendRow(d);
  });

  as.setFrozenRows(1);
  as.sort(2);
}
