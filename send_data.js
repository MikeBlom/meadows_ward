function sendData(){
  var sacramentSubject = "Next 2 weeks of sacrament meetings";
  var data = getSacramentData();
  var sacramentData = data[0];
  var sacramentEmails = getSacramentEmails(data[1]);
  //var sacramentEmails = 'mikewblom@gmail.com'
  sendEmails(sacramentData, sacramentEmails, sacramentSubject);
}

var getSacramentEmails = function(names){
  var positions = [
    {"Bishopric": "Bishop"},
    {"Bishopric": "1st Counselor"},
    {"Bishopric": "2nd Counselor"},
    {"Bishopric": "Executive Secretary"},
    {"Bishopric": "Ward Clerk"},
    {"High Priests": "Group Leader"},
    {"Elders Quorum": "President"},
    {"Relief Society": "President"},
    {"Young Men": "President"},
    {"Young Women": "President"},
    {"Sunday School": "President"},
    {"Primary": "President"},
    {"Ward Mission": "Leader"},
    {"Music": "Chairman"},
    {"Music": "Director"},
    {"Music": "Organist"},
    {"Other": "Ward Bullitin Coordinator"}
  ]
  var ss = SpreadsheetApp.openById('1K_hK9Q8AOSfHxZn5Uu-rKDRwBR3QboKRT1mTUtp2o_I');
  var cs = ss.getSheetByName("Callings");
  var ms = ss.getSheetByName("Members");
  var callingValues = cs.getRange(1, 1, 300, 3).getValues()
  var memberValues = ms.getRange(1,1,500,2).getValues()
  var posNames = getMemberInCallings(positions, callingValues)
  var allNames = posNames.concat(names)

  return getMemberEmails(allNames, memberValues)
}

var getMemberEmails = function(names, members){
  emails = []
  members.forEach(function(member){
    names.forEach(function(name){
      if(name == member[0]){
        emails.push(member[1])
      }
    })
  })

  return emails.toString()
}

var getMemberInCallings = function(positions, callingValues){
  var names = []
  callingValues.forEach(function(calling){

    positions.forEach(function(position){
      var organization = Object.keys(position)
      var posCalling = position[organization]
      if(calling[0] == organization && calling[1] == posCalling){
        names.push(calling[2])
      }
    })

  })

  return names
}

var getSacramentData = function(){
  var sacramentEmails = "mikewblom@gmail.com";
  var ss = SpreadsheetApp.openById('1K_hK9Q8AOSfHxZn5Uu-rKDRwBR3QboKRT1mTUtp2o_I');
  var s = ss.getSheetByName("Sacrament");
  var range = s.getRange(2,1,100);
  var values = range.getValues();
  var now = Number(new Date());
  var rangeNums = [];
  values.forEach(function(date){
    rangeNums.push(Number(date[0]));
  })
  var closestDate = nextDate(rangeNums, now);
  var date = Object.keys(closestDate);
  var index = closestDate[date];
  var neededRange = s.getRange((index + 1), 1, 2, 14);
  var neededValues = s.getRange(1,1,1,14).getValues();
  var names = []
  neededRange.getValues().forEach(function(value, i){
    if(i == 1){
      value.forEach(function(v, i){
        if((i == 3 || i == 4 || i == 5 || i == 6 || i == 7 || i == 8 || i == 9) && v != "N/A" && v.length > 0){
          names.push(v)
        }
      })
    }
    neededValues.push(value);
  })

  return [neededValues, names]
}

var nextDate = function(dates, startDate) {
  var closest = Math.max.apply(null, dates); //Get the highest number in arr in case it match nothing.
  var index = 0;
  for(var i = 0; i < dates.length; i++){ //Loop the array
    if(dates[i] >= startDate && dates[i] < closest) {
      closest = dates[i]; //Check if it's higher than your number, but lower than your closest value
      index = i;
    }
  }

  return {closest: index}; // return the value
}

var sendEmails = function(values, emails, subject){
  var text = "<h3>Dear Member,</h3>Below you will find an automated report which shows our next 2 Sacrament Meetings (<span style='color:red'>Our upcoming Sacrament Meeting on the bottom row</span>).<br/>You are receiving this email because you are either speaking, praying or are in a calling involved in planning or performing in sacrament meeting.<br/>For assistance with prayers, please contact the Ward Executive Secretary.<br/> For assistance with songs and speakers, please contact a member of the bishopric.<br/>For any technical report issues, please contact Mike Blom (mikewblom@gmail.com)<br/><br/><table border='1'>"

  for (var row in values) {
    text += "<tr>";
    for (var col in values[row]) {
      var value = values[row][col];
      if(value.length == 0){
        value = "Not Yet Decided";
      }
      if(col == 0){
        value = "'" + value + "'";
        value = value.slice(0, -24);
      }

      if(row == 0){
        text += "<th style='min-width: 100px; color: blue;'>" + value + "</th>";
      } else {
        text += "<td style='min-width: 100px;'>" + value + "</td>";
      }
    }
    text += "</tr>";
  }
  text += "</table><br/><br/>Thanks for all you do,<br/><br/>Meadows Ward Bishopric";

  MailApp.sendEmail({
    to: emails,
    subject: subject,
    htmlBody: text
  })
}
