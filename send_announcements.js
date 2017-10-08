function sendAnnouncements() {
  var sheets = getSheets();
  var announcements = getAnnouncements(sheets['announcementSheet']);
  var calendar = getCalenderItems(sheets['calendarSheet']);
  var missionaries = getMissionaries(sheets['missionarySheet']);
  var memberContactInfo = getMemberContactInfo(sheets['contactSheet']);
  var directory = getDirectory(memberContactInfo, sheets['callingSheet']);
  var memberEmails = getMemberEmail(memberContactInfo);
  sendEmail(memberEmails, announcements, calendar, missionaries, directory)
}

var getSheets = function(){
  var ss = SpreadsheetApp.openById('1K_hK9Q8AOSfHxZn5Uu-rKDRwBR3QboKRT1mTUtp2o_I');

  return {
    'calendarSheet': ss.getSheetByName('Calendar'),
    'announcementSheet': ss.getSheetByName('Announcements'),
    'memberSheet': ss.getSheetByName('Members'),
    'missionarySheet': ss.getSheetByName('Missionaries'),
    'callingSheet': ss.getSheetByName('Callings'),
    'contactSheet': ss.getSheetByName('Contacts')
  }
}

var getAnnouncements = function(announcementSheet){
  var announcements = [];
  var range = announcementSheet.getDataRange();
  var numRows = range.getNumRows();
  var data = announcementSheet.getRange(2, 1, numRows, 2).getValues();
  data.forEach(function(d){
    if(d[0] == "YES") announcements.push(d[1])
  })

  return announcements
}

var getCalenderItems = function(calendarSheet){
  var announcements = [];
  var range = calendarSheet.getDataRange();
  var numRows = range.getNumRows();
  var data = calendarSheet.getRange(2, 1, numRows, 6).getValues();
  data.forEach(function(d){
    if(d[0] == "YES") announcements.push([d[1], d[2], d[3], d[4], d[5]])
  })

  return announcements
}

var getMissionaries = function(missionarySheet){
  var missionaries = [];
  var range = missionarySheet.getDataRange();
  var numRows = range.getNumRows();
  var missionaries = missionarySheet.getRange(2, 1, numRows, 8).getValues();

  return missionaries
}

// Gets all the members contact information from google contacts. Return format
// [{
//   "phone": "",
//   "email": "",
//   "name": ""
// }]
var getMemberContactInfo = function(contactSheet){
  var contacts = contactSheet.getSheetValues(1, 1, 500, 3)
  var contactData = [];

  contacts.forEach(function(contact){
    var phone = contact[0]
    var email = contact[1]
    var name = contact[2]

    contactData.push({
      "phone": phone,
      "email": email,
      "name": name
    })
  })

  return contactData
}

var sortUnique = function(arr) {
    if (arr.length === 0) return arr;
    arr = arr.sort(function (a, b) { return a*1 - b*1; });
    var ret = [arr[0]];
    for (var i = 1; i < arr.length; i++) { // start loop at 1 as element 0 can never be a duplicate
        if (arr[i-1] !== arr[i]) {
            ret.push(arr[i]);
        }
    }

    return ret;
}

var getMemberEmail = function(contacts){
  var emails = [];
  contacts.forEach(function(contact){
    if(contact["email"].length > 0){
      emails.push(contact["email"])
    }
  })

  var uniqueEmails = sortUnique(emails)
  var shortenedEmails = uniqueEmails.slice(0, 98)
  var uniqueSize = shortenedEmails.length
  var uniqueList = [];
  var grouped = [];
  var count = 0;
  shortenedEmails.forEach(function(email, i){
    grouped.push(email);
    count = count + 1

    if(count > 48 || (i + 1) == uniqueSize){
      uniqueList.push(grouped.join(', '));
      count = 0;
      grouped = [];
    }
  })

  return uniqueList
}

var getDirectory = function(memberContactInfo, callingSheet){
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
    {"Other": "Ward Bullitin Coordinator"}
  ]

  var callingValues = callingSheet.getRange(1, 1, 300, 3).getValues()
  var leadership = getMemberInCalling(positions, callingValues)
  var directoryData = getDirectoryHash(memberContactInfo, leadership)

  return order(directoryData, positions)
}

// Accumulates a hash of information for the leadership directory. Return Format
// [{
//   "type": "",
//   "calling": "",
//   "name": "",
//   "email": "",
//   "phone":
// }]
var getDirectoryHash = function(memberValues, leadership){
  var directory = []
  memberValues.forEach(function(member){
    if(member["name"].length > 0){
      leadership.forEach(function(leader){
        if(member["name"] == leader["name"]){
          var cData = {
            "type": leader["organization"],
            "calling": leader["calling"],
            "name": leader["name"],
            "email": member["email"],
            "phone": member["phone"]
          };

          directory.push(cData)
        };
      });
    };
  });

  return directory
}

// Gets the members of the callings sheet that hold the specified positions. Return format
// [{
//   "name": "",
//   "organization": "",
//   "calling": ""
// }]
var getMemberInCalling = function(positions, callingValues){
  var leadership = []
  callingValues.forEach(function(calling){
    positions.forEach(function(position){
      var organization = Object.keys(position)
      var posCalling = position[organization]
      if(calling[0] == organization && calling[1] == posCalling){
        leadership.push({
          "name": calling[2],
          "organization": calling[0],
          "calling": calling[1]
        })
      }
    })
  })

  return leadership
}

var order = function(contactData, positions){
  ordered = []

  positions.forEach(function(category){
    contactData.forEach(function(contact){
      var calling = ""
      var area = ""
      for(var key in category) {
        calling = category[key]
        area = key
      }

      if(contact["calling"] == calling && key == contact["type"]){
        ordered.push(contact)
      }
    });
  });

  return ordered
}

var sendEmail = function(emails, announcements, calendar, missionaries, directory){
  var text = "<div style='font-family: sans-serif;'><div>Dear Meadows Ward Member,</h3><br/><br/>Below you will find an automated report. For assistance with the contents, please contact a member of the bishopric.<br/>For any technical report issues, please contact Mike Blom (mikewblom@gmail.com)</div><br/>"

  if(announcements.length > 0){
    text += "<div style='background-color: #3C4C69; padding: 10px;'><div style='font-size: 20px; font-weight: 100; color: white;'>Announcements</div></div>";
    text += "<ul style='list-style-type: circle; background-color: white; padding: 10px 0px 10px 26px; margin-top: 0px; border: 1px solid #3C4C69;'>";
    announcements.forEach(function(announcement){
      text += "<li>";
      text += announcement;
      text += "</li>";
    });
    text += "</ul><br/>";
  }

  if(calendar.length > 0){
    text += "<div style='background-color: #3C4C69; padding: 10px;'><div style='font-size: 20px; font-weight: 100; color: white;'>Calendar</div></div><table style='border: 1px solid #3C4C69; width: 100%;'><tr style='background-color: brown; color: white; font-weight: 100;'> <th style='border: 1px solid #3C4C69;'>Start Date and Time</th> <th style='border: 1px solid #3C4C69;'>End Date and Time</th> <th style='border: 1px solid #3C4C69;'>Title</th> <th style='border: 1px solid #3C4C69;'>Description</th> <th style='border: 1px solid #3C4C69;'>Location</th> </tr>";

    calendar.forEach(function(item){
      text += "<tr style='background-color: white;'>";
      text += "<td style='min-width: 100%; border: 1px solid #3C4C69;'>" + Utilities.formatDate(item[0], 'America/Denver', 'MMMM dd, HH:mm') + "</td>";
      text += "<td style='min-width: 100%; border: 1px solid #3C4C69;'>" + Utilities.formatDate(item[1], 'America/Denver', 'MMMM dd, HH:mm') + "</td>";
      text += "<td style='max-width: 200px; border: 1px solid #3C4C69;'>" + checkText(item[2]) + "</td>";
      text += "<td style='max-width: 200px; border: 1px solid #3C4C69;'>" + checkText(item[3]) + "</td>";
      text += "<td style='max-width: 100px; border: 1px solid #3C4C69;'>" + checkText(item[4]) + "</td>";
      text += "</tr>";
    })

    text += "</table><br/>";
  }

  if(missionaries.length > 0){
    text += "<div style='background-color: #3C4C69; padding: 10px;'><div style='font-size: 20px; font-weight: 100; color: white;'>Members of the ward serving missions</div></div>";
    text += "<ul style='list-style-type: circle; background-color: white; padding: 10px 0px 10px 26px; margin-top: 0px; border: 1px solid #3C4C69;'>";
    missionaries.forEach(function(missionary){
      if(missionary[0].length > 0){
        var address = "<div>" + missionary[2] + "<br/>" + missionary[3] + ", " + missionary[4] + ", " + missionary[5] + "<br/>" + missionary[6] + "</div>"
        text += "<li>";
        text += "<b>Name:</b> " + missionary[0] + "<br/>";
        text += "<b>Mission:</b> " + missionary[1] + "<br/>";
        text += "<b>Address:</b> " + address;
        text += "<b>Email:</b> " + missionary[7]
        text += "</li><br/>";
      }
    });
    text += "</ul><br/>";
  }

  if(directory.length > 0){
    text += "<div style='background-color: #3C4C69; padding: 10px;'><div style='font-size: 20px; font-weight: 100; color: white;'>Directory</div></div><table style='border: 1px solid #3C4C69; width: 100%;'><tr style='background-color: brown; color: white; font-weight: 100;'> <th style='border: 1px solid #3C4C69;'>Category</th> <th style='border: 1px solid #3C4C69;'>Calling</th> <th style='border: 1px solid #3C4C69;'>Name</th> <th style='border: 1px solid #3C4C69;'>Phone Number</th> <th style='border: 1px solid #3C4C69;'>Email</th> </tr>";

    directory.forEach(function(item){
      text += "<tr style='background-color: white;'>";
      text += "<td style='min-width: 100%; border: 1px solid #3C4C69;'>" + checkText(item["type"]) + "</td>";
      text += "<td style='min-width: 100%; border: 1px solid #3C4C69;'>" + checkText(item["calling"]) + "</td>";
      text += "<td style='max-width: 200px; border: 1px solid #3C4C69;'>" + checkText(item["name"]) + "</td>";
      text += "<td style='max-width: 200px; border: 1px solid #3C4C69;'>" + checkText(item["phone"]) + "</td>";
      text += "<td style='max-width: 100px; border: 1px solid #3C4C69;'>" + checkText(item["email"]) + "</td>";
      text += "</tr>";
    })

    text += "</table><br/>";
  }

  text += "We look forward to seeing you on Sunday,<br/><br/>Meadows Ward Bishopric"

  if(announcements.length > 0 || calendar.length > 0){
    emails.forEach(function(groupedEmails){
      Logger.log("Sending emails to: " + groupedEmails)
      Logger.log("Remaining: " + MailApp.getRemainingDailyQuota())
      //"mikewblom@gmail.com, mb@masteryconnect.com, bart.brisko@gmail.com, rrrramon4@gmail.com, kmbeddes@gmail.com, asessions6@gmail.com, farmingtonmeadows@gmail.com"
      MailApp.sendEmail({
        to: 'farmingtonmeadows@gmail.com',
        subject: "Meadows Ward Information",
        htmlBody: text,
        bcc: groupedEmails
      });
    });
  }
}
var checkText = function(text){
  if(text == undefined){
    return "N/A"
  }else{
    return text
  }
}
