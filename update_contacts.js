function updateContacts() {
  contactData = getContactData();
  populateSheet(contactData);
}

var getContactData = function(){
  var contacts = ContactsApp.getContacts();
  var contactData = [];

  contacts.forEach(function(contact){
    var phones = contact.getPhones();
    var phone = (phones[0] == undefined) ? "" : phones[0].getPhoneNumber();


    var emails = contact.getEmails();
    var email = (emails[0] == undefined) ? "" : emails[0].getAddress();

    var name = contact.getFullName();

    contactData.push([phone, email, name]);
  })

  return contactData
}

var populateSheet = function(contactData){
  var ss = SpreadsheetApp.openById('1K_hK9Q8AOSfHxZn5Uu-rKDRwBR3QboKRT1mTUtp2o_I');
  var s = ss.getSheetByName('Contacts');
  s.clear();
  var contactNumber = contactData.length;
  var range = s.getRange("A1:C" + contactNumber)
  range.setValues(contactData)
  var h = "h"
}
