function onOpen(e) {
  SpreadsheetApp.getUi() //returns the instance of sheet's UI which lets you add menus etc
    .createAddonMenu()
    .addItem('Do Mail Merge', 'sendEmails')
    .addItem('Get Drafts', 'getDrafts')
    .addItem('Import Contacts', 'importContacts')
    .addToUi();
}

function onInstall(e) {
  onOpen(e);
}

/*var sheet,
  data = null,
  numRows = 0,
  numColumns = 0,
  columnHeader,
  status = 0,
  drafts,
  email = 0;
*/

function sendEmails() {
  var sheet,
    data = null,
    numRows = 0,
    numColumns = 0,
    status = -1,
    drafts,
    email = -1;

  //getDrafts();
  drafts = GmailApp.getDrafts();
  if (!drafts[0])
  {
    SpreadsheetApp.getUi()
      .alert('No mails to be sent');
    return;
  }
  var message;
  for (var i = 0; i < drafts.length; i++) {
    message = drafts[i].getMessage();
    //Logger.log(message.getSubject());
    //console.log('blah');
    //SpreadsheetApp.getUi()
      //.alert(message.getSubject());
  }

  //getData();
  sheet = SpreadsheetApp.getActiveSheet(); //.getSheets()[0];
  var startRow = 1,
    startColumn = 1;
  numRows = sheet.getLastRow(),
  numColumns =  sheet.getLastColumn(),
  //SpreadsheetApp.getUi()
    //.alert(numRows);
  Logger.log(numRows);
  var dataRange = sheet.getRange(startRow, startColumn, numRows, numColumns);
  data = dataRange.getValues();

  //getColumnHeaders();
  var columnHeader = new Array(numColumns + 2)
  for (var i = 0; i < numColumns; i++) {
    //SpreadsheetApp.getUi()
      //.alert(data[0][i]);
    columnHeader[i] = data[0][i];
    if (data[0][i] == 'STATUS') { //can I do this
      status = i;
    }
    if (data[0][i] == 'EMAIL') { //can I do this
      email = i;
    }
  }

  //if(emailNotSent()) {
  if (data == null || email == -1) {
    SpreadsheetApp.getUi()
      .alert('No email address given');
    return;
  }
  if (status == -1 && numColumns > 1) {
    sheet.insertColumnAfter(numColumns);
    sheet.insertColumnAfter(numColumns);
    sheet.getRange(1, numColumns + 1).setValue('STATUS');
    status = numColumns;
    sheet.getRange(1, numColumns + 2).setValue('Sent Email STATUS');
    numColumns = numColumns + 2;
  }
  else {
    var new_mail = 0;
    for (var i = 2; i < numRows; i++) {
      if (sheet.getRange(i, numColumns - 1).getValue() == '') {
        new_mail = 1;
        break;
      }
    }
    if (new_mail == 0) {
      SpreadsheetApp.getUi()
        .alert('No new email to be sent');
      return;
    }
  }

  // Fetch values for each row in the Range.
  var subject = drafts[0].getMessage().getSubject();
    for (var i = 1; i < data.length; ++i) {
      var row = data[i];
      var emailAddress = row[email];
      var message = drafts[0].getMessage().getBody();
      SpreadsheetApp.getUi()
        .alert(message);
      var emailSent = row[status];
      for (var j = 0; j < numColumns - 2; j++) {
        message = message.replace('{{' + columnHeader[j] + '}}', row[j]);
        SpreadsheetApp.getUi()
        .alert(message);
      }
      if (emailSent != 'EMAIL_SENT') {
        GmailApp.sendEmail(emailAddress, subject, message, {htmlBody : message});
        sheet.getRange(startRow + i, status + 1).setValue('EMAIL_SENT');
        // Make sure the cell is updated right away in case the script is interrupted
        SpreadsheetApp.flush();
      }
    }
}

/*function getData() {
  sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 1,
    startColumn = 1;
  numRows = sheet.getLastRow(),
  numColumns =  sheet.getLastColumn(),
  var dataRange = sheet.getRange(startRow, startColumn, numRows, numColumns);
  data = dataRange.getValues();
}

function getColumnHeaders() {
  for (var i = 0; i < numColumns; i++) {
    columnHeader[i] = data[0][i];
    if (data[0][i] == 'STATUS') { //can I do this
      status = i;
    }
    if (data[0][i] == 'EMAIL') { //can I do this
      email = i;
    }
  }
}

function emailNotSent() {
  if (data == null || EMAIL == 0) {
    SpreadsheetApp.getUi()
      .alert('No email address given');
    return false;
  }
  if (status == 0) {
    sheet.insertColumnfter(numColumns);
    sheet.insertColumnAfter(numColumns);
    sheet.getRange(0, numColumns + 1).setValue('STATUS');
    sheet.getRange(0, numColumns + 2).setValue('Sent Email STATUS');
    numColumns = numColumns + 2;
    return true;
  }
  else {
    var new_mail = 0;
    for (var i = 1; i < numRows; i++) {
      if (sheet.getRange(i, numColumns - 1).getValue() == '') {
        new_mail = 1;
        break;
      }
    }
    if (new_mail == 0) {
      SpreadsheetApp.getUi()
        .alert('No new email to be sent');
      return false;
    }
    else {
      return true;
  }
}*/

function getDrafts() {
  drafts = GmailApp.getDrafts();
  if (!drafts[0])
  {
    SpreadsheetApp.getUi()
      .alert('No mails to be sent');
    return;
  }
  var message;
  for (var i = 0; i < drafts.length; i++) {
    message = drafts[i].getMessage();
    //Logger.log(message.getSubject());
    //console.log('blah');
    SpreadsheetApp.getUi()
      .alert(message.getSubject());
  }
}

function importContacts() {

}
