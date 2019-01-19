
// runs when you open Google Sheets
function onOpen(e) {
  SpreadsheetApp.getUi()  // returns the instance of sheet's UI which lets you add menus etc
    .createAddonMenu()
    .addItem('Send Emails', 'sendEmails')  // to add a menu item for performing amil merge
    .addToUi();
}


// runs as soon as install the Add-on
function onInstall(e) {
  onOpen(e);
}


// function to send custom emails
function sendEmails() {

  // get drafts from Gmail
  var drafts = GmailApp.getDrafts();
  if (!drafts[0])  // check if drafts are present or not
  {
    SpreadsheetApp.getUi()
      .alert('No draft mail present');
    return;
  }

  // get data from the spreadsheet
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 1,
    startColumn = 1;
  var numRows = sheet.getLastRow();
  var numColumns =  sheet.getLastColumn();
  var data = null;
  //Logger.log(numRows);
  //SpreadsheetApp.getUi()
  //.alert(numRows);
  var dataRange = sheet.getRange(startRow, startColumn, numRows, numColumns);
  data = dataRange.getValues();

  // get the column names
  var columnHeader = new Array(numColumns + 1)
  var sent = -1, email = -1;
  for (var i = 0; i < numColumns; i++) {
    //SpreadsheetApp.getUi()
      //.alert(data[0][i]);
    columnHeader[i] = data[0][i];
    if (data[0][i] == 'SENT') {  // to get the sent column number
      sent = i;
    }
    if (data[0][i] == 'EMAIL') {  // to get the email column number
      email = i;
    }
  }

  // if there is no email column in the spreadsheet
  if (email == -1) {
    SpreadsheetApp.getUi()
      .alert('No email address is given.\nNo \'EMAIL\' column is present');
    return;
  }

  // if email address is given properly for all the inputs
  var re = new RegExp("^[a-zA-Z.]+@[a-zA-Z_]+\.[a-zA-Z]{2,3}$");
  for(var i = 1; i < numRows; i++) {
    if(!re.test(data[i][email])) {
      if(!data[i][email]) {
        SpreadsheetApp.getUi()
          .alert('cell (' + i + ', ' + email + ') : email address is not given');
        return;
      }
      SpreadsheetApp.getUi()
        .alert('cell (' + i + ', ' + email + ') : ' + data[i][email] + ' is not in proper format');
      return;
    }
  }

  // create a sent column if they are not present
  if (sent == -1) {
    sheet.insertColumnAfter(numColumns);
    sheet.getRange(1, numColumns + 1).setValue('SENT');
    sent = numColumns;
    numColumns = numColumns + 1;
  }

  // to check if all the mails are already sent or there are some new ones to be sent
  else {
    var new_mail = 0;
    for (var i = 2; i <= numRows; i++) {
      if (sheet.getRange(i, sent+1).getValue() == '') {
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

  // send emails
  var subject = drafts[0].getMessage().getSubject();
    for (var i = 1; i < data.length; ++i) {
      var row = data[i];
      var emailAddress = row[email];
      var message = drafts[0].getMessage().getBody();
      var emailSent = row[sent];
      for (var j = 0; j < numColumns - 1; j++) {
        message = message.replace('{{' + columnHeader[j] + '}}', row[j]);
      }
      if (emailSent != 'EMAIL_SENT') {
        SpreadsheetApp.getUi()
        .alert(message);
        GmailApp.sendEmail(emailAddress, subject, message, {htmlBody : message});
        var sentThread = GmailApp.search("is:sent", 0, 1)[0];
        var recipient = sentThread.getMessages()[0].getTo();
        if(recipient != row[email]) {
          SpreadsheetApp.getUi()
          .alert('Row : ' + i + ': \n not able to send email, please check the email address');
          return
        }
        sheet.getRange(startRow + i, sent + 1).setValue('EMAIL_SENT');
        SpreadsheetApp.flush();
      }
    }
}
