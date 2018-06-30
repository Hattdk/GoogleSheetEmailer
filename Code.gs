/**
 * Sends emails with data from the Tickets and Terms spreadsheet.
 * Recipient assigned in the Config sheet
 */


function sendEmails() {
  // open config sheet and pull send-to email address
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Config");
  var emailAddress = sheet.getRange(1,2,1).getValue();
    
  // promotions and job changes
  // Tickets sheet
  tickets(emailAddress);
  
  // terminations
  // Terms sheet
  terminations(emailAddress);
}


function tickets(emailAddress) {
  
  // open the tickets sheet where the report will get copied to
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Tickets");
  
  // start counting at 1 not zero for rows and cols
  var startRow = 2; // First row of data to process; 1 == headers
  var startCol = 1;
  var endCol = 29;
  var numRows = 50; // Number of rows to process
  
  // Fetch the range of cells 
  var dataRange = sheet.getRange(startRow, startCol, numRows, endCol);
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  
  // start looping through each row
  for (i in data) {
    // grab all row data
    var row = data[i];
    var empId = row[1];
    
    
    if (empId != "") {
      var firstName = row[2];
      var prefName = row[4];
      var lastName = row[3];
      var jobCode = row[19];
      var effDate = row[20];
      var effDate = Utilities.formatDate(effDate, "GMT", "MM-dd-yyyy");
      var event = row[21];
      var eReason = row[22];
      var store = row[26];
      var prevStore = row[28];
      if (prevStore == "")
        prevStore = "N/A";
      var msg;
      var body;
      var subject;
      
      // build out the base message
      msg = event + ": " + firstName;
      if (prefName != "")
        msg += " '" + prefName + "'";
      msg += " " + lastName;
      
      // create body using base message
      body = msg + "<br\>Job Change: " + eReason + "<br\>";
      body += "Job Code: " + jobCode + "<br\>";
      body += "Store: " + store + "<br\>";
      body += "Effective Date: " + effDate + "<br\>";
      body += "Previous Store: " + prevStore + "<br\>";
           
      // create subject using base message
      subject = "RMDC User - " + msg;
      sendEmail(emailAddress, subject, body);
      
//      Enter logic to filter users who need BI access
//      var subject = "Power BI User - " + msg;
//      sendEmail(emailAddress, subject, body);
    }
  }
}


function terminations(emailAddress) {
  sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Terms");
  var startRow = 2;
  var startCol = 1;
  var endCol = 13;
  var numRows = 50;
  
  var dataRange = sheet.getRange(startRow, startCol, numRows, endCol);
  var data = dataRange.getValues();
  
  for (i in data) {
    var row = data[i];
    var empId = row[1];
    
    if (empId != "") {
      var firstName = row[2];
      var prefName = row[3];
      var lastName = row[4];
      var jobTitle = row[5];
      var termDate = row[6];
      var termDate = Utilities.formatDate(termDate, "GMT", "MM-dd-yyyy");
      var store = row[10];
      var msg;
      var body;
      var subject;
      
      msg = "Termination: " + firstName;
      if (prefName != "")
        msg += " '" + prefName + "'";
      msg += " " + lastName;
      
      body = msg + "<br\>Job Code: " + jobTitle + "<br\>";
      body += "Term Date: " + termDate + "<br\>";
      body += "Store/Department: " + store + "<br\>";
      
      subject = "RMDC User - " + msg;
      sendEmail(emailAddress, subject, body);
      
//      Enter logic to filter users who need BI deprov
//      var subject = "RMDC User - " + msg;
//      sendEmail(emailAddress, subject, body);
    }
  }
}


function sendEmail(emailAddress, subject, body) {
  MailApp.sendEmail({
    to: emailAddress, 
    subject: subject, 
    htmlBody: body // needed to send html formated email
  });
}


/*
* adds the "Send Emails" button and link to sheet
*/
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [
    {name: 'Send Emails', functionName: 'sendEmails'}
  ];
  spreadsheet.addMenu('Send Emails', menuItems);
}