// sendEmailsMitoCircle.js - a short Google Apps script to email registrants from a form/spreadsheet in Google Apps

// string to let humans and computers know what happened
var EMAIL_SENT = "EMAIL_SENT";

function sendEmailsMitoCircle() {
  var sheet = SpreadsheetApp.getActiveSheet();
  // testing...
  var startRow = 2;  // First row of data to process
  var numRows = 1;   // Number of rows to process
 
  // Ended up not using the header bits in the email. 
  // Get header to print in email
  var colhead = sheet.getRange(1,1,1,13);
  // Get data cells that correspond to the header
  var dataRange = sheet.getRange(startRow, 1, numRows, 13);
  
  // Fetch values for each row in the Range.
  var data = dataRange.getValues();
  var colhead = colhead.getValues();
  
  // for each line in data
  for (var i = 0; i < data.length; ++i) {
    // the whole array contains the line so:
    var row = data[i];
    
    var emailAddress = row[11];  // email address 
    var emailSent = row[14];     // email flag 
    
    if (emailSent != EMAIL_SENT) {  // Prevents sending duplicates
      // Charming message:
      var subject = row[1]+" "+row[2]+" - Registration Confirmed for MitoCircle 2012";
      var message = "Thanks for registering! The MitoCircle Mitochondria and Metabolism conference is on June 14-16.\n\nPlease check our site for updates, to submit abstracts, as well as pay for registration (available soon!)\n\nhttp://www.mitocircle.org/conf2012";

      // Actually send it, turns out noReply only works under Google controlled domains :(
      MailApp.sendEmail(emailAddress, subject, message, {noReply: true});

      // Set flag
      sheet.getRange(startRow + i, 14).setValue(EMAIL_SENT);
      // Make sure the cell is updated right away in case the script is interrupted
      SpreadsheetApp.flush();  
    }
  }
}