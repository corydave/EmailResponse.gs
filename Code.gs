var ss = SpreadsheetApp.getActive();
var sheet = ss.getSheetByName('COMMUNICATION');

function onFormSubmit(e) {

  try {
    
    Logger.log('submit triggered');
    
    var rowNumber = SpreadsheetApp.getActive().getSheetByName('Form Responses 1').getLastRow();
    
    // CHANGE BASED ON CONTEXT
    
    var timeStamp = e.values[0];
    var firstName = e.values[1];
    var lastName = e.values[2];
    var pronoun = e.values[3];
    var institution = e.values[4];
    var email = e.values[5];  
    var volunteer = e.values[6];
    var accessibility = e.values[7];
    
    
    // NEED BOTH - 'messageHTML' IS PREFERRED, BUT 'message' IS FALLBACK
    var message = generateMessage(firstName, lastName, institution, email, accessibility, pronoun);
    var messageHTML = generateMessageHTML(firstName, lastName, institution, email, accessibility, pronoun);
    
    sendEmail(message, email, rowNumber, messageHTML);


  } catch (err) {
    
    Logger.log('onFormSubmit(e) encountered an error: ' + err + ' at ' + err.lineNumber);

  }

}

function sendEmail(message, email, rowNumber, messageHTML) {

    try {
    
      // GmailApp.sendEmail(email,'CCCAT Conference - 2022', message);

        MailApp.sendEmail({
          to: email,
          subject: "CCCAT Conference - 2022",
          body: message,
          htmlBody: messageHTML
        });

      sheet.getRange(rowNumber,9).setValue(new Date());
      sheet.getRange(rowNumber,10).setValue(new Date());
      Logger.log(message);

      
      
    } catch (err) {
      
      Logger.log('ERROR\nTried to email ' + users[i][5] + ' but there was an issue. ' + err.lineNumber);

    }



}

  function generateMessageHTML(firstName, lastName, institution, email, accessibility, pronoun) {

    try {

      var messageHTML = '';
      messageHTML = 'Dear ' + firstName + ' ' + lastName + ',';
      messageHTML += '<br /><br />Thank you for registering for the 2022 CCCAT Conference!';
      messageHTML += '<br /><br />';
      messageHTML += 'You can find the schedule at <a href="https://www.cccatconference.org/schedule">here</a>.';
      messageHTML += '<br /><br />';
      messageHTML += 'A few things you should keep in mind for the conference:';
      messageHTML += '<ul><li>We will send you another email as the conference approaches.</li>';
      messageHTML += '<li>If you have any issues, you can email us at <a href="mailto:cccatconference@gmail.com">cccatconference@gmail.com</a></li>';
      messageHTML += '<li>Any information that changes on the day of the conference (schedule changes, for instance) will be posted on the "Information" tab at the website</li>';
      messageHTML += '<li>This conference is going to be awesome. Buckle your seatbelts!</li></ul>';
      messageHTML += '<br /><br />Celebrating the Scholarship of Teaching and Learning,';
      messageHTML += '<br />The CCCAT Conference Committee';
      messageHTML += '<br />';
      messageHTML += '<a href="https://www.cccatconference.org/">www.cccatconference.org</a>';

      return messageHTML;
      
    } catch (err) {

      Logger.log('generateMessageHTML() - encountered an error: ' + err + ' at ' + err.lineNumber);

    }

  }

  function generateMessage(firstName, lastName, institution, email, accessibility, pronoun) {

    try {
      var message = '';
      message = 'Dear ' + firstName + ' ' + lastName + ',';
      message += '\n\nThank you for registering for the 2022 CCCAT Conference!';
      message += '\n\n';
      message += 'You can find the schedule at https://www.cccatconference.org/schedule';
      message += '\n\n';
      message += 'A few things you should keep in mind for the conference:\n\n';
      message += ' - We will send you another email as the conference approaches.\n\n';
      message += ' - If you have any issues, you can email us at cccatconference@gmail.com\n\n';
      message += ' - Any information that changes on the day of the conference (schedule changes, for instance) will be posted on the "Information" tab at the website\n\n';
      message += ' - This conference is going to be awesome. Buckle your seatbelts!';
      message += '\n';
      message += '\n';
      message += '\n\nCelebrating the Scholarship of Teaching and Learning,';
      message += '\nThe CCCAT Conference Committee';
      message += '\n';
      message += 'https://www.cccatconference.org/';    

      return message;
      
    } catch (err) {

      Logger.log('generateMessage() encountered an error: ' + err + ' at ' + err.lineNumber);

    }



  }




function test() {

  try {

  Logger.log('test');
  
  } catch (err) {

    Logger.log('There was an error with "test()" on line ' + err.lineNumber);

  }
}

function emailPeople() {

  var numOfUsers = sheet.getLastRow();
  numOfUsers = numOfUsers - 1;
  
  var startRow = 103;
  var startCol = 12;
  var numDataCols = 9; 
  var dateCol = 17;
  var numRows = 100;

  var users = sheet.getRange(startRow, startCol, numRows, numDataCols).getValues();
  
  var message = '';
  // var userData = [];
  

  for (var i = 0; i < numRows; i++) {
    
    var firstName = users[i][0]; 
    message = generateMessage(firstName);
    
  
    try {
      
      GmailApp.sendEmail(users[i][3],'CCCAT Conference - THIS FRIDAY',message);
      sheet.getRange(i+2,dateCol).setValue(new Date());
      // sheet.getRange(i+2,10).setValue(new Date());
      Logger.log(message + '\nemail=' + users[i][3]);
      
    } catch (err) {
      
      Logger.log('ERROR\nTried to email ' + users[i][3] + ' but there was an issue. ' + err.lineNumber);

    }

    
  
  }




  // Timestamp | First |	Last | College | Email | Host | Accessibility requests | Pronoun

}








