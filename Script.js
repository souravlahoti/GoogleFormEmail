/* Send Confirmation Email with Google Forms */
 
function Initialize() {
 
    var triggers = ScriptApp.getProjectTriggers();
 
    for (var i in triggers) {
        ScriptApp.deleteTrigger(triggers[i]);
    }
 
    ScriptApp.newTrigger("SendConfirmationMail")
        .forSpreadsheet(SpreadsheetApp.getActiveSpreadsheet())
        .onFormSubmit()
        .create();
 
}
 
function SendConfirmationMail(e) {
    
    try {
        
        var ss, cc, sendername, subject, columns;
        var message, value, textbody, sender;
        var flag =1;
        var sheet = SpreadsheetApp.getActiveSheet();
        var data = sheet.getDataRange().getValues();
        var newData = new Array();
        for(i in data){
        var row = data[i];
        var duplicate = false;
        for(j in newData){
          if(row[2] == newData[j][2]){
            duplicate = true;
            flag= 0;
          }
        }
        if(!duplicate){
          newData.push(row);
        }
      }
      sheet.clearContents();
      sheet.getRange(1, 1, newData.length, newData[0].length).setValues(newData);
    
      
      
      
      
        if (flag ==1)
        {
        // This is your email address and you will be in the CC
        cc = Session.getActiveUser().getEmail();
        
        // This will show up as the sender's name
        sendername = "School's Out Team";
 
        // Optional but change the following variable
        // to have a custom subject for Google Docs emails
        subject = "Your Response Successfully Submitted";
 
        // This is the body of the auto-reply
        message = "We have received your details.<br>Thanks!<br><br>";
 
        ss = SpreadsheetApp.getActiveSheet();
        columns = ss.getRange(1, 1, 1, ss.getLastColumn()).getValues()[0];
 
        // This is the submitter's email address
        sender = e.namedValues["Email Address"].toString();
 
       /* // Only include form values that are not blank
        for ( var keys in columns ) {
            var key = columns[keys];
            if ( e.namedValues[key] ) {
                message += key + ' :: '+ e.namedValues[key] + "<br />"; 
            }
        }
        */
        
        textbody = message.replace("<br>", "\n");
 
        GmailApp.sendEmail(sender, subject, textbody, 
                            {cc: cc, name: sendername, htmlBody: message});
        }
      
      
      
      else
      {
        // This is your email address and you will be in the CC
        cc = Session.getActiveUser().getEmail();
        
        // This will show up as the sender's name
        sendername = "School's Out Team";
 
        // Optional but change the following variable
        // to have a custom subject for Google Docs emails
        subject = "Already Registered.";
 
        // This is the body of the auto-reply
        message = "You have already registered with us.<br>Thanks!<br><br>";
 
        ss = SpreadsheetApp.getActiveSheet();
        columns = ss.getRange(1, 1, 1, ss.getLastColumn()).getValues()[0];
 
        // This is the submitter's email address
        sender = e.namedValues["Email Address"].toString();
 
       /* // Only include form values that are not blank
        for ( var keys in columns ) {
            var key = columns[keys];
            if ( e.namedValues[key] ) {
                message += key + ' :: '+ e.namedValues[key] + "<br />"; 
            }
        }
        */
        
        textbody = message.replace("<br>", "\n");
 
        GmailApp.sendEmail(sender, subject, textbody, 
                            {cc: cc, name: sendername, htmlBody: message});
        }
      
    } catch (e) {
        Logger.log(e.toString());
    }
 
}