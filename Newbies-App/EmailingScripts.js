/**
 * Sends email.
 * @param {string} to - email address of a recipient;
 * @param {string} subject - subject of email message;
 * @param {string} body - body of email message.
 */
function sendEmail_(to, subject, body, bcc) {
  try {
    console.log(Utilities.formatString('sendEmail_ %s %s', to, subject));
    
    MailApp.sendEmail({
      to: to,
      subject: subject,
      htmlBody: body,
      noReply: true,
      bcc: bcc
    });
  } catch (e) {
    // Suppressing errors in email sending because email notifications are not critical for the functioning of the app.
    console.error(JSON.stringify(e));
  }
}


/**
* Generate a notification email to relevant companyName teams
*
* @param {recipients} recipients - object with 2 properties: from (requester email string), to (destination emails string)
* @param {string} eventType - type of event triggering the notification
* @param {data} data - data to be emailed. This object can have the following properties: {newbieObject (JS object), newbieRecord (AppMaker Datasource Record)} and any extra added properties (password)
*/
function generateEmail_(recipients, eventType, data){
  var subject, htmlBlob;
  
  switch(eventType) {
    case "newbieRequest":
      var startDate = new Date(data.newbieObject.startDate);
      
      subject = "[Newbie App] Newbie request: " + data.newbieObject.companyNameEmail + ' - starting on ' + startDate.toISOString().slice(0,10);
      
      htmlBlob = '<h1>Newbie request</h1>';
      htmlBlob += '<p><b>Name: </b>' + data.newbieObject.firstName + ' ' + data.newbieObject.lastName + '<br/>';
      htmlBlob += '<b>Email address: </b>' + data.newbieObject.companyNameEmail + '<br/>';
      htmlBlob += '<b>Start date: </b>' + startDate.toISOString().slice(0,10) + '</p>';
      htmlBlob += '<p><b>Requester: </b>' + data.newbieObject.requesterEmail + '<br/>';
      htmlBlob += '<b>Comments: </b>' + data.newbieObject.notes + '<br/><br/></p>';
      
      if (data.newbieObject.needsLaptop) {
        htmlBlob += '<p><b>This newbie requires a laptop, please make sure a machine is ready for them on their first day</b><br/></p>';
      }
      
      htmlBlob += '<p><b>Make sure you check/increase the license count through <a href="https://promevo-gpanelds.appspot.com/#Billing-Subscriptions">gPanel</a> before approving this request</b></p>';
      htmlBlob += '<span><a href="https://script.google.com/a/macros/companyName.com/s/AKfycbwKBg5YvcmUYJsf4YBW5cbN5e09NqqyDwHGOENTpkvUY52ddF8giwUzlrIx-dVS6NQL/exec#RequestsAdmin" style="height:36px;border-radius:2px;color:white;background-color:rgb(33,150,243);border:none;padding:8px 16px;text-decoration:none;" onclick="">VIEW IN APP</a></span>';
      
      break;
      
    case "newbieRequestV2":
      
      subject = "[Newbie App] Newbie request: " + data.newbieRecord.companyNameEmail + ' in ' + data.newbieRecord.Org + ' starting on ' + data.newbieRecord.StartDate.toISOString().slice(0,10);
      
      htmlBlob = '<h1>Newbie request</h1>';
      htmlBlob += '<p><b>Name: </b>' + data.newbieRecord.FirstName + ' ' + data.newbieRecord.LastName + '<br/>';
      htmlBlob += '<b>Email address: </b>' + data.newbieRecord.companyNameEmail + '<br/>';
      htmlBlob += '<b>Start date: </b>' + data.newbieRecord.StartDate.toISOString().slice(0,10) + '</p>';
      htmlBlob += '<p><b>Requester: </b>' + data.newbieRecord.RequesterEmail + '<br/>';
      htmlBlob += '<b>Comments: </b>' + data.newbieRecord.Notes + '<br/><br/></p>';
      
      if (data.newbieRecord.NeedsLaptop) {
        htmlBlob += '<p><b>This newbie requires a laptop, please make sure a machine is ready for them on their first day</b><br/></p>';
      }
      
      htmlBlob += '<p><b>Make sure you check/increase the license count through <a href="https://promevo-gpanelds.appspot.com/#Billing-Subscriptions">gPanel</a> before approving this request</b></p>';
      htmlBlob += '<span><a href="https://script.google.com/a/macros/companyName.com/s/AKfycbwKBg5YvcmUYJsf4YBW5cbN5e09NqqyDwHGOENTpkvUY52ddF8giwUzlrIx-dVS6NQL/exec#RequestsAdminV2" style="height:36px;border-radius:2px;color:white;background-color:rgb(33,150,243);border:none;padding:8px 16px;text-decoration:none;" onclick="">VIEW IN APP</a></span>';
      
      break;
      
    case "welcomeNewbie":
      subject = "Welcome to companyName IT systems";
      
      htmlBlob =  '<p>Hey ' + data.newbieRecord.FirstName + ', welcome to companyName!</p>';
      htmlBlob += '<p>Here\'s some account info to get to started with the basic IT systems:</p>';
      htmlBlob += '<p>Sign in to G-Suite <a href="https://www.google.com/accounts/AccountChooser?Email=' + data.newbieRecord.companyNameEmail + '&continue=https://apps.google.com/user/hub">here</a></p>';
      htmlBlob += '<p>Username: ' + data.newbieRecord.companyNameEmail + '</p>';
      htmlBlob += '<p>Password: ' + data.password + '</p>';
      htmlBlob += '<p>Your team will go over basics with you on your first day. Here are some  headlines:</p>';
      htmlBlob += '<ul>';
      htmlBlob +=   '<li>Google will ask for a new password when you first login.</li>';
      htmlBlob +=   '<li>You must enable <a href="https://www.google.com/landing/2step/">2-step verification</a> before all services can be enabled (and we strongly advise you turn it on anywhere it\'s available eg. GitHub).</li>';
      htmlBlob +=   '<li>Ask your team mates about which mailing lists and Slack groups you should be in.';
      htmlBlob +=     '<ul>';
      htmlBlob +=        '<li>Check <a href="groups.companyName.com">groups.companyName.com</a> for access to some open mailing lists.</li>';
      htmlBlob +=        '<li>Sign up for <a href="https://companyName.slack.com/">Slack</a> with your new companyName Google account.</li>';
      htmlBlob +=     '</ul>';
      htmlBlob +=   '</li>';
      htmlBlob +=   '<li>Please visit <a href="https://ITHelpURL">ITHelpURL</a> for IT resources and to ask technical questions.</li>';
      htmlBlob += '</ul>';
      htmlBlob += '<p>See you very soon!</p>';
      
      break;
      
    case "denyRequest":
      subject = "[Newbie App] Request for " + data.newbieRecord.companyNameEmail + " denied by " + recipients.from;
      
      htmlBlob =  '<p>Your request for a newbie account for ' + data.newbieRecord.companyNameEmail + ' has been denied.<br/>';
      htmlBlob += 'Here\'s the reason:</p>';
      htmlBlob += '<p>' + data.message + '</p>';
      htmlBlob += '<p>Please submit your request again or email ' + recipients.from + ' for more info.</p>';
      
      break;
      
    case "alertMailboxNotSetup":
      subject = "[Newbie App] Warning: newbie mailbox not setup for " + data.userEmail;
      
      htmlBlob =  '<p>The user ' + data.userEmail + ' has been created and added to groups correctly.</p>';
      htmlBlob += '<p>However, Google reported his mailbox as not setup. Please check if he\'s been assigned a G-Suite Business license on the Admin console.</p>';
      
      break;
      
    case "testing":
      subject = "[Newbie App] Test email from App Maker";
      
      var date = new Date();
      
      htmlBlob =  '<p>This email comes from App Maker</p>';
      htmlBlob += '<p>It was called by ' + data.origin + ' at ' + date + '</p>';
      
      break;
      
    case "error":
      subject = "[Newbie App] Error for " + data.userEmail;
      
      htmlBlob =  '<p>Groupies encountered an error. Here\'s the error log:</p>';
      htmlBlob += '<p>' + data.error + '</p>';
      
      break;
      
    case "newbieRequestEdited":
      subject = "[Newbie App] Newbie EDITED: " + data.newbieRecord.companyNameEmail + ' - starting on ' + data.newbieRecord.StartDate.toISOString().slice(0,10);
      
      htmlBlob = '<h1>Newbie request</h1>';
      htmlBlob += '<p><b>Name: </b>' + data.newbieRecord.FirstName + ' ' + data.newbieRecord.LastName + '<br/>';
      htmlBlob += '<b>Email address: </b>' + data.newbieRecord.companyNameEmail + '<br/>';
      htmlBlob += '<b>Start date: </b>' + data.newbieRecord.StartDate.toISOString().slice(0,10) + '</p>';
      htmlBlob += '<p><b>Requester: </b>' + recipients.from + '<br/>';
      htmlBlob += '<b>Comments: </b>' + data.newbieRecord.Notes + '<br/><br/></p>';
      
      if (data.newbieRecord.NeedsLaptop) {
        htmlBlob += '<p><b>This companyNamebie requires a laptop, please make sure a machine is ready for them on their first day</b><br/></p>';
      }
      
      htmlBlob += '<p><b>Make sure you check/increase the license count through <a href="https://licensePortalURL">License Portal</a> before approving this request</b></p>';
      htmlBlob += '<span><a href="https://script.google.com/a/macros/companyName.com/s/SomeRandomlyGeneratedString/exec#RequestsAdminV2" style="height:36px;border-radius:2px;color:white;background-color:rgb(33,150,243);border:none;padding:8px 16px;text-decoration:none;" onclick="">VIEW IN APP</a></span>';
      
      break;
      
    default:
      return;
  }
  
  var template = HtmlService.createTemplate(htmlBlob);
  var htmlBody = template.evaluate().getContent();
  
  sendEmail_(recipients.to, subject, htmlBody, recipients.bcc);
}