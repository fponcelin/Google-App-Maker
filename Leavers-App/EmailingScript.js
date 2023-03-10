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


function generateEmail_(event, leaver, eventTriggeredBy) {
  var subject, htmlBlob, to, bcc;
  
  switch (event) {
    case 'newLeaver':
      subject = '[Leavers App] Leaver request: ' + leaver.Email + ' - leaving on: ' + leaver.LeaveDate.toISOString().slice(0,10);
      
      htmlBlob =  '<h1>New upcoming Leaver</h1>';
      htmlBlob += '<p><b>Name: </b>' + leaver.FullName + '<br/>';
      htmlBlob += '<b>Email address: </b>' + leaver.Email + '<br/>';
      htmlBlob += '<b>Leave Date: </b>' + leaver.LeaveDate.toISOString().slice(0,10) + '<br/>';
      htmlBlob += '<b>Retention Policy: </b>' + leaver.RetentionPolicy + ' days<br/>';
      htmlBlob += '<b>Comments: </b>' + leaver.Comments + '</p>';
      htmlBlob += '<p><b>Requested by: </b>' + leaver.RequesterEmail + '</p>';
      
      to = 'help@companyName.it';
      bcc = '';
      break;
      
    case 'leaverEdited':
      subject = '[Leavers App] Leaver edited: ' + leaver.Email + ' - leaving on: ' + leaver.LeaveDate.toISOString().slice(0,10);
      
      htmlBlob =  '<h1>Leaver edited with the following information</h1>';
      htmlBlob += '<p><b>Name: </b>' + leaver.FullName + '<br/>';
      htmlBlob += '<b>Email address: </b>' + leaver.Email + '<br/>';
      htmlBlob += '<b>Leave Date: </b>' + leaver.LeaveDate.toISOString().slice(0,10) + '<br/>';
      htmlBlob += '<b>Retention Policy: </b>' + leaver.RetentionPolicy + ' days<br/>';
      htmlBlob += '<b>Comments: </b>' + leaver.Comments + '</p>';
      htmlBlob += '<p><b>Information edited by: </b>' + eventTriggeredBy + '</p>';
      
      to = 'it@companyName.com';
      bcc = '';
      break;
      
    case 'pendingLeaverCancelled':
      subject = '[Leavers App] Leaver cancelled: ' + leaver.Email;
      
      htmlBlob =  '<h1>Upcoming Leaver cancelled</h1>';
      htmlBlob += '<p><b>Name: </b>' + leaver.FullName + '<br/>';
      htmlBlob += '<b>Email address: </b>' + leaver.Email + '</p>';
      htmlBlob += '<p><b>Cancelled by: </b>' + eventTriggeredBy + '</p>';
      
      to = 'it@companyName.com';
      bcc = '';
      break;
      
    case 'restoreRequestSubmitted':
      subject = '[Leavers App] Restore request: ' + leaver.Email;
      
      htmlBlob =  '<h1>Request to restore suspended account</h1>';
      htmlBlob += '<p><b>Name: </b>' + leaver.FullName + '<br/>';
      htmlBlob += '<b>Email address: </b>' + leaver.Email + '<br/>';
      htmlBlob += '<b>Returning date: </b>' + leaver.RestoreRequestDate.toISOString().slice(0,10) + '</p>';
      htmlBlob += '<p><b>Request by: </b>' + eventTriggeredBy + '</p>';
      htmlBlob += '<p><b>Comments: </b>' + leaver.RestoreRequestComments + '</p>';
      
      to = 'help@companyName.it';
      bcc = '';
      break;
      
    case 'restoreRequestCancelled':
      subject = '[Leavers App] Cancelled restore request: ' + leaver.Email;
      
      htmlBlob =  '<h1>Cancelled request to restore suspended account</h1>';
      htmlBlob += '<p><b>Name: </b>' + leaver.FullName + '<br/>';
      htmlBlob += '<b>Email address: </b>' + leaver.Email + '</p>';
      htmlBlob += '<p><b>Request by: </b>' + eventTriggeredBy + '</p>';
      
      to = 'it@companyName.com';
      bcc = '';
      break;
      
    case 'restoreRequestDenied':
      subject = '[Leavers App] Denied request to restore suspended account: ' + leaver.Email;
      
      htmlBlob =  '<h1>' + eventTriggeredBy + ' denied your request to restore suspended account</h1>';
      htmlBlob += '<p><b>Name: </b>' + leaver.FullName + '<br/>';
      htmlBlob += '<b>Email address: </b>' + leaver.Email + '</p>';
      htmlBlob += '<p><b>Comments from IT: </b>' + leaver.RestoreApproverComments + '</p>';
      
      to = leaver.RestoreRequesterEmail;
      bcc = 'it@companyName.com';
      break;
      
    case 'leaverSuspended':
      subject = '[Leavers App] Account suspended: ' + leaver.Email;
      
      htmlBlob =  '<h1>' + eventTriggeredBy + ' triggered account suspension</h1>';
      htmlBlob += '<p><b>Name: </b>' + leaver.FullName + '<br/>';
      htmlBlob += '<b>Email address: </b>' + leaver.Email + '<br/>';
      htmlBlob += '<b>Renamed email: </b>' + leaver.RenamedEmail + '<br/>';
      
      if (leaver.DeparturesGroupEmail !== null)
        htmlBlob += '<p><b>Email alias added to group: </b>' + leaver.DeparturesGroupEmail + '</p>';
      
      htmlBlob += '<p>Sadly I can\'t transfer Drive data at this stage, <b>please run the following command in GAM:</b><br/>';
      htmlBlob += '<code>gam create datatransfer ' + leaver.RenamedEmail + ' drive ' + leaver.DriveTransferEmail + ' all</code></p>';
      htmlBlob += '<p>Don\'t forget to do the usual clean up as per the <a href="https://help.companyName.it/support/solutions/articles/[articleID]" target="_blank">off-boarding list</a></p>';
      
      to = 'it@companyName.com';
      bcc = '';
      break;
      
    case 'leaverRestored':
      subject = '[Leavers App] Suspended account has been restored: ' + leaver.Email;
      
      htmlBlob =  '<h1>' + eventTriggeredBy + ' approved your request to restore suspended account</h1>';
      htmlBlob += '<p><b>Name: </b>' + leaver.FullName + '<br/>';
      htmlBlob += '<b>Email address: </b>' + leaver.Email + '<br/>';
      htmlBlob += '<p><b>Comments from IT: </b>' + leaver.RestoreApproverComments + '</p>';
      
      to = leaver.RestoreRequesterEmail;
      bcc = 'it@companyName.com';
      break;
      
    case 'restoreInstructions':
      subject = 'Your companyName account has been restored: ' + leaver.Email;
      
      htmlBlob =  '<h1>Welcome back to companyName!</h1>';
      htmlBlob += '<p><b>Your email address: </b>' + leaver.Email + '<br/>';
      htmlBlob += '<b>Your temporary password: </b>' + leaver.RestoreTempPassword + '</p>';
      htmlBlob += '<p><b><a href="https://www.google.com/accounts/AccountChooser?Email=' + leaver.Email + '&continue=https://apps.google.com/user/hub">Please click here</a></b> to sign in.</p>';
      
      to = leaver.PersonalEmail;
      bcc = 'it@companyName.com';
      
      break;
      
    case 'leaverDeleted':
      subject = '[Leavers App] Account deleted: ' + leaver.RenamedEmail;
      
      htmlBlob =  '<h1>' + eventTriggeredBy + ' triggered account deletion</h1>';
      htmlBlob += '<p><b>Name: </b>' + leaver.FullName + '<br/>';
      htmlBlob += '<b>Email address: </b>' + leaver.Email + '<br/>';
      htmlBlob += '<b>Renamed email: </b>' + leaver.RenamedEmail + '</p>';
      
      if (leaver.DeparturesGroupEmail) {
        htmlBlob += '<p><b>Note: </b>The original email address (' + leaver.Email + ') was added as an alias to the group ' + leaver.DeparturesGroupEmail + '.<br/>';
        htmlBlob += 'It will remain there unless manually deleted.</p>';
      }
      
      to = 'it@companyName.com';
      bcc = '';
      
      break;
      
    case 'error':
      subject = '[Leavers App] An error occurred on the Leavers app while processing ' + leaver.Email;
      
      htmlBlob =  '<p>Leavers encountered an error while processing ' + leaver.Email + '<br/>';
      htmlBlob += 'Current status: ' + leaver.Status + '</p>';
      htmlBlob += '<p>Here\'s the error log:<br/>';
      htmlBlob += eventTriggeredBy + '</p>';
      htmlBlob += '<p>You can view detailed logs <a href="https://console.cloud.google.com/logs/viewer?project=project-id-[ProjectIDNumber]">here</a></p>';
      
      to = 'it@companyName.com';
      bcc = '';
      
      break;
      
    default:
      return;
  }
  
  var template = HtmlService.createTemplate(htmlBlob);
  var htmlBody = template.evaluate().getContent();
  
  sendEmail_(to, subject, htmlBody, bcc);
}


function emailFromClientScript(event, leaverID, eventTriggeredBy) {
  var leaver = app.models.Leavers.getRecord(leaverID);
  
  console.log('event: ' + event);
  console.log('leaverEmail: ' + leaver.Email);
  
  generateEmail_(event, leaver, eventTriggeredBy);
}

