
function addGoogleDivTemplate(email) {
  var htmlBlob = '<div>';
  htmlBlob += '<div dir="ltr" style="border:1px solid #f0f0f0;max-width:650px;font-family:Arial,sans-serif;color:#000000">';
  htmlBlob += '<div style="background-color:#f5f5f5;padding:10px 12px">';
  htmlBlob += '<table cellpadding="0" cellspacing="0" style="width:100%"><tbody><tr><td style="width:50%">';
  htmlBlob += '<span style="font:20px/24px arial;color:#333333"><a href="https://groups.google.com/a/companyName.com/d/forum/';
  htmlBlob += email.replace('@companyName.com', '');
  htmlBlob += '" style="text-decoration:none;color:#000000" target="_blank">';
  htmlBlob += email.replace('@companyName.com', '');
  htmlBlob += '</a></span>';
  htmlBlob += '</td><td style="text-align:right;width:50%">';
  htmlBlob += '<span style="font:20px/24px arial"><a style="color:#dd4b39;text-decoration:none" href="https://groups.google.com/a/companyName.com/d/overview" target="_blank">Google Groups</a></span>';
  htmlBlob += '</td><td style="text-align:right">';
  htmlBlob += '<a href="https://groups.google.com/a/companyName.com/d/overview" target="_blank"><img style="border:0;vertical-align:middle;padding-left:10px" src="http://www.google.com/images/icons/product/groups-32.png" alt="Logo for Google Groups"></a>';
  htmlBlob += '</td></tr></tbody></table></div>';
  htmlBlob += '<div style="margin:30px 30px 30px 30px;line-height:21px">';
  htmlBlob += '<span style="font-size:20px;color:#000000;line-height:36px">Congratulations, you&#39;ve created <a href="https://groups.google.com/a/companyName.com/d/forum/';
  htmlBlob += email.replace('@companyName.com', '');
  htmlBlob += '" style="color:#1155cc;text-decoration:none" target="_blank"><b>';
  htmlBlob += email.replace('@companyName.com', '');
  htmlBlob += '</b></a></span><br></div>';
  htmlBlob += '<div style="margin:30px 30px 30px 30px;line-height:21px">';
  htmlBlob += '<span style="font-size:20px;color:#000000;line-height:36px">Start inviting people to your group</span><br>';
  htmlBlob += '<span style="font-size:13px;color:#333333">With Google Groups you can invite anyone to join your group using their email address. Know who you want in your group already? You can skip the invite and directly add up to 10 people at once.<br>';
  htmlBlob += '<span><a href="https://support.google.com/groups/answer/2465464" style="color:#1155cc;text-decoration:none" target="_blank">Learn about adding members</a></span></span></div>';
  htmlBlob += '<div style="margin:30px 30px 30px 30px;line-height:21px">';
  htmlBlob += '<span style="font-size:20px;color:#000000;line-height:36px">Share Google Docs and Calendars directly with your group</span><br>';
  htmlBlob += '<span style="font-size:13px;color:#333333">Working with shared documents and calendars has never been easier than with Google Groups. Share your files with the group and it will automatically be updated whenever someone is added to or removed from the group.<br>';
  htmlBlob += '<span><a href="https://support.google.com/a/answer/167101" style="color:#1155cc;text-decoration:none" target="_blank">Learn about sharing with groups</a></span></span></div>';
  htmlBlob += '<div style="margin:30px 30px 30px 30px;line-height:21px">';
  htmlBlob += '<span style="font-size:20px;color:#000000;line-height:36px">Become organised</span><br>';
  htmlBlob += '<span style="font-size:13px;color:#333333">Google Groups has several options for organising content in your discussion group, including: categories, tags and tracking topic resolutions.<br>';
  htmlBlob += '<span><a href="https://support.google.com/a/answer/126169" style="color:#1155cc;text-decoration:none" target="_blank">Learn about advanced Groups features</a></span></span></div>';
  htmlBlob += '<div style="margin:30px 30px 30px 30px;line-height:21px">';
  htmlBlob += '<a style="border-radius:2px;display:inline-block;padding:0px 8px;background-color:#498af2;color:#ffffff;font-size:11px;border:solid 1px #3079ed;font-weight:bold;text-decoration:none;min-width:54px;text-align:center;line-height:27px" href="https://groups.google.com/a/companyName.com/d/forum/';
  htmlBlob += email.replace('@companyName.com', '');
  htmlBlob += '" target="_blank">Visit your group</a></div>';
  htmlBlob += '<div style="background-color:#f5f5f5;padding:5px 12px"><table cellpadding="0" cellspacing="0" style="width:100%"><tbody><tr><td style="padding-top:4px;font-family:arial,sans-serif;color:#636363;font-size:11px">';
  htmlBlob += '<a href="http://support.google.com/groups/bin/answer.py?answer=3D46601" style="color:#1155cc;text-decoration:none" target="_blank">Visit</a> the help centre.';
  htmlBlob += '</td></tr></tbody></table></div></div></div>';

  return htmlBlob;
}


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
* @param {recipients} recipients - object with 3 properties: from (requester email string), to (destination emails string), bcc
* @param {string} eventType - type of event triggering the notification
* @param {data} data - data to be emailed. This object can have the following properties: {requestRecord (AppMaker Datasource Record)}
*/
function generateEmail_(recipients, eventType, data){
  var subject, htmlBlob, ownersMembersJSON;
  
  switch(eventType) {
    case "groupRequest":
      ownersMembersJSON = JSON.parse(data.requestRecord.OwnersMembersJSON);
      subject = "[Groupies App] Group request from " + recipients.from;
      
      htmlBlob =  '<h1>Group request</h1>';
      htmlBlob += '<p><b>Name: </b>' + data.requestRecord.GroupName + '<br/>';
      htmlBlob += '<b>Email address: </b>' + data.requestRecord.GroupEmail + '<br/>';
      htmlBlob += '<b>Requester: </b>' + data.requestRecord.RequesterEmail + '<br/>';
      htmlBlob += '<b>Description: </b>' + data.requestRecord.Description + '</p>';
      htmlBlob += '<p><b>Team Drive: </b>' + data.requestRecord.TeamDrive + '<br/>';
      htmlBlob += '<b>Group Calendar: </b>' + data.requestRecord.Calendar + '</p>';
      htmlBlob += '<b>Owners:</b>';
      htmlBlob += '<ul>';
      
      for (var i=0; i < ownersMembersJSON.owners.length; i++) {
        htmlBlob += '<li>' + ownersMembersJSON.owners[i].email + '</li>';
      }
      
      htmlBlob += '</ul>';
      htmlBlob += '<b>Members:</b>';
      htmlBlob += '<ul>';
      
      for (i=0; i < ownersMembersJSON.members.length; i++) {
        htmlBlob += '<li>' + ownersMembersJSON.members[i].email + '</li>';
      }
      
      htmlBlob += '</ul>';
      htmlBlob += '<p><a href="https://script.google.com/a/macros/companyName.com/s/SomeLongRandomString/exec#Admin" style="height:36px;border-radius:2px;color:white;background-color:rgb(33,150,243);border:none;padding:8px 16px;text-decoration:none;" onclick="">VIEW IN APP</a></p>';
      
      break;
      
    case "requestConfirmed":
      ownersMembersJSON = JSON.parse(data.requestRecord.OwnersMembersJSON);
      subject = "Your group request for " + data.requestRecord.GroupEmail + " has been received";
      
      htmlBlob =  '<h1>Group request</h1>';
      htmlBlob += '<p>Thank you for your request. Our IT team is reviewing your group and will process it shortly. Here\'s a summary of your request:</p>';
      htmlBlob += '<p><b>Name: </b>' + data.requestRecord.GroupName + '<br/>';
      htmlBlob += '<b>Email address: </b>' + data.requestRecord.GroupEmail + '<br/>';
      htmlBlob += '<b>Requester: </b>' + data.requestRecord.RequesterEmail + '<br/>';
      htmlBlob += '<b>Description: </b>' + data.requestRecord.Description + '</p>';
      htmlBlob += '<p><b>Team Drive: </b>' + data.requestRecord.TeamDrive + '<br/>';
      htmlBlob += '<b>Group Calendar: </b>' + data.requestRecord.Calendar + '</p>';
      htmlBlob += '<b>Owners:</b>';
      htmlBlob += '<ul>';
      
      for (var j=0; j < ownersMembersJSON.owners.length; j++) {
        htmlBlob += '<li>' + ownersMembersJSON.owners[j].email + '</li>';
      }
      
      htmlBlob += '</ul>';
      htmlBlob += '<b>Members:</b>';
      htmlBlob += '<ul>';
      
      for (j=0; j < ownersMembersJSON.members.length; j++) {
        htmlBlob += '<li>' + ownersMembersJSON.members[j].email + '</li>';
      }
      
      htmlBlob += '</ul>';
      break;
      
    case "groupCreated":
      subject = "Welcome to the " + data.requestRecord.GroupName + " mailing group";
      
      htmlBlob =  '<p>Hello friend!</p>';
      htmlBlob += '<p>Welcome to the mailing group for ' + data.requestRecord.GroupName + '. This group was requested by ' + data.requestRecord.RequesterEmail + '<br/>';
      htmlBlob += addGoogleDivTemplate(data.requestRecord.GroupEmail);
      
      if (data.requestRecord.TeamDrive)
        htmlBlob += '<p>A Team Drive with the same name as the group as also been created.</p>';
      
      htmlBlob += '<p>Enjoy!</p>';
      
      break;
      
    case "denyRequest":
      subject = "Request for " + data.requestRecord.GroupEmail + " denied by " + recipients.from;
      
      htmlBlob =  '<p>Your request for a mailing group for ' + data.requestRecord.GroupEmail + ' has been denied.<br/>';
      htmlBlob += 'Here\'s the reason:</p>';
      htmlBlob += '<p>' + data.message + '</p>';
      htmlBlob += '<p>Please submit your request again or email ' + recipients.from + ' for more info.</p>';
      
      break;
      
    case "testing":
      subject = "[Groupies App] Test email from App Maker";
      
      var date = new Date();
      
      htmlBlob =  '<p>This email comes from App Maker</p>';
      htmlBlob += '<p>It was called by ' + data.origin + ' at ' + date + '</p>';
      
      break;
      
    case "error":
      subject = "[Groupies App] Error for " + data.requestRecord.GroupEmail;
      
      htmlBlob =  '<p>Groupies encountered an error. Here\'s the error log:</p>';
      htmlBlob += '<p>' + data.error + '</p>';
      
      break;
      
    default:
      return;
  }
  
  var template = HtmlService.createTemplate(htmlBlob);
  var htmlBody = template.evaluate().getContent();
  
  sendEmail_(recipients.to, subject, htmlBody, recipients.bcc);
}