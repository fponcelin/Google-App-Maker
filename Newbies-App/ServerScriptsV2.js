function newbieRequestSubmitted(key) {
  var newbie = app.models.Newbies.getRecord(key);
  console.log('newbie request received: %s', newbie.companyNameEmail);
  generateEmail_({from: newbie.RequesterEmail, to: "help@companyName.it", bcc: ""}, "newbieRequestV2", {newbieRecord: newbie});
}

/**
* Inserts a user in a mailing group
*
* @param {string} groupEmail - Email address of the group
* @param {User Resource} userResource - AdminDirectory user resource JSON of the newbie to be inserted in the group
*/
function insertNewbieinGroup_(groupEmail, userResource) {
  var requestBody = {
    kind: "admin#directory#member",
    etag: userResource.etag,
    id: userResource.id,
    email: userResource.primaryEmail,
    role: 'MEMBER',
    type: 'USER'
  };
  
  try {
    var member = AdminDirectory.Members.insert(requestBody, groupEmail);
    console.log('New group member: %s in %s', member.email, groupEmail);
  } catch (e) {
    console.error(JSON.stringify(e));
    generateEmail_({from: '', to: 'it@companyName.com', bcc: ''}, "error", {userEmail: userResource.primaryEmail, error: JSON.stringify(e)});
  }
}

/**
* Gets the groups from the Newbie record and calls the insertNewbieInGroup function
*
* @param {User Resource} userResource - AdminDirectory user resource JSON of the newbie to be inserted in the group
* @param {Datasource Record} newbieRecord - Newbies datasource record
*/
function assignGroupsV2_(userResource, newbieRecord) {
  var groupsJSON = JSON.parse(newbieRecord.Groups);

  for (var i=0; i < groupsJSON.emails.length; i++) {
    insertNewbieinGroup_(groupsJSON.emails[i], userResource);
  }

  var user = AdminDirectory.Users.get(userResource.primaryEmail); //get the updated user after groups insertion

  if (!user.isMailboxSetup) {
    generateEmail_({from: "", to: "it@companyName.com", bcc: ""}, "alertMailboxNotSetup", {userEmail: userResource.primaryEmail});
  }
}

/**
* Creates the newbie's companyName Google account and assign to relevant groups
*
* @param {Datasource Record} newbieRecord - Newbies datasource record
*/
function createGoogleAccountV2_(newbieRecord){
  var orgUnitPath = '/' + newbieRecord.Org;
  
  if (newbieRecord.Contract.indexOf('Permanent') == -1)
    orgUnitPath += '/Non-perm';
  
  var password = Math.random().toString(36);
  var user = {
    primaryEmail: newbieRecord.companyNameEmail,
    name: {
      givenName: newbieRecord.FirstName,
      familyName: newbieRecord.LastName
    },
    password: password,
    orgUnitPath: orgUnitPath,
    changePasswordAtNextLogin: true
  };
  
  
  //*********** USER INSERTION IN G-SUITE ****************
  try {
    user = AdminDirectory.Users.insert(user); //This line creates the user in G-Suite Directory ########## comment out for testing ##########
  } catch(e) {
    throw new app.ManagedError('User insertion failed, please check licenses');
  }
  
  console.log(user);
  newbieRecord.IsAccountCreated = true;
  newbieRecord.AccountCreationDate = new Date();
  newbieRecord.ApprovalStatus = 'Approved';
  app.saveRecords([newbieRecord]);
  
  assignGroupsV2_(user, newbieRecord);
  
  //******************************************************
  
  generateEmail_({from: newbieRecord.RequesterEmail, to: newbieRecord.PersonalEmail, bcc: "it@companyName.com"}, "welcomeNewbie", {newbieRecord: newbieRecord, password: password});
  //############## Comment out for testing #############
}

/**
* Approves the newbie request, calls for the google account creation and the email notifications
*
* @param {int} newbieID - Newbie record primary key
*/
function approveNewbieRequestV2(newbieID, adminEmail){
  var newbieRecord = app.models.Newbies.getRecord(newbieID);
  
  newbieRecord.IsApproved = true;
  newbieRecord.ApprovalDate = new Date();
  newbieRecord.ApprovalStatus = "ApprovalInProgress";
  newbieRecord.ApproverEmail = adminEmail;
  app.saveRecords([newbieRecord]);
  
  createGoogleAccountV2_(newbieRecord);
}


function newbieRequestEdited(user, key) {
  var newbieRecord = app.models.Newbies.getRecord(key);
  generateEmail_({from: user, to: "help@companyName.it", bcc: ""}, "newbieRequestEdited", {newbieRecord: newbieRecord});
}

/**
* Denies the newbie request, calls for the email notification
*
* @param {int} newbieID - Newbie record primary key
* @param {string} message - message to be sent by email to the requester
* @param {strin} adminEmail - email address of the admin denying the request
*/
function denyNewbieRequest(newbieID, message, adminEmail){
  var newbieRecord = app.models.Newbies.getRecord(newbieID);
  
  newbieRecord.ApprovalDate = new Date();
  newbieRecord.ApprovalStatus = "Denied";
  app.saveRecords([newbieRecord]);
  
  generateEmail_({from: adminEmail, to: newbieRecord.RequesterEmail, bcc: "it@companyName.com"}, "denyRequest", {newbieRecord: newbieRecord, message: message});
}


/**
* Gets the counts of total and used GSuite licenses and calculates the difference
* Code copied and adapted from https://stackoverflow.com/a/47372561
*
* @return {string} - Calculated amount of licenses with date of data point
*/
function getLicenseCount() {
  var tryDate = new Date();
  var dateString = tryDate.getFullYear().toString() + "-" + (tryDate.getMonth() + 1).toString() + "-" + tryDate.getDate().toString();
  var response;
  while (true) {
    try {
      response = AdminReports.CustomerUsageReports.get(dateString,{parameters : "accounts:gsuite_enterprise_total_licenses,accounts:gsuite_enterprise_used_licenses"});
      break;
    } catch(e) {
      //Logger.log(e.warnings.toString());
      tryDate.setDate(tryDate.getDate()-1);
      dateString = tryDate.getFullYear().toString() + "-" + (tryDate.getMonth() + 1).toString() + "-" + tryDate.getDate().toString();
      continue;
    }
  }
  
  var totalLicenseCount = response.usageReports[0].parameters[0].intValue;
  var usedLicenseCount = response.usageReports[0].parameters[1].intValue;
  Logger.log("Total licenses:" + totalLicenseCount.toString());
  Logger.log("Used licenses:" + usedLicenseCount.toString());
  
  var availableLicenseCount = totalLicenseCount - usedLicenseCount;
  
  return availableLicenseCount.toString() + ' as of ' + dateString;
}