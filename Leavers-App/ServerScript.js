/**
 * Gets leaver's group membership
 * @param {string} email - leaver's email;
 * @return {JSON} array of email addresses in JSON format.
 */
function getLeaverGroups_(email) {
  var pageToken, page, groups;
  var emails = [];
  do {
    page = AdminDirectory.Groups.list({
      userKey: email,
      maxResults: 100,
      pageToken: pageToken
    });
    groups = page.groups;
    if (groups) {
      for (var i=0; i < groups.length; i++) {
        emails.push(groups[i].email);
      }
    }
    pageToken = page.nextPageToken;
  } while (pageToken);
  
  if (emails.length > 0)
    return { 'emails': emails };
  else
    return false;
}

/**
 * Generates the leave date in YYYYMMDD format.
 * @param {date} date - leaver's leave date;
 * @return {string} formatted leave date.
 */
function yyyymmdd_(date) {
    var y = date.getFullYear();
    var m = date.getMonth() + 1;
    var d = date.getDate();
    var mm = m < 10 ? '0' + m : m;
    var dd = d < 10 ? '0' + d : d;
    return '' + y + mm + dd;
}

/**
 * Updates the leaver's SQL record with group membership, Org Unit path and renamed email address prior to account suspension.
 * @param {SQL Record} leaver - leaver record;
 * @param {AdminDirectory User} gsuiteUser - matching AdminDirectory GSuite user;
 * @param {string} adminEmail - email address of the IT admin approving the account suspension.
 */
function saveLeaverMetadata_(leaver, gsuiteUser) {
  console.log('Saving leaver metadata...');
  
  var leaverGroups = getLeaverGroups_(leaver.Email);
  if (leaverGroups) {
    leaver.Groups = JSON.stringify(leaverGroups);
    console.log('Leaver groups: %s', leaver.Groups);
  } else {
    leaver.Groups = null;
  }
  
  leaver.OrgUnitPath = gsuiteUser.orgUnitPath;
  console.log('Leaver OU: %s', leaver.OrgUnitPath);
  
  leaver.RenamedEmail = leaver.Email.replace('@companyName.com', '') + '_' + yyyymmdd_(leaver.LeaveDate) + '@companyName.com';
  console.log('Leaver RenamedEmail: %s', leaver.RenamedEmail);
  
  app.saveRecords([leaver]);
}


/**
 * Remove user from Google groups.
 * @param {SQL Record} leaver - leaver record;
 */
function deleteGroups_(leaver) {
  if (leaver.Groups !== null) {
    var groups = JSON.parse(leaver.Groups);
    console.log('Removing Leaver from groups...');

    try {
      for (var i=0; i < groups.emails.length; i++) {
        console.log('Removing user ' + leaver.Email + ' from group: ' + groups.emails[i]);
        AdminDirectory.Members.remove(groups.emails[i], leaver.Email); //################## Comment out for testing ##############
      }
    } catch (e) {
      console.error(JSON.stringify(e));
      generateEmail_('error', leaver, 'deleteGroups function error: ' + JSON.stringify(e));
      throw new app.ManagedError('Error while removing the user from mailing groups (step 2). Change the status of this leaver back to "pending" and try again.');
    }
  } else {
    console.log('No groups to be removed from');
  }
}

/**
 * Wipes the user's calendar.
 * @param {SQL Record} leaver - leaver record;
 */
function wipeCalendar_(leaver) {
  console.log('Wiping Leaver calendar...');
  var errors = [];
  
  console.log('Getting Calendar events...');
  var events = Calendar.Events.list(leaver.Email);

  console.log('Deleting %s events...', events.items.length);
  for (var i=0; i<events.items.length; i++) {
    console.log('Deleting event %s/%s - event ID: %s', i+1, events.items.length, events.items[i].id);
    console.log('Event status: %s', events.items[i].status);
    if (events.items[i].status != 'cancelled') {
      try {
        Calendar.Events.remove(leaver.Email, events.items[i].id);
      } catch (e) {
        console.error(JSON.stringify(e));
        errors.push('<br/>eventID: ' + events.items[i].id + '- error: ' + JSON.stringify(e));
      }
    }
  }

  if (errors.length > 0)
    generateEmail_('error', leaver, errors.length + ' wipeCalendar function errors: ' + errors);
  console.log('Calendar wiped');
}

/**
 * Deprovisions the user: delete Application Specific Passwords (if any), invalidates backup Verification Codes (if any) and deletes 3rd party Tokens (if any)
 * @param {SQL Record} leaver - leaver record;
 */
function deprovision_(leaver) {
  console.log('Deprovisioning Leaver...');
  
  try {
    //Delete Application Specific Passwords
    console.log('Deleting ASPs...');
    var asps = AdminDirectory.Asps.list(leaver.Email);
    if (asps.items) {
      for (var i=0; i<asps.items.length; i++) {
        console.log('Deleting ASP: ' + asps.items[i]);
        AdminDirectory.Asps.remove(asps.items[i].userKey, asps.items[i].codeId); //################## Comment out for testing ##############
      }
      console.log('ASPs deleted');
    } else {
      console.log('No ASPs to delete');
    }

    //Invalidate backup Verification Codes
    console.log('Invalidating backup Verification Codes...');
    AdminDirectory.VerificationCodes.invalidate(leaver.Email); //################## Comment out for testing ##############

    //Delete 3rd party Tokens
    console.log('Deleting 3rd party Tokens...');
    var tokens = AdminDirectory.Tokens.list(leaver.Email);
    if (tokens.items) {
      for (var j=0; j<tokens.items.length; j++) {
        console.log('Deleting Token: ' + tokens.items[j]);
        AdminDirectory.Tokens.remove(tokens.items[j].userKey, tokens.items[j].clientId); //################## Comment out for testing ##############
      }
      console.log('Tokens deleted');
    } else {
      console.log('No Tokens to delete');
    }
  } catch (e) {
    console.error(JSON.stringify(e));
    generateEmail_('error', leaver, 'deprovision function error: ' + JSON.stringify(e));
    throw new app.ManagedError('Error while deprovisioning user (step 3).  Change the status of this leaver back to "pending" and try again.');
  }
}

/**
 * Suspends the user's GSuite account, renames their primary email, moves it to the Leavers sub-OU, hides it from the Directory and randomises the password.
 * @param {SQL Record} leaver - leaver record;
 * @return {AdminDirectory User} gsuiteUser - updated User resource.
 */
function suspendAccount_(leaver) {
  var requestBody = {
    includeInGlobalAddressList: false,
    primaryEmail: leaver.RenamedEmail,
    orgUnitPath: leaver.OrgUnitPath.slice(0, 4) + '/Leavers',
    password: Math.random().toString(36),
    suspended: true
  };
  
  var gsuiteUser;
  
  try {
    console.log('Suspending GSuite user...');
    gsuiteUser = AdminDirectory.Users.update(requestBody, leaver.Email); //################## Comment out for testing ##############
    console.log('User suspended');
  } catch (e) {
    console.error(JSON.stringify(e));
    generateEmail_('error', leaver, 'suspendAccount function error: ' + JSON.stringify(e));
    throw new app.ManagedError('Error while suspending user (step 4). Please process user manually on GSuite admin panel, then amend user\'s metadata in this app accordingly');
  }
  
  return gsuiteUser;
}

/**
 * ############## Apps Script Transfer API doesn't seem to exist at this time, function unused ###############
 * Generates a Drive data transfer request
 * @param {SQL Record} leaver - leaver record;
 */
function transferDriveData_(leaver) {
  var requestBody = {
    kind: "admin#datatransfer#DataTransfer",
    oldOwnerUserId: leaver.RenamedEmail,
    newOwnerUserId: leaver.DriveTransferEmail,
    applicationDataTransfers: [{
      applicationTransferParams: [{
        key: "PRIVACY_LEVEL",
        value: [
          "SHARED"
        ]
      }]
    }]
  };
  
}

/**
 * Deletes the user's original email from his aliases and adds it to the selected Departures group.
 * @param {SQL Record} leaver - leaver record;
 */
function moveAliasToDeparturesGroup_(leaver) {
  console.log('Moving email alias...');
  try {
    //Delete alias from user account
    console.log('Deleting alias %s from User %s...', leaver.Email, leaver.RenamedEmail);
    AdminDirectory.Users.Aliases.remove(leaver.RenamedEmail, leaver.Email); //################## Comment out for testing ##############
    console.log('Alias deleted from User');

    //Add alias to relevant group
    console.log('Adding alias %s to Group %s...', leaver.Email, leaver.DeparturesGroupEmail);
    AdminDirectory.Groups.Aliases.insert({alias: leaver.Email}, leaver.DeparturesGroupEmail); //################## Comment out for testing ##############
    console.log('Alias added to Group');
  } catch (e) {
    console.error(JSON.stringify(e));
    generateEmail_('error', leaver, 'moveAliasToDeparturesGroup function error: ' + JSON.stringify(e));
    throw new app.ManagedError('Error while transferring alias to Departure group (step 5). Please use GAM or GSuite admin panel to process alias then amend user\'s metadata on this app accordingly.');
  }
}


/**
 * Main function for triggering account suspension process.
 * @param {string} leaverID - leaver's SQL record primary key;
 * @param {string} adminEmail - email address of the IT admin approving the account suspension.
 */
function approveLeaverSuspension(leaverID, adminEmail) {
  var leaver = app.models.Leavers.getRecord(leaverID);
  var gsuiteUser;
  try {
    gsuiteUser = AdminDirectory.Users.get(leaver.Email);
  } catch(e) {
    console.error(JSON.stringify(e));
    generateEmail_('error', leaver, 'approveLeaverSuspension function error: ' + JSON.stringify(e));
    throw new app.ManagedError('Error while getting user (step 0). Can\'t find ' + leaver.Email + ' on GSuite. Please make necessary changes (on GSuite or here on the user\'s record), set status back to "pending" and try again');
  }
  
  if (gsuiteUser.primaryEmail == leaver.Email && !gsuiteUser.suspended) {
    saveLeaverMetadata_(leaver, gsuiteUser);

    deleteGroups_(leaver);
    //wipeCalendar_(leaver); //*********** function unused, calendar.wipe function doesn't work in JS and individual events deletion is unreliable ***********
    deprovision_(leaver);
    gsuiteUser = suspendAccount_(leaver);
    //transferDriveData_(leaver); *********** function unused, no API for it :-( ***********
    
    if (leaver.DeparturesGroupEmail !== null) {
      moveAliasToDeparturesGroup_(leaver);
    }

    if (gsuiteUser.suspended) { //################## Comment out for testing ##############
      leaver.Status = 'suspended';
      leaver.SuspensionDate = new Date();
      app.saveRecords([leaver]);

    //*********** Notify IT ************
    generateEmail_('leaverSuspended', leaver, adminEmail);
    }
  }
}


/**
 * Unsuspends the user's GSuite account, renames their primary email, moves it to their original OU, unhides it from the Directory and sets the temp password.
 * @param {SQL Record} leaver - leaver record;
 * @return {AdminDirectory User} gsuiteUser - updated User resource.
 */
function unsuspendAccount_(leaver, tmpPwd) {
  var requestBody = {
    includeInGlobalAddressList: true,
    primaryEmail: leaver.Email,
    orgUnitPath: leaver.OrgUnitPath,
    password: tmpPwd,
    changePasswordAtNextLogin: true,
    suspended: false
  };
  var gsuiteUser;
  
  try {
    console.log('Unsuspending User %s...', leaver.Email);
    gsuiteUser = AdminDirectory.Users.update(requestBody, leaver.RenamedEmail); //################## Comment out for testing ##############
    console.log('User unsuspended');
  } catch (e) {
    console.error(JSON.stringify(e));
    generateEmail_('error', leaver, 'unsuspendAccount function error: ' + JSON.stringify(e));
    throw new app.ManagedError('Error while unsuspending user (step 2). Please check user on GSuite admin panel for any discrepancies with the user\'s record in this app, amend as needed and set status back to "restoreRequested". Then approve restore again');
  }
    
  return gsuiteUser;
}


/**
 * Adds the user as member of the groups they were in before suspension.
 * @param {SQL Record} leaver - leaver record;
 */
function restoreGroups_(leaver) {
  if (leaver.Groups !== null) {
    var groups = JSON.parse(leaver.Groups);
    var requestBody = {
      kind: "admin#directory#member",
      email: leaver.Email,
      role: 'MEMBER'
    };

    try {
      console.log('Restoring Groups...');
      for (var i=0; i<groups.emails.length; i++) {
        console.log('Inserting User %s in Group %s', leaver.Email, groups.emails[i]);
        var member = AdminDirectory.Members.insert(requestBody, groups.emails[i]); //################## Comment out for testing ##############
      }
      console.log('Groups restored');
    } catch (e) {
      console.error(JSON.stringify(e));
      generateEmail_('error', leaver, 'restoreGroups function error: ' + JSON.stringify(e));
      throw new app.ManagedError('Error while adding user back to mailing groups (step 4). Add groups manually with GAM or GSuite admin panel. Then set this record\'s status as "restored".');
    }
  } else {
    console.log('No Groups to restore');
  }
}


/**
 * Main function for triggering account restore process.
 * @param {string} leaverID - leaver's SQL record primary key;
 * @param {string} adminEmail - email address of the IT admin approving the account suspension.
 */
function approveLeaverRestore(leaverID, adminEmail) {
  var leaver = app.models.Leavers.getRecord(leaverID);
  var gsuiteUser;
  try {
    gsuiteUser = AdminDirectory.Users.get(leaver.RenamedEmail);
  } catch(e) {
    console.error(JSON.stringify(e));
    generateEmail_('error', leaver, 'approveLeaverRestore function error: ' + JSON.stringify(e));
    throw new app.ManagedError('Error while getting user (step 0). Can\'t find ' + leaver.RenamedEmail + ' on GSuite. Please make necessary changes (on GSuite or here on the user\'s record), set status back to "restoreRequested" and try again');
  }
  
  console.log('Checking if Alias was applied to Departure Group...');
  if (leaver.DeparturesGroupEmail !== null) {
    try {
      console.log('Deleting alias %s from Group %s...', leaver.Email, leaver.DeparturesGroupEmail);
      AdminDirectory.Groups.Aliases.remove(leaver.DeparturesGroupEmail, leaver.Email); //################## Comment out for testing ##############
      console.log('Alias deleted from Group');
    } catch (e) {
      console.error(JSON.stringify(e));
      generateEmail_('error', leaver, 'approveLeaverRestore function error while deleting alias from departures group: ' + JSON.stringify(e));
      throw new app.ManagedError('Error while deleting alias ' + leaver.Email + ' from ' + leaver.DeparturesGroupEmail + ' (step 1). Please delete alias manually with GAM or GSuite admin panel, then remove the Departure group from the user\'s record in this app and set status back to "restoreRequested". Then approve restore again');
    }
  }
  
  var tmpPwd = Math.random().toString(36);
  gsuiteUser = unsuspendAccount_(leaver, tmpPwd);
  
  try {
    console.log('Deleting alias %s from User %s...', leaver.RenamedEmail, leaver.Email);
    AdminDirectory.Users.Aliases.remove(leaver.Email, leaver.RenamedEmail); //################## Comment out for testing ##############
    console.log('Alias deleted');
  } catch (e) {
    console.error(JSON.stringify(e));
    generateEmail_('error', leaver, 'approveLeaverRestore function error while deleting renamed alias from user: ' + JSON.stringify(e));
    throw new app.ManagedError('Error while deleting renamed user alias (step 3). Delete alias manually with GAM or GSuite admin panel and restore mailing groups manually. Then set this record\'s status as "restored".');
  }
  
  restoreGroups_(leaver);
  
  if (!gsuiteUser.suspended) { //################## Comment out for testing ##############
    leaver.Status = 'restored';
    leaver.RestoreDate = new Date();
    leaver.RestoreTempPassword = tmpPwd;
    app.saveRecords([leaver]);
    
    //*********** Notify IT, requester and user ************
    generateEmail_('leaverRestored', leaver, adminEmail);
    generateEmail_('restoreInstructions', leaver, adminEmail);
  }
}


/**
 * Main function for triggering account deletion process.
 * @param {string} leaverID - leaver's SQL record primary key;
 * @param {string} adminEmail - email address of the IT admin approving the account suspension.
 */
function approveLeaverDeletion(leaverID, adminEmail) {
  var leaver = app.models.Leavers.getRecord(leaverID);
  var gsuiteUser = AdminDirectory.Users.get(leaver.RenamedEmail);
  
  try {
    console.log('Deleting User %s...', leaver.RenamedEmail);
    gsuiteUser = AdminDirectory.Users.remove(leaver.RenamedEmail); //################## Comment out for testing ##############
    console.log('User deleted');
  } catch (e) {
    console.error(JSON.stringify(e));
    generateEmail_('error', leaver, 'approveLeaverDeletion function error: ' + JSON.stringify(e));
  }
  
  if (gsuiteUser === null) { //################## Comment out for testing ##############
    leaver.Status = 'deleted';
    leaver.DeletionDate = new Date();
    app.saveRecords([leaver]);

    //*********** Notify IT ************
    generateEmail_('leaverDeleted', leaver, adminEmail);
  }
}