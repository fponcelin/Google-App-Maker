/**
* Loops through all @companyName.com users & groups primary email and aliases checking if a provided email address is available
*
* @param {string} email - email address to check
* @return {bool} false if address unavailable, true otherwise
*/
function scanAllUsers(email) {
  var pageToken, page;
  try {
    do {
      page = AdminDirectory.Groups.list({
        domain: 'companyName.com',
        maxResults: 100,
        pageToken: pageToken
      });
      var groups = page.groups;
      if (groups) {
        for (var i=0; i < groups.length; i++) {
          var group = groups[i];
          if (group.email == email)
            return false;

          if (group.aliases !== undefined) {
            if (group.aliases.indexOf(email) >= 0)
              return false;
          }
        }
      }
      pageToken = page.nextPageToken;
    } while (pageToken);
    
    do {
      page = AdminDirectory.Users.list({
        domain: 'companyName.com',
        maxResults: 100,
        pageToken: pageToken
      });
      var users = page.users;
      if (users) {
        for (var j=0; j < users.length; j++) {
          var user = users[j];
          if (user.primaryEmail == email)
            return false;

          if (user.aliases !== undefined) {
            if (user.aliases.indexOf(email) >= 0)
              return false;
          }
        }
      }
      pageToken = page.nextPageToken;
    } while (pageToken);

    return true;
  } catch(e) {
    console.error(JSON.stringify(e));
    generateEmail_({from: '', to: 'it@companyName.com', bcc: ''}, "error", {requestRecord: requestRecord, error: JSON.stringify(e)});
  }
}


/**
 * Generates UUID
 */
function generateUUID(){
    var dt = new Date().getTime();
    var uuid = 'xxxxxxxx-xxxx-4xxx-yxxx-xxxxxxxxxxxx'.replace(/[xy]/g, function(c) {
        var r = (dt + Math.random()*16)%16 | 0;
        dt = Math.floor(dt/16);
        return (c=='x' ? r :(r&0x3|0x8)).toString(16);
    });
    return uuid;
}

/**
* Creates a Team Drive and modify permissions to add group and remove helpdesk@companyName.com (because the app runs as Helpdesk and therefore adds it to the TD)
*
* @param {Datasource Record} requestRecord - Group request record
*/
function createTeamDrive_(requestRecord) {
  try {
    console.log('Creating Team Drive named ' + requestRecord.GroupName + '...');
    
    var td = Drive.Teamdrives.insert({'name': requestRecord.GroupName}, generateUUID());
    
    console.log('Team Drive data: ' + JSON.stringify(td));
    
    var tdPermissions = Drive.Permissions.list(td.id, {supportsTeamDrives: true});
    var appmakerPermissionID = tdPermissions.items[0].id; //This is the permission ID of the account executing the app (helpdesk@companyName.com)
    
    console.log('Modifying group permissions to Team Drive...');

    //Adding group as Content Manager to the Team Drive
    var tdPermissionResource = {
      "role": "fileOrganizer",
      "type": "user",
      "value": requestRecord.GroupEmail
    };

    Drive.Permissions.insert(tdPermissionResource, td.id, {supportsTeamDrives: true});
    
    //Adding group owners as Managers of the Team Drive
    var usersJSON = JSON.parse(requestRecord.OwnersMembersJSON);
    for (var i=0; i<usersJSON.owners.length; i++) {
      tdPermissionResource = {
        "role": "organizer",
        "type": "user",
        "value": usersJSON.owners[i].email
      };
      Drive.Permissions.insert(tdPermissionResource, td.id, {supportsTeamDrives: true});
    }

    Drive.Permissions.remove(td.id, appmakerPermissionID, {supportsTeamDrives: true});
    
  } catch (e) {
    console.error(JSON.stringify(e));
    generateEmail_({from: '', to: 'it@companyName.com', bcc: ''}, "error", {requestRecord: requestRecord, error: JSON.stringify(e)});
    throw new app.ManagedError('Failed to create Team Drive named ' + requestRecord.GroupName + ' (step 5).');
  }
}

/**
* Creates a Calendar and adds ACLs to add group as owner and domain as reader
*
* @param {Datasource Record} requestRecord - Group request record
*/
function createCalendar_(requestRecord) {
  try {
    console.log('Creating Calendar named %s...', requestRecord.GroupName);
    
    var calendar = Calendar.Calendars.insert({summary: requestRecord.GroupName});
    
    console.log('Calendar %s has been created', calendar.id);
    
    console.log('Adding Group owner ACL...');
    var aclRule = {
      role: 'owner',
      scope: {
        type: 'group',
        value: requestRecord.GroupEmail
      }
    };
        
    Calendar.Acl.insert(aclRule, calendar.id);
    
    console.log('Adding domain reader ACL...');
    aclRule = {
      role: 'reader',
      scope: {
        type: 'domain',
        value: 'companyName.com'
      }
    };
    
    Calendar.Acl.insert(aclRule, calendar.id);
    
    console.log('Unsubscribing helpdesk from Calendar...');
    Calendar.CalendarList.remove(calendar.id);
    
  } catch (e) {
    console.error(JSON.stringify(e));
    generateEmail_({from: '', to: 'it@companyName.com', bcc: ''}, "error", {requestRecord: requestRecord, error: JSON.stringify(e)});
    throw new app.ManagedError('Failed to create Calendar named ' + requestRecord.GroupName + ' (step 6).');
  }
}


/**
* Inserts a user in a mailing group
*
* @param {string} groupEmail - Email address of the group
* @param {User Resource} userResource - AdminDirectory user resource JSON of the newbie to be inserted in the group
*/
function insertMemberInGroup_(groupEmail, userEmail, userRole) {
  try {
    var userResource = AdminDirectory.Users.get(userEmail);

    var requestBody = {
      kind: "admin#directory#member",
      etag: userResource.etag,
      id: userResource.id,
      email: userEmail,
      role: userRole,
      type: 'USER'
    };
  
  
    var member = AdminDirectory.Members.insert(requestBody, groupEmail);
    console.log('New group member: %s in %s', member.email, groupEmail);
  } catch (e) {
    console.error(JSON.stringify(e));
    generateEmail_({from: '', to: 'it@companyName.com', bcc: ''}, "error", {requestRecord: requestRecord, error: JSON.stringify(e)});
    throw new app.ManagedError('Failed to insert group ' + userRole +': ' + userEmail + ' (step 3/4).');
  }
}

/**
* Updates a group's settings to allow external posts
*
* @param {Datasource Record} requestRecord - Group request record
*/
function updateGroupSettings_(requestRecord) {
  try {
  var group = AdminGroupsSettings.Groups.get(requestRecord.GroupEmail);
  group.whoCanPostMessage = requestRecord.IsExternal ? 'ANYONE_CAN_POST' : 'ALL_IN_DOMAIN_CAN_POST';
  group.whoCanViewGroup = 'ALL_MEMBERS_CAN_VIEW';
  group.whoCanViewMembership = 'ALL_IN_DOMAIN_CAN_VIEW';
  group.isArchived = true;
  group.showInGroupDirectory = true;
  AdminGroupsSettings.Groups.patch(group, requestRecord.GroupEmail);
  } catch(e) {
    console.error(JSON.stringify(e));
    generateEmail_({from: '', to: 'it@companyName.com', bcc: ''}, "error", {requestRecord: requestRecord, error: JSON.stringify(e)});
    throw new app.ManagedError('Failed to update settings for group ' + requestRecord.GroupEmail + ' (step 2).');
  }
}


/**
* Creates a mailing group then inserts the users in the group
*
* @param {Datasource Record} requestRecord - Group request record
*/
function createGoogleGroup_(requestRecord, adminEmail) {
  var group = {
    email: requestRecord.GroupEmail,
    name: requestRecord.GroupName,
    description: requestRecord.Description + ' - Requested by ' + requestRecord.RequesterEmail + ' on ' + requestRecord.RequestDate
  };
  
  try {
    group = AdminDirectory.Groups.insert(group);
    console.log('New group created: %s, approved by: %s', group.email, adminEmail);
  } catch(e) {
    console.error(JSON.stringify(e));
    generateEmail_({from: '', to: 'it@companyName.com', bcc: ''}, "error", {requestRecord: requestRecord, error: JSON.stringify(e)});
    throw new app.ManagedError('Failed to create group ' + requestRecord.GroupEmail + ' (step 1).');
  }
}


/**
* Approves the group request, calls for the Google group creation
*
* @param {int} requestID - Request record primary key
* @param {string} adminEmail - Email address of the admin approving the request
*/
function approveGroupRequest(requestID, adminEmail) {
  var requestRecord = app.models.GroupRequests.getRecord(requestID);
  
  requestRecord.ApprovalStatus = "Approved";
  requestRecord.IsActioned = true;
  requestRecord.ApproverEmail = adminEmail;
  app.saveRecords([requestRecord]);
  
  createGoogleGroup_(requestRecord, adminEmail);
  updateGroupSettings_(requestRecord);

  var ownersMembersJSON = JSON.parse(requestRecord.OwnersMembersJSON);

  for (var i=0; i < ownersMembersJSON.owners.length; i++) {
    insertMemberInGroup_(requestRecord.GroupEmail, ownersMembersJSON.owners[i].email, 'OWNER');
  }

  for (i=0; i < ownersMembersJSON.members.length; i++) {
    insertMemberInGroup_(requestRecord.GroupEmail, ownersMembersJSON.members[i].email, 'MEMBER');
  }

  if (requestRecord.TeamDrive)
    createTeamDrive_(requestRecord);

  if (requestRecord.Calendar)
    createCalendar_(requestRecord);

  generateEmail_({from: adminEmail, to: requestRecord.GroupEmail, bcc: 'it@companyName.com'}, 'groupCreated', {requestRecord: requestRecord});
}


/**
* Denies the group request, emails the requester with the admin's message
*
* @param {int} requestID - Request record primary key
* @param {string} message - Message to be emailed
* @param {string} adminEmail - Email address of the admin approving the request
*/
function denyGroupRequest(requestID, message, adminEmail){
  var requestRecord = app.models.GroupRequests.getRecord(requestID);
  
  requestRecord.ApprovalStatus = "Denied";
  requestRecord.IsActioned = true;
  requestRecord.ApproverEmail = adminEmail;
  app.saveRecords([requestRecord]);
  
  generateEmail_({from: adminEmail, to: requestRecord.RequesterEmail, bcc: "it@companyName.com"}, "denyRequest", {requestRecord: requestRecord, message: message});
}

/**
* Creates a request Record and saves it to the GroupRequests datasource
*
* @param {object} request - Javascript request object
* @return {Datasource Record} requestRecord - Group request record
*/
function createRequestRecord_(request) {
  var requestRecord = app.models.GroupRequests.newRecord();
  
  requestRecord.GroupName = request.groupName;
  requestRecord.GroupEmail = request.groupEmail;
  requestRecord.Description = request.description;
  requestRecord.IsExternal = request.isExternal;
  requestRecord.RequesterEmail = request.requesterEmail;
  requestRecord.OwnersMembersJSON = request.ownersMembersJSON;
  requestRecord.RequestDate = new Date();
  requestRecord.TeamDrive = request.teamDrive;
  requestRecord.Calendar = request.calendar;
  
  app.saveRecords([requestRecord]);
  
  return requestRecord;
}

/**
* Calls createRequestRecord with the passed in request object, then calls generateEmail_ passing in the created request record
*
* @param {object} request - Javascript request object
* @return {string} group email address
*/
function createGroupRequest(request) {
  var requestRecord = createRequestRecord_(request);
  generateEmail_({from: requestRecord.RequesterEmail, to: 'help@companyName.it', bcc: ''}, 'groupRequest', {requestRecord: requestRecord});
  generateEmail_({from: '', to: requestRecord.RequesterEmail, bcc: ''}, 'requestConfirmed', {requestRecord: requestRecord});
  
  return requestRecord.GroupEmail;
}



