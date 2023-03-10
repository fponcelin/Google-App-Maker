/**
 * Resolves icon URL for status.
 * @param {string} status - selected status;
 * @return {string} URL for status icon.
 */
function getIconForStatus(status) {
  switch (status) {
  case "Pending":
    return "resources/google-apps-script-[UUID]/0v1SNJ1HwBIM.png";
    
  case "Approved":
    return "resources/google-apps-script-[UUID]/ugoKrnkuruyY.png";
    
  case "Denied":
    return "resources/google-apps-script-[UUID]/4i4ffit68CNL.png";
    
  default:
    return;
  }
}

/**
 * Resolves group owners.
 * @param {string} ownersMembersJSONString - group owners/members stringified JSON;
 * @return {string} list of owners.
 */
function getGroupRequestOwners(ownersMembersJSONString) {
  var ownersMembersJSON = JSON.parse(ownersMembersJSONString);
  var ownersString = '';
  
  for (var i=0; i < ownersMembersJSON.owners.length; i++) {
    ownersString += ownersMembersJSON.owners[i].email + ' ';
  }
  return ownersString;
}

/**
 * Resolves group members.
 * @param {string} ownersMembersJSONString - group owners/members stringified JSON;
 * @return {string} list of members.
 */
function getGroupRequestMembers(ownersMembersJSONString) {
  var ownersMembersJSON = JSON.parse(ownersMembersJSONString);
  var membersString = '';
  
  for (var i=0; i < ownersMembersJSON.members.length; i++) {
    membersString += ownersMembersJSON.members[i].email + ' ';
  }
  return membersString;
}