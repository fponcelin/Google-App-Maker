/**
 * Helper function to display the Leaver's retention policy in human readable format
 * @param {number} days - Leaver's retention policy in days
 * @return {string} value to display
 */
function convertRetentionPolicy(days) {
  switch (days) {
    case 0:
      return "Immediate deletion";
      
    case 90:
      return "90 days";
      
    case 365:
      return "1 year";
      
    case 2190:
      return "6 years";
      
    default:
      return "no retention policy - call IT!";
  }
}


function sliceOUString(orgUnitPath) {
  if (orgUnitPath)
    orgUnitPath = orgUnitPath.slice(1,4);
  
  return orgUnitPath;
}


function isDateFuture(date) {
  if (date > new Date())
    return true;
  return false;
}


function calculateDaysRemaining(leaver) {
  if (leaver) { //check if there is a valid record when function is called
    if (leaver.Status != 'pending') {
      var leaveDate = leaver.LeaveDate;
      var retentionPolicy = leaver.RetentionPolicy;
      var today = new Date();

      var remainingTimeInMilliseconds = leaveDate.getTime() + retentionPolicy * 86400000 - today.getTime();
      var remainingTimeInDays = Math.ceil(remainingTimeInMilliseconds / 86400000);
      return remainingTimeInDays;
    }
  }
}

/**
 * Hides the suspension approval popup and resets the fields and labels.
 */
function hideAndResetApproveSuspensionPopup() {
  app.popups.ApproveSuspensionDialog.visible = false;
  
  var widgets = app.popups.ApproveSuspensionDialog.descendants;
  widgets.DepartureGroupOptionCheckbox.value = false;
  widgets.DriveTransferOptionCheckbox.value = false;
  widgets.gamCalendarCheckbox.value = false;
  widgets.LeaverEmailLabel.text = 'Email';
  widgets.LeaverNameLabel.text = 'Name';
  widgets.ApproveSuspensionSpinner.visible = false;
  widgets.CancelButton.enabled = true; 
  widgets.DriveTransferOptionCheckbox.enabled = true;
  widgets.DepartureGroupOptionCheckbox.enabled = true;
  widgets.gamCalendarCheckbox.enabled = true;
}


/**
 * Hides the restore request popup, clears unsaved changes to datasource and resets the buttons.
 */
function hideAndResetRestoreRequestPopup() {
  app.popups.RestoreRequestDialog.visible = false;
  
  var widgets = app.popups.RestoreRequestDialog.descendants;
  widgets.RestoreRequestSpinner.visible = false;
  widgets.CancelButton.enabled = true;
  widgets.SubmitRestoreRequestButton.enabled = true;
  
  widgets.Content.datasource.clearChanges();
}

/**
 * Hides the Process restore request popup, clears unsaved changes to datasource and resets the buttons.
 */
function hideAndResetProcessRestorePopup() {
  app.popups.ProcessRestoreRequestDialog.visible = false;
  
  var widgets = app.popups.ProcessRestoreRequestDialog.descendants;
  widgets.CancelButton.enabled = true;
  widgets.ApproveRestoreButton.enabled = true;
  widgets.DenyRestoreButton.enabled = true;
  widgets.ProcessRestoreSpinner.visible = false;
  
  widgets.Content.datasource.clearChanges();
}


function copyText(text) { 
  var textArea = document.createElement('textarea');

  textArea.style.position = 'fixed';
  textArea.style.top = 0;
  textArea.style.left = 0;
  textArea.style.width = '2em';
  textArea.style.height = '2em';
  textArea.style.padding = 0;
  textArea.style.border = 'none';
  textArea.style.outline = 'none';
  textArea.style.boxShadow = 'none';
  textArea.style.background = 'transparent';
  textArea.value = text;
  document.body.appendChild(textArea);
  textArea.select();
  try {
    document.execCommand('copy');
    document.body.removeChild(textArea);
    console.log('copied');
    return true;
  } catch(e) {
    console.log(JSON.stringify(e));
    document.body.removeChild(textArea);
    return false;
  }
  
}