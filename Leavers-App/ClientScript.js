
/**
 * Filters the list in the admin view according to selected status in the drop down
 * @param {string} status - Status to filter
 */
function filterAdminListByStatus(status) {
  var datasource = app.datasources.Leavers;
  var accordion = app.pages.Admin.descendants.LeaversAccordion;
  var rows = accordion.children._values;
  
  for (var i=0; i<rows.length; i++) {
    //console.log('row data: ' + rows[i].datasource.item.Email);
    if (status !== null) {
      if (rows[i].datasource.item.Status != status)
        rows[i].visible = false;
      else
        rows[i].visible = true;
    } else
      rows[i].visible = true;
  }
}



/**
 * Filters the list of leavers according to the selected tab on the main page
 * @param {Widget} widget - TabPanel widget
 * @param {number} tabIndex - selected tab index
 */
function filterLeaversList(widget, tabIndex) {
  switch (tabIndex) {
    case 0:
      app.pages.Main.descendants.PendingLeaversAccordion.visible = false;
      widget.datasource.query.clearFilters();
      widget.datasource.query.filters.Status._equals = 'pending';
      widget.datasource.load();
      break;
      
    case 1:
      app.pages.Main.descendants.SuspendedLeaversAccordion.visible = false;
      widget.datasource.query.clearFilters();
      widget.datasource.query.filters.Status._in = [
        'suspended',
        'restoreRequested'
      ];
      widget.datasource.load();
      break;
      
    case 2:
      app.pages.Main.descendants.DeletedAccountsList.visible = false;
      widget.datasource.query.clearFilters();
      widget.datasource.query.filters.Status._equals = 'deleted';
      widget.datasource.query.filters.DeletionDate._greaterThanOrEquals = new Date(new Date().setDate(new Date().getDate()-30));
      widget.datasource.load();
      break;
      
    default:
      return;
  }
}


function filterSuspendedAccordionList(checkbox) {
  var accordion = app.pages.Main.descendants.SuspendedLeaversAccordion;
  var rows = accordion.children._values;
  
  for (var i=0; i<rows.length; i++) {
    if (checkbox) {
      if (calculateDaysRemaining(rows[i].datasource.item) > 0)
        rows[i].visible = false;
      else
        rows[i].visible = true;
    } else
      rows[i].visible = true;
  }
}


/**
* Creates a new user record in the NewLeaver database when an item is clicked in the User Picker. 
* Updates the labels in the form with the selected data and clears that record.
* @param {Widget} widget - the userPicker widget
* @param {Object} newValue - the generated object from the selected userPicker list item
*/
function userPickerClicked(widget, newValue) {
  var user = widget.datasource.item;
  user.Email = (newValue) ? newValue.PrimaryEmail : null;
  user.Name = (newValue) ? newValue.FullName : null;
  user.PhotoUrl = (newValue) ? newValue.ThumbnailPhotoUrl : null;

  if (user.Email) {
    var widgets;
    if (widget.root == app.pages.Main) {
      widgets = app.pages.Main.descendants;
      widgets.LeaverWarningLabel.visible = false;
      widgets.LeaverNameLabel.text = user.Name;
      widgets.LeaverEmailLabel.text = user.Email;
      widgets.LeaveDateBox.enabled = true;
      widgets.RetentionPolicyDropDown.enabled = true;
      widgets.LeaverCommentsTextArea.enabled = true;
      widgets.SubmitLeaverButton.enabled = true;
      widget.datasource.clearChanges();
      widget.value = null;
      
      var leavers = app.datasources.Leavers.items;
      for (var i=0; i<leavers.length; i++) {
        if (leavers[i].Status == 'pending') {
          if (leavers[i].Email == user.Email) {
            widgets.LeaverWarningLabel.visible = true;
            widgets.SubmitLeaverButton.enabled = false;
          }
        }
      }
      
    } else if (widget.root == app.popups.ApproveSuspensionDialog) {
      widgets = app.popups.ApproveSuspensionDialog.descendants;
      widgets.LeaverNameLabel.text = user.Name;
      widgets.LeaverEmailLabel.text = user.Email;
      widget.datasource.clearChanges();
      widget.value = null;
    }
  }
}

/**
 * Creates and saves a new Leaver record in the SQL database
 * @param {Widget} widget - Submit button
 */
function submitLeaverButtonClicked(widget) {
  var widgets = app.pages.Main.descendants;
  if (!widget.parent.validate())
    return;
  
  widgets.SubmitLeaverButton.enabled = false;
  
  var myDatasource = app.datasources.Leavers;
  var leaver = myDatasource.modes.create.item;
    
  leaver.FullName = widgets.LeaverNameLabel.text;
  leaver.Email = widgets.LeaverEmailLabel.text;
  leaver.Comments = widgets.LeaverCommentsTextArea.value;
  leaver.LeaveDate = widgets.LeaveDateBox.value;
  leaver.RetentionPolicy = widgets.RetentionPolicyDropDown.value;
  leaver.RequesterEmail = app.user.email;
  leaver.Status = "pending";
  
  widgets.PendingLeaversAccordion.visible = false;
  widgets.SuspendedLeaversAccordion.visible = false;
  widgets.DeletedAccountsList.visible = false;
  
  myDatasource.createItem(function(newRecord) {
    myDatasource.saveChanges(function() {
      var leaverID = newRecord._key;

      myDatasource.clearChanges();
      filterLeaversList(app.pages.Main.descendants.TabsPanel, app.pages.Main.descendants.TabsPanel.selectedTab);
      
      widgets.LeaverNameLabel.text = 'Name';
      widgets.LeaverEmailLabel.text = 'Email';
      widgets.LeaverCommentsTextArea.value = '';
      widgets.LeaveDateBox.value = null;
      widgets.RetentionPolicyDropDown.value = null;
      widgets.LeaveDateBox.enabled = false;
      widgets.RetentionPolicyDropDown.enabled = false;
      widgets.LeaverCommentsTextArea.enabled = false;
      widgets.LeaverWarningLabel.visible = false;
      

      //*********** Notify IT ************
      if (!app.user.role.Admins)
        google.script.run.emailFromClientScript('newLeaver', leaverID, app.user.email);
      });
 });
}





/**
 * Saves edits to selected pending Leaver.
 * Keeps the popup visible until database has processed the updated record
 * @param {Widget} widget - Save changes button
 */
function saveLeaverChangesButtonClicked(widget) {
  widget.root.descendants.EditLeaverSpinner.visible = true;
  widget.root.descendants.CancelButton.enabled = false;
  widget.enabled = false;
  widget.datasource.saveChanges(function(){
    widget.datasource.clearChanges(); //This line is to clear any changes that might linger after saving. The app can't reload the database when switching tabs otherwise.
    widget.root.visible = false;
    widget.root.descendants.EditLeaverSpinner.visible = false;
    widget.root.descendants.CancelButton.enabled = true;
    widget.enabled = true;
    widget.datasource.load();
    
    //*********** Notify IT ************
    if (!app.user.role.Admins)
      google.script.run.emailFromClientScript('leaverEdited', widget.datasource.item._key, app.user.email);
  });
}

/**
 * Deletes selected pending Leaver from the database.
 * Keeps the popup visible until database has processed the updated record
 * @param {Widget} widget - Delete leaver button
 */
function cancelPendingLeaverButtonClicked(widget) {
  var widgets = app.popups.CancelLeaverConfirmationDialog.descendants;
  widgets.CancelPendingLeaverDialogSpinner.visible = true;
  widgets.CancelButton.enabled = false;
  widgets.ConfirmButton.enabled = false;
  
  if (!app.user.role.Admins) {
    //*********** Notify IT then delete the user ************
    google.script.run
      .withSuccessHandler(function() {
      widget.datasource.deleteItem();
      widget.datasource.saveChanges(function(){
        widget.datasource.clearChanges(); //This line is to clear any changes that might linger after saving. The app can't reload the database when switching tabs otherwise.
        widget.root.visible = false;
        widgets.CancelPendingLeaverDialogSpinner.visible = false;
        widgets.CancelButton.enabled = true;
        widgets.ConfirmButton.enabled = true;
      });
    })
      .emailFromClientScript('pendingLeaverCancelled', widget.datasource.item._key, app.user.email); 
  } else {
    widget.datasource.deleteItem();
    widget.datasource.saveChanges(function(){
      widget.datasource.clearChanges(); //This line is to clear any changes that might linger after saving. The app can't reload the database when switching tabs otherwise.
      widget.root.visible = false;
      widgets.CancelPendingLeaverDialogSpinner.visible = false;
      widgets.CancelButton.enabled = true;
      widgets.ConfirmButton.enabled = true;
    });
  }
}




/**
 * Saves the custom options to the Leaver record, changes its status and calls for the server script to process the account suspension.
 * @param {Widget} widget - Approve Suspension button
 */
function approveSuspensionButtonClicked(widget) {
  var widgets = app.popups.ApproveSuspensionDialog.descendants;
  
  if (widgets.DriveTransferOptionCheckbox.value) {
    if (widgets.LeaverEmailLabel.text.indexOf('@companyName.com') != -1) {
      widget.datasource.item.DriveTransferEmail = widgets.LeaverEmailLabel.text;
    } else {
      return;
    }
  } else {
    widget.datasource.item.DriveTransferEmail = 'leavers@companyName.com';
  }
  
  if (widgets.DepartureGroupOptionCheckbox.value) {
    if (widgets.DepartureGroupDropdown.validate()) {
      widget.datasource.item.DeparturesGroupEmail = widgets.DepartureGroupDropdown.value;
    } else {
      return;
    }
  }
  
  widgets.ApproveSuspensionSpinner.visible = true;
  widgets.ApproveSuspensionButton.enabled = false;
  widgets.CancelButton.enabled = false;
  widgets.DriveTransferOptionCheckbox.enabled = false;
  widgets.DepartureGroupOptionCheckbox.enabled = false;
  widgets.gamCalendarCheckbox.enabled = false;
  widget.datasource.item.Status = 'suspensionInProgress';
  widget.datasource.item.SuspensionApproverEmail = app.user.email;

  widget.datasource.saveChanges(function() {
    google.script.run
    .withSuccessHandler(function() {
      filterLeaversList(app.pages.Main.descendants.TabsPanel, app.pages.Main.descendants.TabsPanel.selectedTab);
      hideAndResetApproveSuspensionPopup();
    })
    .withFailureHandler(function(error){
      //################ Add error handling function ###################
      app.popups.ErrorSuspensionDialog.descendants.ErrorDescriptionLabel.text = error.message;
      app.popups.ErrorSuspensionDialog.visible = true;
      hideAndResetApproveSuspensionPopup();
    })
    .approveLeaverSuspension(widget.datasource.item._key, app.user.email);
  });
}

/**
 * Shows the appropriate popup according to the status of the selected leaver and the current user priviledges:
 * - Leaver is "suspended" - show RestoreRequestDialog
 * - Leaver is "restoreRequested" and user is not Admin - show CancelRequestConfirmationDialog
 * - Leaver is "restoreRequested" and user is Admin - show ProcessRestoreDialog
 * @param {Widget} widget - RestoreSuspendedAccount button
 */
function restoreSuspendedLeaverButtonClicked(widget) {
  if (widget.datasource.item.Status == 'suspended') {
    app.popups.RestoreRequestDialog.visible = true;
  } else if (widget.datasource.item.Status == 'restoreRequested') {
    if (app.user.roles.indexOf('Admins') >= 0) { //if user is admin
      app.popups.ProcessRestoreRequestDialog.visible = true;
    } else {
      app.popups.CancelRestoreRequestConfirmationDialog.visible = true;
    }
  } else {
    return;
  }
}

/**
 * Changes status of selected Leaver to "restoreRequested", saves provided info to Leaver record and notifies IT.
 * Keeps the popup visible until database has processed the updated record
 * @param {Widget} widget - Save changes button
 */
function submitRestoreRequestButtonClicked(widget) {
  var widgets = app.popups.RestoreRequestDialog.descendants;
  if (!(!widgets.RestoreCommentsTextArea.validate() || !widgets.PersonalEmailTextBox.validate() || !widgets.RestoreRequestDateBox.validate())) {
    
    widgets.RestoreRequestSpinner.visible = true;
    widgets.CancelButton.enabled = false;
    widgets.SubmitRestoreRequestButton.enabled = false;
    
    widget.datasource.item.Status = 'restoreRequested';
    widget.datasource.item.RestoreRequesterEmail = app.user.email;

    widget.datasource.saveChanges(function() {
      widget.datasource.clearChanges(); //This line is to clear any changes that might linger after saving. The app can't reload the database when switching tabs otherwise.
      hideAndResetRestoreRequestPopup();
      
      //*********** Notify IT ************
      google.script.run.emailFromClientScript('restoreRequestSubmitted', widget.datasource.item._key, app.user.email);
    });
  }
}

/**
 * Changes status of selected Leaver to "suspended", deletes info relevant to the restore request from Leaver record and notifies IT.
 * Keeps the popup visible until database has processed the updated record
 * @param {Widget} widget - confirm button
 */
function confirmCancelRestoreButtonClicked(widget) {
  var widgets = app.popups.CancelRestoreRequestConfirmationDialog.descendants;
  widgets.CancelRestoreRequestDialogSpinner.visible = true;
  widgets.CancelButton.enabled = false;
  widgets.ConfirmButton.enabled = false;
  
  widget.datasource.item.Status = 'suspended';
  widget.datasource.item.RestoreRequestComments = null;
  widget.datasource.item.PersonalEmail = null;
  
  widget.datasource.saveChanges(function() {
      widget.datasource.clearChanges(); //This line is to clear any changes that might linger after saving. The app can't reload the database when switching tabs otherwise.
      widgets.CancelRestoreRequestDialogSpinner.visible = false;
      widgets.CancelButton.enabled = true;
      widgets.ConfirmButton.enabled = true;
      
      //*********** Notify IT ************
      google.script.run.emailFromClientScript('restoreRequestCancelled', widget.datasource.item._key, app.user.email);
    });
}

/**
 * Changes status of selected Leaver to "restoreInProgress" and calls the server script to process the account restore.
 * Keeps the popup visible until server script has executed
 * @param {Widget} widget - Approve button
 */
function approveRestoreButtonClicked(widget) {
  var widgets = app.popups.ProcessRestoreRequestDialog.descendants;
  
  widgets.ApproveRestoreButton.enabled = false;
  widgets.CancelButton.enabled = false;
  widgets.DenyRestoreButton.enabled = false;
  widgets.ProcessRestoreSpinner.visible = true;
  
  widget.datasource.item.Status = 'restoreInProgress';
  widget.datasource.item.RestoreApproverEmail = app.user.email;
  
  widget.datasource.saveChanges(function() {
    google.script.run
    .withSuccessHandler(function() {
      widget.datasource.load();
      hideAndResetProcessRestorePopup();
    })
    .withFailureHandler(function(error) {
      app.popups.ErrorRestoreDialog.descendants.ErrorDescriptionLabel.text = error.message;
      app.popups.ErrorRestoreDialog.visible = true;
      hideAndResetProcessRestorePopup();
    })
    .approveLeaverRestore(widget.datasource.item._key, app.user.email);
  });
}

/**
 * Changes status of selected Leaver to "suspended", deletes info relevant to the restore request from Leaver record and notifies the requester.
 * Keeps the popup visible until database has processed the updated record
 * @param {Widget} widget - confirm button
 */
function denyRestoreButtonClicked(widget) {
  var widgets = app.popups.ProcessRestoreRequestDialog.descendants;
  
  widgets.ApproveRestoreButton.enabled = false;
  widgets.CancelButton.enabled = false;
  widgets.DenyRestoreButton.enabled = false;
  widgets.ProcessRestoreSpinner.visible = true;
  
  widget.datasource.item.Status = 'suspended';
  widget.datasource.item.RestoreRequestComments = null;
  widget.datasource.item.PersonalEmail = null;
  widget.datasource.item.RestoreRequestDate = null;
  
  widget.datasource.saveChanges(function() {
    hideAndResetProcessRestorePopup();
    
    //*********** Notify IT ************
      google.script.run.emailFromClientScript('restoreRequestDenied', widget.datasource.item._key, app.user.email);
  });
}

/**
 * Changes status of selected Leaver to "deletionInProgress" and calls the server script to process the account deletion.
 * Keeps the popup visible until server script has executed
 * @param {Widget} widget - Approve button
 */
function confirmDeleteButtonClicked(widget) {
  var widgets = app.popups.ApproveDeletionConfirmationDialog.descendants;
  
  widgets.CancelButton.enabled = false;
  widgets.ConfirmButton.enabled = false;
  widgets.ApproveDeletionDialogSpinner.visible = true;
  
  widget.datasource.item.Status = 'deletionInProgress';
  widget.datasource.item.DeletionApproverEmail = app.user.email;
  
  widget.datasource.saveChanges(function() {
    google.script.run
    .withSuccessHandler(function() {
      widget.datasource.load();
      widgets.ApproveDeletionDialogSpinner.visible = false;
      widgets.CancelButton.enabled = true;
      widgets.ConfirmButton.enabled = true;
      widget.root.visible = false;
    })
    .approveLeaverDeletion(widget.datasource.item._key, app.user.email);
  });
}