/**
* Generates an array of mailing group addresses
*
* @param {string} newbieOrg - Organisation unit
* @param {string} newbieTeam - Team email base handle
* @param {string} contractType - Contract type
*
* @return [string] groups - Array of mailing groups addresses
*/
function getNewbieGroups(newbieOrg, newbieTeam, contractType, workLocation) {
  var groups = [];
  var isStandardTeam = true;
  var isEmployee = false;
  var isIntern = false;
  var isFreelancer = false;
  
  //Check if newbie is Employee
  if (contractType.indexOf("Permanent") != -1)
    isEmployee = true;
  
  //Check if newbie is Freelancer
  else if (contractType.indexOf("Freelancer") != -1)
    isFreelancer = true;
  
  //Check if newbie is intern
  else if (contractType.indexOf("Intern") != -1)
      isIntern = true;
  
  //If newbie team or contract is "other", don't assign any group
  if (newbieTeam == 'other' || contractType.indexOf("Other") != -1)
    isStandardTeam = false;
  
  switch (newbieOrg) {
      
    case 'AAA':
      if (isEmployee || isIntern)
        groups.push('aaa@companyName.com');
      
      isStandardTeam = false;
      break;
      
    case 'BBB':
      if (isEmployee || isIntern)
        groups.push('bbb@companyName.com');
      else if (isFreelancer)
        groups.push('freelancers.bbb@companyName.com');
      if (!(newbieTeam == 'finance' || newbieTeam == 'people'))
        isStandardTeam = false;
      if (newbieTeam == 'artists')
        groups.push('artists@companyName.com');
      break;
      
    case 'CCC':
      if (isEmployee || isIntern)
        groups.push('ccc@companyName.com');
      else if (isFreelancer) {
        groups.push('freelancers.ccc@companyName.com');
        if (newbieTeam == 'design') {
          isStandardTeam = false;
          groups.push('design.ccc.temp@companyName.com');
        }
      }
      if (newbieTeam == 'delivery') {
        isStandardTeam = false;
        if (isEmployee)
          groups.push('delivery.ccc.perm@companyName.com');
        else
          groups.push('delivery.ccc.temp@companyName.com');
      }
      
      break;
      
    case 'DDD':
      if (isEmployee)
        groups.push('ddd@companyName.com');
      else if (isFreelancer) {
        groups.push('freelancers.ddd@companyName.com');
        if (newbieTeam == 'design') {
          groups.push('design.ddd.temp@companyName.com');
          isStandardTeam = false;
        }
      if (newbieTeam == 'developers') {
          groups.push('developers.ddd.temp@companyName.com');
          isStandardTeam = false;
        }  
      } else if (isIntern)
        groups.push('interns.ddd@companyName.com');
      break;
      
    case 'EEE':
      if (isEmployee || isIntern)
        groups.push('eee@companyName.com');
      else if (isFreelancer)
        groups.push('freelancers.eee@companyName.com');
      break;
      
    case 'FFF':
      if (isEmployee || isIntern)
        groups.push('fff@companyName.com');
      else if (isFreelancer)
        groups.push('freelancers.fff@companyName.com');
      break;
    
    default:
      break;
  }
  
  if (isStandardTeam)
    groups.push(newbieTeam + '.' + newbieOrg.toLowerCase() + '@companyName.com');
  
  console.log(groups);
  return groups;
}


/**
* Generates an array of eligible companyName addresses using data entered in the request form
* Current address patterns: firstname@, lastname@, personalemailaddress@, firstnameinitiallastname@
*
* @returns [string] emails - array of email addresses
*/
function generateEmails() {
  var widgets = app.pages.RequestNewbieV2.descendants;
  var emails = [];
  
  var firstName = removeDiacritics(widgets.FirstNameTextBox.value.toLowerCase().trim());
  firstName = firstName.replace(/ /g, "");
  emails.push(firstName + '@companyName.com');
  
  var lastName = removeDiacritics(widgets.LastNameTextBox.value.toLowerCase().trim());
  lastName = lastName.replace(/ /g, "");
  emails.push(lastName + '@companyName.com');
  
  var personalEmailHandle = widgets.PersonalEmailTextBox.value.slice(0, widgets.PersonalEmailTextBox.value.indexOf('@'));
  emails.push(personalEmailHandle + '@companyName.com');
  
  //var firstNameWithSurnameInitial = firstName + lastName.slice(0,1);
  
  var firstNameInitialWithSurname = firstName.slice(0,1) + lastName;
  emails.push(firstNameInitialWithSurname + '@companyName.com');
  
  return emails;
}

/**
* Loops through the pre-loaded database of currently used email addresses on the domain and checks each email provided against that list
* Also check against the list of "illegal" emails (emails that were used by sensitive people e.g. IT, leadership, etc).
*
* @param {string} newbieOrg - Organisation unit
* @param {string} newbieTeam - Team email base handle
* @param {string} contractType - Contract type
*
* @return [string] availableEmails - Array of emails with no match in the exisiting emails list
*/
function validateEmails(emails) {
  var availableEmails = [];
  
  var match = false;
  for (var i=0; i<emails.length; i++) {
    for (var j=0; j<app.datasources.ExistingEmails.items.length; j++) {
      if (emails[i] == app.datasources.ExistingEmails.items[j].Email)
        match = true;
    }
    
    if (!match) {
      for (var jj=0; jj<app.datasources.IllegalEmails.items.length; jj++) {
        if (emails[i] == app.datasources.IllegalEmails.items[jj].Email)
          match = true;
      }
    }
    
    if (!match) {
      for (var jjj=0; jjj<app.datasources.Newbies.items.length; jjj++) {
        if (emails[i] == app.datasources.Newbies.items[jjj].companyNameEmail)
          match = true;
      }
    }
    
    if (!match) {
      for (var jjjj=0; jjjj<app.datasources.RogoEmails.items.length; jjjj++) {
        if (emails[i] == app.datasources.RogoEmails.items[jjjj].Email)
          match = true;
      }
    }
    
    if (!match)
      availableEmails.push(emails[i]);
    match = false;
  }
  
  return availableEmails;
}

function populateEmailSuggestions(emails) {
  var datasource = app.datasources.SuggestedEmails.modes.create;
  
  for (var i=0; i<emails.length; i++) {
    var suggestion = datasource.item;
    suggestion.Email = emails[i];
    console.log('suggestedEmail: ' + emails[i]);
    datasource.createItem();
  }
}




/**
* Validates the newbie request form to ensure all fields are filled correctly
* Gets the list of mailing groups based on info provided and calls the function populateGroupsPanel to display the groups in the UI
* Adds the groups to the popup properties so they can be retrieved when creating the Newbie database record
* Generates and validates the email suggestions and calls the function populateEmailSuggestions to display the emails in the UI
* Displays the confirmation dialog
*
* @param {Widget} widget - Button widget used to submit the Newbie request form
*/
function reviewButtonClicked(widget) {
  var widgets = app.pages.RequestNewbieV2.descendants;
  var popupWidgets = app.popups.ConfirmRequestDialogV2A.descendants;
  
  if (!widgets.NewbieForm.validate())
    return;
  
  widgets.ReviewButton.enabled = false;
  
  var groups = getNewbieGroups(widgets.OrgDropdown.value, widgets.TeamDropdown.value, widgets.ContractRadioGroup.value, widgets.StudioDropdown.value);
  populateGroupsPanel(groups);
  app.popups.ConfirmRequestDialogV2A.properties.Groups = groups;
  

  var emails = generateEmails();
  emails = validateEmails(emails);
  populateEmailSuggestions(emails);
  
  
  app.popups.ConfirmRequestDialogV2A.visible = true;
  widgets.ReviewButton.enabled = true;
}

/**
* If the field is not empty, check its value against the list of existing email addresses
* Updates the UI accordingly
*
* @param {Widget} widget - Custom email text field
*/
function validateEmailInput(widget) {
  $(".selectedRowLabel").hide();
  var widgets = app.popups.ConfirmRequestDialogV2A.descendants;
  if (widget.value && widget.validate()) {
    if (validateEmails([widget.value + '@companyName.com']).length > 0) {
      widgets.EmailAvailableLabel.visible = true;
      widgets.EmailUnavailableLabel.visible = false;
      widgets.SelectedEmailLabel.text = widget.value + '@companyName.com';
      widgets.ConfirmButton.enabled = true;
    } else {
      widgets.EmailAvailableLabel.visible = false;
      widgets.EmailUnavailableLabel.visible = true;
      widgets.SelectedEmailLabel.text = 'This email is unavailable';
      widgets.ConfirmButton.enabled = false;
    }
  } else {
    widgets.SelectedEmailLabel.text = 'No email selected';
    widgets.ConfirmButton.enabled = false;
    widgets.EmailAvailableLabel.visible = false;
    widgets.EmailUnavailableLabel.visible = false;
  }
}

/**
* Creates the Newbie record on the database and calls for the server function to send the email notification
*/
function confirmNewbieRequestButtonClicked() {
  var widgets = app.pages.RequestNewbieV2.descendants;
  var popupWidgets = app.popups.ConfirmRequestDialogV2A.descendants;
  
  popupWidgets.ConfirmationSpinner.visible = true;
  popupWidgets.CancelButton.enabled = false;
  popupWidgets.ConfirmButton.enabled = false;
  
  var datasource = app.datasources.Newbies;
  var newbie = datasource.modes.create.item;
  
  newbie.FirstName = widgets.FirstNameTextBox.value;
  newbie.LastName = widgets.LastNameTextBox.value;
  newbie.PersonalEmail = widgets.PersonalEmailTextBox.value;
  newbie.StartDate = widgets.StartDateBox.value;
  newbie.EndDate = widgets.EndDateBox.value;
  newbie.Org = widgets.OrgDropdown.value;
  newbie.Team = widgets.TeamDropdown.value;
  newbie.Contract = widgets.ContractRadioGroup.value;
  newbie.Location = widgets.StudioDropdown.value;
  newbie.Notes = widgets.CommentsTextArea.value;
  newbie.NeedsLaptop = widgets.NeedsLaptopCheckbox.value;
  
  newbie.RequesterEmail = app.user.email;
  newbie.RequestDate = new Date();
  
  newbie.companyNameEmail = popupWidgets.SelectedEmailLabel.text;
  
  var groups = app.popups.ConfirmRequestDialogV2A.properties.Groups;
  var groupsJSON = {'emails': groups};
  
  newbie.Groups = JSON.stringify(groupsJSON);
  
  datasource.createItem(function(newbie) {
    datasource.saveChanges(function() {
      console.log('Newbie Key: ' + newbie._key);
      google.script.run
        .withSuccessHandler(function() {
          hideAndResetConfirmRequestPopup();
          resetRequestNewbieForm();
          app.popups.RequestSubmittedDialog.visible = true;
      })
        .newbieRequestSubmitted(newbie._key);
    });
  });
  
  
}

/**
* Checks if the email typed is available when editing a request.
*
* @param {Widget} widget - the email text field.
*/
function validateEmailEdit(widget) {
  var widgets = app.popups.EditNewbieDialog.descendants;
  
  if (widget.value && widget.validate()) {
    if (validateEmails([widget.value]).length > 0) {
      widgets.EmailAvailableLabel.visible = true;
      widgets.EmailUnavailableLabel.visible = false;
      widgets.EmailWarningLabel.visible = false;
    } else {
      if (app.user.role.Admins) {
        widgets.EmailAvailableLabel.visible = false;
        widgets.EmailUnavailableLabel.visible = false;
        widgets.EmailWarningLabel.visible = true;
      } else {
        widgets.EmailAvailableLabel.visible = false;
        widgets.EmailUnavailableLabel.visible = true;
        widgets.EmailWarningLabel.visible = false;
      }
    }
  } else {
    widgets.EmailAvailableLabel.visible = false;
    widgets.EmailUnavailableLabel.visible = true;
    widgets.EmailWarningLabel.visible = false;
  }
}

/**
* Discards changes applied to a Newbie request and close popup.
*/
function cancelEditNewbieButtonClicked() {
  var widgets = app.popups.EditNewbieDialog.descendants;
  
  app.datasources.Newbies.clearChanges();
  app.popups.EditNewbieDialog.visible = false;
  
  widgets.EmailAvailableLabel.visible = true;
  widgets.EmailUnavailableLabel.visible = false;
  widgets.EmailWarningLabel.visible = false;
}

/**
* Saves changes applied to a Newbie request and emails help@ to notify of the changes.
*
* @param {Widget} widget - the button clicked.
*/
function saveEditNewbieButtonClicked(widget) {
  var widgets = app.popups.EditNewbieDialog.descendants;
  widgets.Spinner1.visible = true;
  
  var groups = getNewbieGroups(widgets.OrgDropdown.value, widgets.TeamDropdown.value, widgets.ContractDropdown.value, widgets.WorkLocationDropdown.value);
  var groupsJSON = {'emails': groups};
  
  widget.datasource.item.Groups = JSON.stringify(groupsJSON);
  
  app.datasources.Newbies.saveChanges(function() {
    widgets.Spinner1.visible = false;
    app.popups.EditNewbieDialog.visible = false;
    
    if (!app.user.role.Admins) {
      google.script.run.newbieRequestEdited(app.user.email, widget.datasource.item._key);
    }
  });
}



/**
* Disables the buttons on the admin page and calls for the approveNewbieRequest server script. Reloads the datasource in the callback.
* If the user creation fails, the Error popup appears with the error contents displayed.
*
* @param {Widget} widget - the button clicked.
*/
function approveButtonClickedV2(widget) {
  var newbieID = widget.datasource.item._key;
  widget.enabled = false;
  widget.parent.descendants.DenyButton.enabled = false;
  widget.parent.descendants.EditButton.enabled = false;

  google.script.run
    .withSuccessHandler(function() {
      app.datasources.Newbies.load();
    })
    .withFailureHandler(function(error) {
      app.popups.ErrorDialog.descendants.ErrorDescriptionLabel.text = error.message;
      app.popups.ErrorDialog.visible = true;
    
      widget.enabled = true;
      widget.parent.descendants.DenyButton.enabled = true;
      widget.parent.descendants.EditButton.enabled = true;
      app.datasources.Newbies.load();
    })
    .approveNewbieRequestV2(newbieID, app.user.email);
}


function getTeamName(emailHandle) {
  if (!app.datasources.MailingGroups.loaded)
    return;
  
  var mailingGroups = app.datasources.MailingGroups.items;
  for (var i=0; i<mailingGroups.length; i++) {
    if (mailingGroups[i].EmailBaseHandle == emailHandle)
      return mailingGroups[i].Name;
  }
  return emailHandle;
}

/**
* Sets the properties of the DenyDialog page fragment and displays the dialog
*
* @param {Widget} widget - the button clicked.
*/
function denyButtonClicked(widget) {
  var newbieID = widget.datasource.item._key;
  var denyDialog = app.pageFragments.DenyDialog;
  denyDialog.properties.NewbieEmail = widget.datasource.item.companyNameEmail;
  denyDialog.properties.NewbieID = newbieID;
  app.showDialog(denyDialog);
}


/**
* Calls for the denyNewbieRequest server script. In the callback clear the textarea, reload the datasource and close the dialog.
*
* @param {Widget} widget - the button clicked.
*/
function denyAndSendButtonClicked(widget) {
  var newbieID = widget.parent.properties.NewbieID;
  var message = widget.parent.descendants.ResponseTextArea.value;

  google.script.run
    .withSuccessHandler(function() {
      widget.parent.descendants.ResponseTextArea.value = '';
      app.datasources.Newbies.load();
      app.closeDialog();
    })
    .withFailureHandler(function(error) {
      app.closeDialog();
      console.error('denyNewbieRequest error: ' + JSON.stringify(error));
    })
    .denyNewbieRequest(newbieID, message, app.user.email);
}