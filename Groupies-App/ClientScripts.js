/**
* Checks availability of provided address handle in the companyName Directory
* Displays results in the AvailabilityLabel
*
* @param {string} emailHandle - email handle string to be tested
*/
function checkAvailable(emailHandle) {
  var email = emailHandle + '@companyName.com';
  app.pages.RequestGroup.descendants.AvailabilityLabel.text = '';
  app.pages.RequestGroup.descendants.Spinner1.visible = true;
  
  google.script.run
  .withSuccessHandler(function(isAvailable) {
    app.pages.RequestGroup.descendants.Spinner1.visible = false;
    if (isAvailable) {
      app.pages.RequestGroup.descendants.AvailabilityLabel.text = email + " is available. Hurray!";
      app.pages.RequestGroup.descendants.AvailabilityLabel.styles = ["greenLabel"];
      app.pages.RequestGroup.properties.IsEmailValid = true;
    } else {
      app.pages.RequestGroup.descendants.AvailabilityLabel.text = email + " is already taken. Try something else.";
      app.pages.RequestGroup.descendants.AvailabilityLabel.styles = ["redLabel"];
      app.pages.RequestGroup.properties.IsEmailValid = false;
    }
  })
  .scanAllUsers(email);
}

/**
* Clears the temporary databases of owners and members
*/
function clearTempDatasources() {
  app.datasources.GroupMembers.load(function() {
    if (app.datasources.GroupMembers.items.length > 0) {
      for (var i = app.datasources.GroupMembers.items.length - 1; i >= 0; i--) {
        app.datasources.GroupMembers.items[i]._delete();
        
      }
    }
    app.datasources.GroupMembers.saveChanges();
  });
  
  app.datasources.GroupOwners.load(function() {
    if (app.datasources.GroupOwners.items.length > 0) {
      for (var i = app.datasources.GroupOwners.items.length - 1; i >= 0; i--) {
        app.datasources.GroupOwners.items[i]._delete();
        
      }
    }
    app.datasources.GroupOwners.saveChanges();
  });
}

/**
* Creates and saves a new user record in the owners/members databases when an item is clicked in the User Picker
*/
function userPickerClicked(widget, newValue) {
  var user = widget.datasource.item;
  user.Email = (newValue) ? newValue.PrimaryEmail : null;
  user.Name = (newValue) ? newValue.FullName : null;
  user.PhotoUrl = (newValue) ? newValue.ThumbnailPhotoUrl : null;

  if (user.Email) {
    console.log("user email: " + user.Email);
    widget.datasource.saveChanges();
    widget.value = null;
    
  }
}


/**
* Resets the request form and clears the fields
*/
function clearRequestForm() {
  var widgets = app.pages.RequestGroup.descendants;
  widgets.GroupNameTextBox.value = null;
  widgets.GroupEmailTextBox.value = null;
  widgets.DescriptionTextArea.value = null;
  widgets.IsExternalRadioGroup.value = false;
  widgets.SubmitRequestButton.enabled = true;
  widgets.CreateTDCheckbox.value = false;
  widgets.CreateCalCheckbox.value = false;
}

/**
* Validates the fields in the form, checks that the requested email address has been validated by the user and that there's at least an owner or a member in the relevant lists.
*
* @return {bool} false if validation fails, true otherwise
*/
function validateRequestForm() {
  if (!app.pages.RequestGroup.properties.IsEmailValid) {
    app.pages.RequestGroup.descendants.AvailabilityLabel.text = '<- Make sure the requested address is available';
    app.pages.RequestGroup.descendants.AvailabilityLabel.styles = ['redLabel'];
    
    app.pageFragments.InfoDialog.descendants.TitleLabel.text = 'Check availability';
    app.pageFragments.InfoDialog.descendants.MessageLabel.text = 'Please click the button to check availability of your desired group address before submitting.';
    app.pageFragments.InfoDialog.descendants.MessageLabel1.text = '';
    app.showDialog(app.pageFragments.InfoDialog);
    
    return false;
  }
  if (!(app.datasources.GroupOwners.items.length > 0 || app.datasources.GroupMembers.items.length > 0)) {
    app.pageFragments.InfoDialog.descendants.TitleLabel.text = 'Missing data';
    app.pageFragments.InfoDialog.descendants.MessageLabel.text = 'Please add at least one Owner or one Member to your group.';
    app.pageFragments.InfoDialog.descendants.MessageLabel1.text = '';
    app.showDialog(app.pageFragments.InfoDialog);
    
    return false;
  }
  if (!(app.pages.RequestGroup.validate())) {
    app.pageFragments.InfoDialog.descendants.TitleLabel.text = 'NOPE!';
    app.pageFragments.InfoDialog.descendants.MessageLabel.text = 'Something\'s missing, but I\'m totally unhelpful and won\'t tell you what!';
    app.pageFragments.InfoDialog.descendants.MessageLabel1.text = '';
    app.showDialog(app.pageFragments.InfoDialog);
    
    return false;
  }
  
  return true;
}

/**
* Creates a request object with data from the form, including a stringified JSON with group owners and members.
* Calls the createGroupRequest() server script, passing the request object as parameter.
* In the success callback: gets the group email address passed in, calls clearRequestForm(), clearTempDataSources(), updates the labels in InfoDialog and show the dialog
*/
function submitGroupRequest(widget) {
  widget.enabled = false;
  var ownersMembersJSON = {"owners":[], "members":[]};
  var request = {
    groupName : app.pages.RequestGroup.descendants.GroupNameTextBox.value,
    groupEmail : app.pages.RequestGroup.descendants.GroupEmailTextBox.value + '@companyName.com',
    description : app.pages.RequestGroup.descendants.DescriptionTextArea.value,
    isExternal : app.pages.RequestGroup.descendants.IsExternalRadioGroup.value,
    requesterEmail : app.user.email,
    teamDrive : app.pages.RequestGroup.descendants.CreateTDCheckbox.value,
    calendar : app.pages.RequestGroup.descendants.CreateCalCheckbox.value
  };
  
  for (var i = 0; i < app.datasources.GroupOwners.items.length; i++) {
    ownersMembersJSON.owners.push({"email":app.datasources.GroupOwners.items[i].Email});
  }
  
  for (i = 0; i < app.datasources.GroupMembers.items.length; i++) {
    var isOwner = false;
    
    for (var j = 0; j < ownersMembersJSON.owners.length; j++) {
      if (ownersMembersJSON.owners[j].email == app.datasources.GroupMembers.items[i].Email)
        isOwner = true;
    }
    
    if (!isOwner)
      ownersMembersJSON.members.push({"email":app.datasources.GroupMembers.items[i].Email});
  }
  
  request.ownersMembersJSON = JSON.stringify(ownersMembersJSON);
  
  google.script.run
  .withSuccessHandler(function(groupEmail) {
    clearRequestForm();
    clearTempDatasources();
    app.pageFragments.InfoDialog.descendants.TitleLabel.text = 'Thank you!';
    app.pageFragments.InfoDialog.descendants.MessageLabel.text = 'Your request for ' + groupEmail + ' has been received and will be processed by our wonderful IT team ^_^';
    app.pageFragments.InfoDialog.descendants.MessageLabel1.text = 'You will receive an email when we approve/deny your request. Stay tuned...';
    app.showDialog(app.pageFragments.InfoDialog);
    })
  .createGroupRequest(request);
}


/**
* Approves the group request.
* Disables both buttons on the current row and calls the approveGroupRequest server script.
*/
function approveRequestButtonPushed(widget) {
  var requestID = widget.datasource.item._key;
  widget.enabled = false;
  widget.parent.descendants.DenyButton.enabled = false;

  google.script.run
  .withSuccessHandler(function() {
    app.datasources.GroupRequests.load();
  })
  .withFailureHandler(function(error) {
    app.popups.ErrorDialog.descendants.ErrorDescriptionLabel.text = error.message;
    app.popups.ErrorDialog.visible = true;
  })
  .approveGroupRequest(requestID, app.user.email);
}


function denyRequestButtonPushed(widget) {
  var requestID = widget.datasource.item._key;
  var denyDialog = app.pageFragments.DenyDialog;
  denyDialog.properties.GroupEmail = widget.datasource.item.GroupEmail;
  denyDialog.properties.GroupID = requestID;
  app.showDialog(denyDialog);
}


function loadSelectedGroupData(widget) {
  app.pages.GroupsBrowser.descendants.GroupNameLabel.text = widget.datasource.item.Name;
  app.pages.GroupsBrowser.descendants.GroupAddressLabel.text = widget.datasource.item.Address;
  app.pages.GroupsBrowser.descendants.GroupDescriptionLabel.text = widget.datasource.item.Description;
  
  app.pages.GroupsBrowser.descendants.MembersList.datasource.query.parameters.groupKey = widget.datasource.item.Address;
  app.pages.GroupsBrowser.descendants.MembersList.datasource.load();
}

function searchBoxEdit(widget) {
  var listWidget;
  
  switch (widget.name) {
    case "MyGroupsSearchBox": 
      listWidget = app.pages.GroupsBrowser.descendants.MyGroupsList;
      break;
      
    case "AllGroupsSearchBox":
      listWidget = app.pages.GroupsBrowser.descendants.AllGroupsList;
      break;
      
    default:
      return;
  }
  
  var rows = listWidget.children._values;
  for (var i=0; i<rows.length; i++) {   
    var widgets = rows[i].children._values;
    
    var address = widgets[0].text;
    var addressHandle = address.substring(0, address.indexOf("@"));
    
    if (addressHandle.indexOf(widget.value) == -1 && widget.value !== null)
      rows[i].visible = false;
    else
      rows[i].visible = true;
    
  }
  
}


