function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('---->')
      .addItem('Click Keys to make a change', 'fakeMenuItem')
      .addToUi(); 
  ui.createMenu('Keys')
      .addItem('Add / Update Keys ', 'openFormDialog')
       .addItem('Search', 'searchData')
      
      .addToUi();
  ui.createMenu('<----')
      .addItem('Click Keys to make a change', 'fakeMenuItem')
      .addToUi();
}

function openFormDialog() {
  var html = HtmlService.createHtmlOutputFromFile('Form')
      .setWidth(500)
      .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, 'Enter Key Information');
}

function fakeMenuItem() {
  // This function can be left blank since it's a fake menu item
  // or you can add any desired functionality here
}

function processForm(formObject) {
  Logger.log("Form Object: " + JSON.stringify(formObject)); // Log the form object to check if it's received properly

  if (!formObject || !formObject.keyName || !formObject.signInOut || !formObject.person || !formObject.lockboxLocation) {
    Logger.log("Form data is missing or incomplete.");
    return "Form data is missing or incomplete."; // Return a message indicating missing or incomplete form data
  }

  var keyName = formObject.keyName;
  var signInOut = formObject.signInOut;
  var person = formObject.person;
  var lockboxLocation = formObject.lockboxLocation;

  Logger.log("Key Name: " + keyName); // Log the key name to check if it's received properly

  // Check if the key is restricted or non-restricted
  var isRestricted = /^[A-F]\d{1,2}$/i.test(keyName);
  Logger.log("Is Restricted Key: " + isRestricted); // Log whether the key is restricted or not

  // If the key is restricted, call the processRestrictedKeyForm function
  if (isRestricted) {
    Logger.log("Processing Restricted Key Form");
    processRestrictedKeyForm(formObject);
  } else {
    // If the key is non-restricted, call the processNonRestrictedKeyForm function
    Logger.log("Processing Non-Restricted Key Form");
    processNonRestrictedKeyForm(formObject);
  }
}





// Logic for processing restricted keys
function processRestrictedKeyForm(formObject) {
  Logger.log("Form Object: " + JSON.stringify(formObject)); // Log the formObject for debugging

  if (!formObject || !formObject.keyName || !formObject.signInOut || !formObject.person || !formObject.lockboxLocation) {
    Logger.log("Form data is missing or incomplete.");
    return "Form data is missing or incomplete."; // Return a message indicating missing or incomplete form data
  }

  Logger.log("Processing form...");

  var keyName = formObject.keyName;
  var signInOut = formObject.signInOut;
  var person = formObject.person;
  var lockboxLocation = formObject.lockboxLocation;

  Logger.log("Received form data - Key Name: " + keyName + ", Sign In/Out: " + signInOut + ", Person: " + person + ", Lockbox Location: " + lockboxLocation);

  // Check if the key is already assigned to a person or a lockbox
  var currentAssignment = isKeyAssigned(keyName);

  Logger.log("Current Assignment: " + JSON.stringify(currentAssignment));

  if (signInOut === "Signing Out") {
    Logger.log("Processing Sign Out...");
    // If the key is a restricted key and assigned to a lockbox, prompt for reassignment
    if (/^[A-F]\d{1,2}$/i.test(keyName) && currentAssignment && currentAssignment.assignedTo === "lockbox") {
      // Check if the provided location differs from the current assignment's location
      if (lockboxLocation !== currentAssignment.lockboxLocation) {
        var ui = SpreadsheetApp.getUi();
        var response = ui.alert("Warning", "This restricted key is currently assigned to a lockbox at " + currentAssignment.lockboxLocation + ". Are you sure you want to sign it out from " + lockboxLocation + "?", ui.ButtonSet.YES_NO);
        if (response == ui.Button.NO) {
          Logger.log("Form submission canceled.");
          return "Form submission canceled."; // Return a message indicating cancellation
        }
      }
      // Remove the key from the old lockbox
      removeFromDigitalLockbox(keyName);
    }

    // If the key is a restricted key and assigned to a person, prompt for confirmation
    if (/^[A-F]\d{1,2}$/i.test(keyName) && currentAssignment && currentAssignment.assignedTo === "person" && currentAssignment.person !== person) {
      var ui = SpreadsheetApp.getUi();
      var response = ui.alert("Warning", "This restricted key is currently assigned to " + currentAssignment.person + ". Are you sure you want to reassign it under " + person + "'s name?", ui.ButtonSet.YES_NO);
      if (response == ui.Button.NO) {
        Logger.log("Form submission canceled.");
        return "Form submission canceled."; // Return a message indicating cancellation
      }
    }

    // Update Log page with the information
    updateLog(keyName, signInOut, person, lockboxLocation);
    Logger.log("Log updated with sign-out information.");

    // Update FTL Key Assigned page
    updateSignedOutKeys(keyName, person);
    Logger.log("FTL Key Assigned page updated with sign-out information.");

    // Close the dialog box after processing
    Logger.log("Form submitted successfully.");
    return "Form submitted successfully!";

  } else if (signInOut === "Signing In") {
    Logger.log("Processing Sign In...");
    // If the key is a restricted key and assigned to a lockbox
    if (/^[A-F]\d{1,2}$/i.test(keyName) && currentAssignment && currentAssignment.assignedTo === "lockbox") {
        // Check if the provided location differs from the current assignment's location
        if (lockboxLocation !== currentAssignment.lockboxLocation) {
            var ui = SpreadsheetApp.getUi();
            var response = ui.alert("Warning", "This restricted key is currently assigned to a lockbox at " + currentAssignment.lockboxLocation + ". Do you want to move it to " + lockboxLocation + "?", ui.ButtonSet.YES_NO);
            if (response == ui.Button.NO) {
                Logger.log("Form submission canceled.");
                return "Form submission canceled."; // Return a message indicating cancellation
            }
            // Remove the key from the old lockbox
            removeFromDigitalLockbox(keyName);
        }
    }

    // If the key is a restricted key and assigned to a person, prompt for confirmation
    if (/^[A-F]\d{1,2}$/i.test(keyName) && currentAssignment && currentAssignment.assignedTo === "person" && currentAssignment.person !== person) {
      var ui = SpreadsheetApp.getUi();
      var response = ui.alert("Warning", "This restricted key is currently assigned to " + currentAssignment.person + ". Are you sure you want to sign it in under " + person + "'s name?", ui.ButtonSet.YES_NO);
      if (response == ui.Button.NO) {
        Logger.log("Form submission canceled.");
        return "Form submission canceled."; // Return a message indicating cancellation
      }
    }

    // Update Log page with the information
    updateLog(keyName, signInOut, person, lockboxLocation);
    Logger.log("Log updated with sign-in information.");

    // Update FTL Digital Lockbox page
    updateLockbox(keyName, lockboxLocation);
    Logger.log("FTL Digital Lockbox page updated.");

    // Remove the key assignment from the FTL.Keys.Assigned tab
    removeKeyAssignment(keyName);
    Logger.log("Key assignment removed from FTL Keys Assigned page.");

    // Close the dialog box after processing
    Logger.log("Form submitted successfully.");
    return "Form submitted successfully!";
  }
}



function processNonRestrictedKeyForm(formObject) {
  Logger.log("Non-Restricted Key Form Object: " + JSON.stringify(formObject)); // Log the formObject for debugging

  if (!formObject || !formObject.keyName || !formObject.signInOut || !formObject.person || !formObject.lockboxLocation) {
    Logger.log("Non-Restricted Key Form data is missing or incomplete.");
    return "Non-Restricted Key Form data is missing or incomplete."; // Return a message indicating missing or incomplete form data
  }

  Logger.log("Processing Non-Restricted Key Form...");

  var keyName = formObject.keyName;
  var signInOut = formObject.signInOut;
  var person = formObject.person;
  var lockboxLocation = formObject.lockboxLocation;

  // Implement logic for processing non-restricted keys here
  // Currently, this function doesn't have any specific logic implemented

  // Log the received form data
  Logger.log("Received form data - Key Name: " + keyName + ", Sign In/Out: " + signInOut + ", Person: " + person + ", Lockbox Location: " + lockboxLocation);

  // Return a success message
  return "Non-Restricted Key Form submitted successfully!";
}







function searchData() {
  var ui = SpreadsheetApp.getUi();
  var searchQuery = ui.prompt('Search', 'Enter a key or person\'s name:', ui.ButtonSet.OK_CANCEL);
  if (searchQuery.getSelectedButton() == ui.Button.OK) {
    var result = searchForKeyOrPerson(searchQuery.getResponseText());
    if (result && result.length > 0) {
      var message = 'Search Results:\n\n' + result.join('\n');
      ui.alert('Search Results', message, ui.ButtonSet.OK);
    } else {
      ui.alert('Search Results', 'No matching records found.', ui.ButtonSet.OK);
    }
  }
}

function searchForKeyOrPerson(query) {
  var logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Log");
  var data = logSheet.getDataRange().getValues();
  var results = [];
  
  // Search for matches in the log data
  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === query || data[i][3] === query) {
      var date = data[i][0].toLocaleDateString(); // Get the date in a readable format
      var action = data[i][2] === "Signing In" ? "Signed In" : "Signed Out";
      var keyName = data[i][1];
      var fromTo = data[i][2] === "Signing In" ? "From" : "To";
      var person = data[i][3];
      var lockbox = data[i][4];
      results.push([date, action, keyName, fromTo, person, lockbox]);
    }
  }
  
  // Format the results with spacing and lines between rows
  var formattedResults = [];
  results.forEach(function(row) {
    formattedResults.push(row.join(' | ')); // Join each piece of data with a pipe symbol
    formattedResults.push('-------------------------------------'); // Add a line between rows
  });

  return formattedResults;
}

function isKeyAssigned(keyName) {
  var assignedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FTL.Keys.Assigned");
  var lockboxSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FTL.Digital.Lockbox");
  var assignedData = assignedSheet.getDataRange().getValues();
  var lockboxData = lockboxSheet.getDataRange().getValues();
  
  Logger.log("Key Name: " + keyName); // Log the key name to check if it's received properly

  // Check if the key is assigned to a person
  for (var i = 1; i < assignedData.length; i++) {
    if (assignedData[i][0] === keyName && assignedData[i][1] !== "") {
      Logger.log("Key " + keyName + " is assigned to person: " + assignedData[i][1]);
      return { assignedTo: "person", person: assignedData[i][1] };
    }
  }
  
  // Check if the key is assigned to a lockbox
  for (var j = 1; j < lockboxData.length; j++) {
    if (lockboxData[j][0] === keyName) {
      Logger.log("Key " + keyName + " is assigned to lockbox at: " + lockboxData[j][1]);
      return { assignedTo: "lockbox", lockboxLocation: lockboxData[j][1] };
    }
  }
  
  // If the key is not assigned to anyone or any lockbox, return null
  Logger.log("Key " + keyName + " is not assigned to anyone.");
  return null;
}


function isRestrictedKey(keyName) {
  // Check if the key is a restricted key (A-F followed by 1-99)
  return /^[A-F]\d{1,2}$/i.test(keyName);
}

// This function will handle the assignment check for non-restricted keys
function isNonRestrictedKeyAssigned(keyName) {
   // You can implement the logic specific to non-restricted keys here
  // For now, we'll just return null as a placeholder
  return null;
}

function updateLog(keyName, signInOut, person, lockboxLocation) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Log");
  var timestamp = new Date();
  
  // Add the new row of information
  var newRow = [timestamp, keyName, signInOut, person, lockboxLocation];
  sheet.appendRow(newRow);
}

function updateSignedOutKeys(keyName, person) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FTL.Keys.Assigned");
  var data = sheet.getDataRange().getValues();
  var found = false;

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == keyName) {
      var currentPerson = data[i][1]; // Assuming the person's name is in column 2

      // If the key is already assigned to the same person, update the assignment
      sheet.getRange(i + 1, 2).setValue(person);
      found = true;
      break;
    }
  }

  // If the key is not found in the "FTL.Keys.Assigned" tab, add a new entry
  if (!found) {
    var lastRow = sheet.getLastRow() + 1;
    sheet.getRange(lastRow, 1).setValue(keyName);
    sheet.getRange(lastRow, 2).setValue(person); // Assuming the person's name is in column 2
  }

  // Sort the keys assignment data by key name
  sortKeysAssigned();
}

function removeKeyAssignment(keyName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FTL.Keys.Assigned");
  var range = sheet.getDataRange();
  var values = range.getValues();
  
  for (var i = 1; i < values.length; i++) {
    if (values[i][0] == keyName) {
      sheet.deleteRow(i + 1); // Delete the row where the key is assigned
      Logger.log("Key assignment removed: " + keyName); // Log the key removal
      break;
    }
  }
  
  // Sort the keys assignment data by key name
  sortKeysAssigned();
}

function updateLockbox(keyName, lockboxLocation) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FTL.Digital.Lockbox");
  var range = sheet.getDataRange();
  var values = range.getValues();
  var keyFound = false;

  for (var i = 1; i < values.length; i++) {
    if (values[i][0] == keyName) {
      sheet.getRange(i + 1, 2).setValue(lockboxLocation); // Update the lockbox location
      keyFound = true;
      break;
    }
  }

  // If the key is not found, add a new entry
  if (!keyFound) {
    var lastRow = sheet.getLastRow() + 1;
    sheet.getRange(lastRow, 1).setValue(keyName);
    sheet.getRange(lastRow, 2).setValue(lockboxLocation);
  }

  // Sort the keys assignment data by key name
  sortDigitalLockbox();

}

function removeFromDigitalLockbox(keyName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FTL.Digital.Lockbox");
  var range = sheet.getDataRange();
  var values = range.getValues();
  
  for (var i = 1; i < values.length; i++) {
    if (values[i][0] == keyName) {
      sheet.deleteRow(i + 1); // Delete the row where the key is located in the lockbox
      break;
    }
  }
  
  // Sort the digital lockbox data by key name
  sortDigitalLockbox();
}

function sortDigitalLockbox() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FTL.Digital.Lockbox");
  var range = sheet.getRange("A2:B" + sheet.getLastRow()); // Excluding header row
  var values = range.getValues();

  // Custom sorting function to sort the keys
  values.sort(function(row1, row2) {
    // Extract letter and numeric parts of key names
    var parts1 = extractParts(row1[0]);
    var parts2 = extractParts(row2[0]);

    // Compare letter parts first
    if (parts1.letter !== parts2.letter) {
      return parts1.letter.localeCompare(parts2.letter);
    }

    // If letter parts are the same, compare numeric parts
    return parts1.number - parts2.number;
  });

  // Set the sorted values back to the range
  range.setValues(values);
}

function sortKeysAssigned() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FTL.Keys.Assigned");
  var range = sheet.getRange("A2:B" + sheet.getLastRow()); // Excluding header row
  var values = range.getValues();

  // Custom sorting function to sort the keys
  values.sort(function(row1, row2) {
    // Extract letter and numeric parts of key names
    var parts1 = extractParts(row1[0]);
    var parts2 = extractParts(row2[0]);

    // Compare letter parts first
    if (parts1.letter !== parts2.letter) {
      return parts1.letter.localeCompare(parts2.letter);
    }

    // If letter parts are the same, compare numeric parts
    return parts1.number - parts2.number;
  });

  // Set the sorted values back to the range
  range.setValues(values);
}

// Function to extract letter and numeric parts of the key name
function extractParts(keyName) {
  Logger.log("Key name: " + keyName);
  // Extract letter and numeric parts using regex
  var match = keyName.match(/^([A-Za-z]+)(\d+)$/);
  if (match) {
    return {
      letter: match[1], // Letter part
      number: parseInt(match[2]) // Numeric part
    };
  }
  
  // If no match, return empty values
  return { letter: '', number: 0 };
}
