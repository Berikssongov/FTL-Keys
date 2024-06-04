function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🔑')
    .addItem('Process Key Form ', 'openFormDialog')
    .addItem('Search Keys', 'openSearchDialog')
    .addToUi();
  ui.createMenu('🔒')
    .addItem('!!! Compile Restricted Keys Report !!!', 'compileReport')
    .addToUi();
    
}

function openFormDialog() {
  var html = HtmlService.createHtmlOutputFromFile('Form')
      .setWidth(400)
      .setHeight(625);
  SpreadsheetApp.getUi().showModalDialog(html, 'Enter Key Information');
}

function fakeMenuItem() {
  // This function can be left blank since it's a fake menu item
  // or you can add any desired functionality here
}

function processForm(formObject) {
  Logger.log("Form Object: " + JSON.stringify(formObject)); // Log the form object to check if it's received properly

  if (!formObject || !formObject.keyName || !formObject.signInOut || !formObject.person || !formObject.lockboxLocation) {
    logToSheet('Form data is missing.');
    Logger.log("Form data is missing or incomplete.");
    return "Form data is missing or incomplete."; // Return a message indicating missing or incomplete form data
  }

  // Check if the form data is complete
  if (!formObject.keyName || !formObject.signInOut || !formObject.person || !formObject.lockboxLocation) {
    logToSheet('Form data is incomplete.');
    return "Form data is incomplete."; // Return a message indicating incomplete form data
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
    // If the key is non-restricted, call the processNonRestrictedKeyForm function from NonRestrictedKeyHandler
    Logger.log("Processing Non-Restricted Key Form");
    var result = processNonRestrictedKeyForm(formObject);
    Logger.log("Non-Restricted Key Form Result: " + result);
    return result;
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

function openSearchDialog() {
  var html = HtmlService.createHtmlOutputFromFile('searchDialog')
      .setWidth(500)
      .setHeight(400);
  SpreadsheetApp.getUi().showModalDialog(html, 'Search Keys and People');
}

function searchForKeyOrPerson(query) {
  logToSheet('searchForKeyOrPerson called with query: ' + query);
  
  var logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Log");
  var data = logSheet.getDataRange().getValues();
  var keyHistory = [];

  for (var i = 1; i < data.length; i++) {
    var keyName = data[i][1] ? data[i][1] : "";
    var personName = data[i][3] ? data[i][3] : "";
    var date = new Date(data[i][0]).toLocaleDateString();
    var action = data[i][2] === "Signing In" ? "Signed In" : "Signed Out";
    var fromTo = data[i][2] === "Signing In" ? "From" : "To";
    var lockbox = data[i][4];
    var logEntry = `${date} | ${action} | ${keyName} | ${fromTo} | ${personName} | ${lockbox}`;

    if (keyName === query || personName === query) {
      keyHistory.push(logEntry);
    }
  }

  var currentKeys = getCurrentKeys(query);
  var currentAssignee = getCurrentAssignee(query);

  logToSheet('search results: ' + JSON.stringify({
    keyHistory: keyHistory,
    currentKeys: currentKeys,
    currentAssignee: currentAssignee
  }));
  
  return {
    keyHistory: keyHistory,
    currentKeys: currentKeys,
    currentAssignee: currentAssignee
  };
}

function getCurrentKeys(person) {
  logToSheet('getCurrentKeys called with person: ' + person);

  var assignedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FTL.Keys.Assigned");
  
  if (!assignedSheet) {
    logToSheet('Error: FTL.Keys.Assigned sheet not found');
    return ['Error: FTL.Keys.Assigned sheet not found'];
  }
  
  var data = assignedSheet.getDataRange().getValues();
  var currentKeys = [];

  for (var i = 1; i < data.length; i++) {
    if (data[i][1] === person) {
      var keyName = data[i][0] ? data[i][0] : "";
      var lockbox = data[i][2] ? data[i][2] : "";
      currentKeys.push(`${keyName} (${lockbox})`);
    }
  }

  logToSheet('getCurrentKeys results: ' + JSON.stringify(currentKeys));

  return currentKeys;
}

function getCurrentAssignee(key) {
  logToSheet('getCurrentAssignee called with key: ' + key);

  var assignedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FTL.Keys.Assigned");
  
  if (!assignedSheet) {
    logToSheet('Error: FTL.Keys.Assigned sheet not found');
    return 'Error: FTL.Keys.Assigned sheet not found';
  }
  
  var data = assignedSheet.getDataRange().getValues();
  var assignee = 'Not Assigned';

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] === key) {
      assignee = data[i][1] ? data[i][1] : 'Not Assigned';
      break;
    }
  }

  logToSheet('getCurrentAssignee result: ' + assignee);

  return assignee;
}

function processSearch(query) {
  logToSheet('Server: Performing search with query: ' + query);

  var logSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Log");
  var assignedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FTL.Keys.Assigned");
  var lockboxSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FTL.Digital.Lockbox");

  var logData = logSheet.getDataRange().getValues();
  var assignedData = assignedSheet.getDataRange().getValues();
  var lockboxData = lockboxSheet.getDataRange().getValues();

  var keyHistory = [];
  var currentAssignee = [];
  var currentKeys = [];
  var lockboxKeys = {};

  // Search in the log for key history
  for (var i = 1; i < logData.length; i++) {
    if (logData[i][1] === query || logData[i][3] === query) {
      var date = logData[i][0].toLocaleDateString();
      var action = logData[i][2] === "Signing In" ? "Signed In" : "Signed Out";
      var keyName = logData[i][1];
      var fromTo = logData[i][2] === "Signing In" ? "From" : "To";
      var person = logData[i][3];
      var lockbox = logData[i][4];
      keyHistory.push(`${date} | ${action} | ${keyName} | ${fromTo} | ${person} | ${lockbox}`);
    }
  }

  // Search in the assigned sheet for current keys assigned to a person or the key itself
  for (var i = 1; i < assignedData.length; i++) {
    if (assignedData[i][1] === query) {
      currentKeys.push(assignedData[i][0]);
    }
    if (assignedData[i][0] === query) {
      currentAssignee.push(assignedData[i][1]);
    }
  }

  // Search in the lockbox sheet for keys in lockboxes
  for (var i = 1; i < lockboxData.length; i++) {
    if (lockboxData[i][0] === query) {
      if (!lockboxKeys[lockboxData[i][1]]) {
        lockboxKeys[lockboxData[i][1]] = 0;
      }
      lockboxKeys[lockboxData[i][1]]++;
    }
  }

  // Convert lockboxKeys object to array format
  var lockboxKeysArray = Object.keys(lockboxKeys).map(key => ({
    name: key,
    count: lockboxKeys[key]
  }));

  return {
    keyHistory: keyHistory,
    currentAssignee: Array.from(new Set(currentAssignee)),
    currentKeys: currentKeys,
    lockboxKeys: lockboxKeysArray
  };
}


// Function to log to a specific sheet in the spreadsheet
function logToSheet(logText) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Debug");
  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1).setValue(logText);
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
