function processNonRestrictedKeyForm(formObject) {
  // Check if the formObject is defined
  if (!formObject) {
    logToSheet('Form data is missing.');
    return "Form data is missing."; // Return a message indicating missing form data
  }

  // Log the form object
  logToSheet("Non-Restricted Key Form Object: " + JSON.stringify(formObject));

  // Check if the form data is complete
  if (!formObject.keyName || !formObject.signInOut || !formObject.person || !formObject.lockboxLocation) {
    logToSheet('Form data is incomplete.');
    return "Form data is incomplete."; // Return a message indicating incomplete form data
  }

  // Log the form information
  var message = "Key Name: " + formObject.keyName + ", " +
                "Sign In/Out: " + formObject.signInOut + ", " +
                "Person: " + formObject.person + ", " +
                "Lockbox Location: " + formObject.lockboxLocation;
  logToSheet('Form Information: ' + message);

  // Log the key and location being searched for
  logToSheet("Searching for Key: " + formObject.keyName + " at Location: " + formObject.lockboxLocation);

  // Process sign-out
  if (formObject.signInOut === "Signing Out") {
    // Assign the key to the specified person
    var assignmentResult = assignKey(formObject.keyName, formObject.person, formObject.lockboxLocation);
    logToSheet("Assignment Result: " + JSON.stringify(assignmentResult));

    // Check if the assignment was successful
    if (assignmentResult.success) {
      // Update Log page with the information
      updateLog(formObject.keyName, formObject.signInOut, formObject.person, formObject.lockboxLocation);
      logToSheet("Log updated with sign-out information.");
    } else {
      // Show an alert with the error message
      ui.alert(assignmentResult.message);
    }
  }

  // Rest of the function logic...
}


// Function to assign a key to a person and remove it from the lockbox if available
function assignKey(keyName, person, lockboxLocation) {
    var lockboxSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FTL.Digital.Lockbox");
    var lockboxData = lockboxSheet.getDataRange().getValues();
    var keyFound = false;

    // Log the lockbox data to check if it's retrieved properly
    logToSheet("Lockbox Data: " + JSON.stringify(lockboxData));

    // Log the key name and lockbox location being searched
    logToSheet("Searching for Key: " + keyName + " at Location: " + lockboxLocation);

    // Check if the key is available in the specified lockbox location
    for (var i = 1; i < lockboxData.length; i++) {
        var currentKeyName = lockboxData[i][0].toString();
        var currentLocation = lockboxData[i][1].toString();
        logToSheet("Comparing Key: " + currentKeyName + " at Location: " + currentLocation);
        if (currentKeyName === keyName && currentLocation === lockboxLocation) {
            // If the key is found in the lockbox at the specified location, remove it
            lockboxSheet.deleteRow(i + 1);
            keyFound = true;
            logToSheet("Key " + keyName + " removed from lockbox at " + lockboxLocation + ".");
            break;
        }
    }

    if (!keyFound) {
        // If the key is not found in the specified lockbox location, return an error message
        logToSheet("Key " + keyName + " not found at " + lockboxLocation + ".");
        return { success: false, message: "The requested key is not available in the specified lockbox location." };
    }

    // Add the key assignment to the FTL.Keys.Assigned sheet
    var assignedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FTL.Keys.Assigned");
    var lastRow = assignedSheet.getLastRow() + 1;
    assignedSheet.getRange(lastRow, 1).setValue(keyName);
    assignedSheet.getRange(lastRow, 2).setValue(person); // Assuming the person's name is in column 2

    // Sort the keys assignment data by key name
    sortKeysAssigned();

    // Return a success message
    return { success: true, message: "Key assigned successfully!" };
}



// Function to update the log page
function updateLog(keyName, signInOut, person, lockboxLocation) {
  // Log the update operation
  logToSheet("Log Updated: Key " + keyName + " " + signInOut + " by " + person + " at " + lockboxLocation);

  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Log");
  var timestamp = new Date();

  // Add the new row of information
  var newRow = [timestamp, keyName, signInOut, person, lockboxLocation];
  sheet.appendRow(newRow);
}


// Function to log to a specific sheet in the spreadsheet
function logToSheet(logText) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Debug");
  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1).setValue(logText);
}
