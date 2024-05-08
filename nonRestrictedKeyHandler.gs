// Function to process the non-restricted key form
function processNonRestrictedKeyForm(formObject) {
  // Step 1: Check if the formObject is defined
  if (!formObject) {
    logToSheet('Form data is missing.');
    return "Form data is missing."; // Return a message indicating missing form data
  }

  // Step 2: Log the form object
  logToSheet("Non-Restricted Key Form Object: " + JSON.stringify(formObject));

  // Step 3: Check if the form data is complete
  if (!formObject.keyName || !formObject.signInOut || !formObject.person || !formObject.lockboxLocation) {
    logToSheet('Form data is incomplete.');
    return "Form data is incomplete."; // Return a message indicating incomplete form data
  }

  // Step 4: Log the form information
  var message = "Key Name: " + formObject.keyName + ", " +
                "Sign In/Out: " + formObject.signInOut + ", " +
                "Person: " + formObject.person + ", " +
                "Lockbox Location: " + formObject.lockboxLocation;
  logToSheet('Form Information: ' + message);

  // Step 5: Log the key and location being searched for
  logToSheet("Searching for Key: " + formObject.keyName + " at Location: " + formObject.lockboxLocation);

  // Step 6: Process sign-out or sign-in based on the value of formObject.signInOut
  if (formObject.signInOut === "Signing Out") {
    // Step 6a: Assign the key to the specified person
    var assignmentResult = assignKey(formObject.keyName, formObject.person, formObject.lockboxLocation);
    logToSheet("Assignment Result: " + JSON.stringify(assignmentResult));

    // Step 6b: Check if the assignment was successful
    if (assignmentResult.success) {
      // Step 6c: Update Log page with the information
      updateLog(formObject.keyName, formObject.signInOut, formObject.person, formObject.lockboxLocation);
      logToSheet("Log updated with sign-out information.");
    } else {
      // Step 6d: Show an alert with the error message
      ui.alert(assignmentResult.message);
    }
 } else if (formObject.signInOut === "Signing In") {
    // Step 7: Check if the person has the key assigned
    var assignedSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FTL.Keys.Assigned");
    var assignedRange = assignedSheet.getDataRange();
    var assignedValues = assignedRange.getValues();

    var keyAssigned = false;
    for (var i = 1; i < assignedValues.length; i++) {
        if (assignedValues[i][0] === formObject.keyName && assignedValues[i][1] === formObject.person) {
            keyAssigned = true;
            break;
        }
    }

    if (!keyAssigned) {
        // Key not assigned to the specified person
        logToSheet("The key " + formObject.keyName + " is not assigned to " + formObject.person + ".");
        ui.alert("Warning", "The key " + formObject.keyName + " is not assigned to " + formObject.person + ".", ui.ButtonSet.OK);
        return "The key is not assigned to the specified person."; // Return an error message
    }

    // Step 8: Remove the key assignment
    var removed = false;
    for (var i = assignedValues.length - 1; i >= 1; i--) {
        if (assignedValues[i][0] === formObject.keyName && assignedValues[i][1] === formObject.person) {
            assignedSheet.deleteRow(i + 1); // Adding 1 because row numbering starts from 1, not 0
            removed = true;
            break;
        }
    }

    if (!removed) {
        // Error occurred while removing the key assignment
        logToSheet("Error occurred while removing the key assignment.");
        return "Error occurred while removing the key assignment."; // Return an error message
    }

    // Step 9: Add the key to the specified lockbox location
    var lockboxSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FTL.Digital.Lockbox");
    var lastRow = lockboxSheet.getLastRow() + 1;
    lockboxSheet.getRange(lastRow, 1).setValue(formObject.keyName);
    lockboxSheet.getRange(lastRow, 2).setValue(formObject.lockboxLocation); // Assuming the lockbox location is in column 2

    // Step 10: Update Log page with sign-in information
    updateLog(formObject.keyName, formObject.signInOut, formObject.person, formObject.lockboxLocation);
    logToSheet("Log updated with sign-in information.");

    // Sort all keys (restricted and non-restricted)
    sortAllKeys();

    // Step 11: Implement error handling and logging (if necessary)
 }
}

// Function to assign a key to a person and remove it from the lockbox if available
function assignKey(keyName, person, lockboxLocation) {
    // Log the type and value of keyName
    logToSheet("Type of keyName: " + typeof keyName);
    logToSheet("Value of keyName: " + keyName);

    // Ensure keyName is a string
    if (typeof keyName !== 'string') {
        return { success: false, message: "Invalid key name format." };
    }
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
    logToSheet("Key " + keyName + " successfully assigned to " + person + ".");

  // Sort all keys (restricted and non-restricted)
  sortAllKeys();

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

  // Return a success message
  return { success: true, message: "Log updated successfully!" };
}

// Function to log to a specific sheet in the spreadsheet
function logToSheet(logText) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Debug");
  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow + 1, 1).setValue(logText);
}

// Function to sort all keys (restricted and non-restricted)
function sortAllKeys() {
  try {
    sortKeys("FTL.Digital.Lockbox");
    sortKeys("FTL.Keys.Assigned");
    
    // Log success message
    logToSheet("Sorting keys completed successfully.");
  } catch (error) {
    // Log any errors that occur during sorting
    logToSheet("Error occurred during sorting: " + error);
  }
}

// Function to sort keys for a specific sheet
function sortKeys(sheetName) {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(sheetName);
  var range = sheet.getRange("A2:B" + sheet.getLastRow()); // Excluding header row
  var values = range.getValues();
  
  // Custom sorting function to sort the keys
  values.sort(function(row1, row2) {
    // Extract letter and numeric parts of key names
    var parts1 = extractParts(row1[0]);
    var parts2 = extractParts(row2[0]);

    // Check if either key is restricted
    var isRestricted1 = isRestricted(parts1.letter);
    var isRestricted2 = isRestricted(parts2.letter);

    // If one key is restricted and the other is not, prioritize the restricted key
    if (isRestricted1 && !isRestricted2) {
      return -1; // Move row1 (restricted key) up
    } else if (!isRestricted1 && isRestricted2) {
      return 1; // Move row2 (restricted key) up
    }

    // If both keys are either restricted or non-restricted, sort them based on letter and number
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

// Function to check if a key is restricted (starts with A-F followed by 1-99)
function isRestricted(letter) {
  return /^[A-F]\d{1,2}$/i.test(letter);
}

// Function to check if a key is restricted
function isRestricted(keyName) {
  return /^[A-F]\d{1,2}$/i.test(keyName);
}

// Function to extract letter and numeric parts of the key name
function extractParts(keyName) {
  var match = keyName.toString().match(/^([A-Za-z]*)(\d+)$/);
  if (match) {
    return {
      letter: match[1].toUpperCase(), // Convert to uppercase
      number: parseInt(match[2])
    };
  }
  // If no match, return empty values
  return { letter: '', number: parseInt(keyName) };
}




