// Configuration: Specify the maximum number of each type of restricted key
var maxKeys = {
  "A": 20,
  "B": 20,
  "C": 31,
  "D": 10,
  "E": 15,
  "F": 15
};

// Function to create or clear the report sheet
function createOrClearReportSheet(sheetName) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(sheetName);
  
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    logToSheet("Created new sheet: " + sheetName);
  } else {
    sheet.clear();
    logToSheet("Cleared existing sheet: " + sheetName);
  }
  
  return sheet;
}

// Function to get data of assigned keys
function getAssignedKeysData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FTL.Keys.Assigned");
  var data = sheet.getRange("A2:B" + sheet.getLastRow()).getValues();
  logToSheet("Retrieved assigned keys data.");
  return data;
}

// Function to get data of keys in the lockbox
function getLockboxKeysData() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FTL.Digital.Lockbox");
  var data = sheet.getRange("A2:B" + sheet.getLastRow()).getValues();
  logToSheet("Retrieved lockbox keys data.");
  return data;
}

function compileReport() {
  var ui = SpreadsheetApp.getUi();
  var passwordPrompt = ui.prompt('WARNING!!', 'Running This Report Will Override The Last Report', ui.ButtonSet.OK_CANCEL);

  if (passwordPrompt.getSelectedButton() == ui.Button.OK) {
    var enteredPassword = passwordPrompt.getResponseText();
    if (checkPassword(enteredPassword)) {
      var reportSheetName = "Restricted";
      var reportSheet = createOrClearReportSheet(reportSheetName);

      var assignedKeysData = getAssignedKeysData();
      var lockboxKeysData = getLockboxKeysData();
      
      // Filter for restricted keys (e.g., keys starting with 'A' to 'F')
      var restrictedAssignedKeys = assignedKeysData.filter(row => isRestrictedKey(row[0]));
      var restrictedLockboxKeys = lockboxKeysData.filter(row => isRestrictedKey(row[0]));

      logToSheet("Filtered restricted assigned keys.");
      logToSheet("Filtered restricted lockbox keys.");

      // Set up headers
      var headers = ["", "A", "B", "C", "D", "E", "F"];
      reportSheet.appendRow(headers);
      
      // Populate the grid starting from the second row
      var startingRow = 2; // Header row is 1
      var maxRowNumber = getMaxRowNumber(restrictedAssignedKeys);
      var rowNumbers = [];
      for (var i = 1; i <= maxRowNumber; i++) {
      rowNumbers.push([i]);
}
reportSheet.getRange(startingRow, 1, maxRowNumber, 1).setValues(rowNumbers);

      restrictedAssignedKeys.forEach(row => {
        var key = row[0];
        var person = row[1];
        var parts = extractParts(key);
        var letterIndex = headers.indexOf(parts.letter);
        if (letterIndex !== -1) {
          reportSheet.getRange(parts.number + startingRow - 1, letterIndex + 1).setValue(person);
        }
      });

      restrictedLockboxKeys.forEach(row => {
        var key = row[0];
        var location = row[1];
        var parts = extractParts(key);
        var letterIndex = headers.indexOf(parts.letter);
        if (letterIndex !== -1) {
          reportSheet.getRange(parts.number + startingRow - 1, letterIndex + 1).setValue(location);
        }
      });

      logToSheet("Compiled report data and updated the 'Restricted' sheet.");

      // Auto resize columns for better readability
      reportSheet.autoResizeColumns(1, headers.length);
      
      // Enable text wrapping for all cells
      var range = reportSheet.getDataRange();
      range.setWrap(true);

      // Remove unused columns
      for (var i = 7; i > headers.length; i--) {
        reportSheet.deleteColumn(i);
      }

      // Remove blank rows
      var blankRows = getMaxRowNumber(restrictedAssignedKeys) - restrictedAssignedKeys.length - restrictedLockboxKeys.length;
      if (blankRows > 0) {
        reportSheet.deleteRows(restrictedAssignedKeys.length + startingRow - 1, blankRows);
      }

      // Optimize: Set "?" for blank cells
      var rangeToProcess = reportSheet.getRange(startingRow, 1, getMaxRowNumber(restrictedAssignedKeys), headers.length);
      var values = rangeToProcess.getValues();
      for (var row = 0; row < values.length; row++) {
        for (var col = 0; col < values[row].length; col++) {
          if (!values[row][col]) {
            values[row][col] = "?";
          }
        }
      }
      rangeToProcess.setValues(values);

    // gray out cells exceeding the maximum key count
      for (var key in maxKeys) {
        var maxCount = maxKeys[key];
        var letterIndex = headers.indexOf(key);
        if (letterIndex !== -1) {
          for (var i = maxCount + startingRow; i <= getMaxRowNumber(restrictedAssignedKeys) + startingRow; i++) {
           var cell = reportSheet.getRange(i, letterIndex + 1);
           cell.setBackground("gray");
           cell.clearContent();
          }
        }
      }


      SpreadsheetApp.getActiveSpreadsheet().toast('The restricted keys report has been successfully compiled.', 'Compilation Successful');
    } else {
      ui.alert('Incorrect Password', 'The password you entered is incorrect.', ui.ButtonSet.OK);
      logToSheet("Please do not use this if you do not know it's use");
    }
  } else {
    logToSheet("Compilation cancelled by user.");
  }
}



// Function to check the entered password
function checkPassword(password) {
  // Replace this with your actual password or implement your own logic for password validation
  var actualPassword = "1585";
  return password === actualPassword;
}

// Function to get the maximum row number based on the highest numbered restricted key
function getMaxRowNumber(keysData) {
  var maxRowNumber = 0;
  keysData.forEach(row => {
    var parts = extractParts(row[0]);
    maxRowNumber = Math.max(maxRowNumber, parts.number);
  });
  return maxRowNumber;
}
