/** @OnlyCurrentDoc */
/** Global variables and constants */
// Variables
var spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
var lock = LockService.getScriptLock();
var ui = SpreadsheetApp.getUi();

// Consolidate sheet references into an object
var sheets = {
  expenses: spreadsheet.getSheetByName("Expenses"),
  incomes: spreadsheet.getSheetByName("Incomes"),
  transactions: spreadsheet.getSheetByName("Transactions"),
  report: spreadsheet.getSheetByName("Report"),
  loan: spreadsheet.getSheetByName("Loan"),
  investment: spreadsheet.getSheetByName("Investment"),
  prediction: spreadsheet.getSheetByName("Prediction"),
  statement: spreadsheet.getSheetByName("Statement"),
  roi: spreadsheet.getSheetByName("Record of Incomes"),
  roe: spreadsheet.getSheetByName("Record of Expenses"),
  rot: spreadsheet.getSheetByName("Record of Transactions"),
  test: spreadsheet.getSheetByName("test")
};

// Constants
var CONFIRMATION_DIALOG_TITLE = "⚠️ Warning";
var CONFIRMATION_DIALOG_MESSAGE = "Something is wrong. Do you want to continue?";
var RESET_SUCCESSFUL_TITLE = "🔄 Reset Successful";
var ERROR_TITLE = "🛑 Error!!";

// Column for prediction and statement sheet
var STATEMENT_COLUMN_NUM = 14

/** End of global variables and constants */

/** Utility Functions */

// Function to display a confirmation dialog
function showConfirmationDialog() {
  var result = ui.alert(CONFIRMATION_DIALOG_TITLE, CONFIRMATION_DIALOG_MESSAGE, ui.ButtonSet.YES_NO);
  return result == ui.Button.YES;
}

// Function to check if another instance of the script is already running
function isScriptAlreadyRunning() {
  return !lock.tryLock(0);
}

// Function to display a message
function showMessage(message, title) {
  ui.alert(title || "Alert", message, ui.ButtonSet.OK);
}

// Function to get sheet by name
function getSheet(name) {
  var sheet = sheets[name];

  if (!sheet) {
    throw new Error("Sheet '" + name + "' not found.");
  }

  return sheet;
}

/**
 * Function to get data from a target sheet and range.
 * @param {string} sheetName - Name of the target sheet.
 * @param {string | string[]} ranges - Range(s) to retrieve data from.
 * @returns {any[][]} - Data fetched from the target sheet and range.
 */
function getData(sheetName, ranges) {
  var sheet = getSheet(sheetName);

  if (!Array.isArray(ranges)) {
    ranges = [ranges]; // Convert single range to array for consistency
  }

  // Fetch data from each range and create an array of arrays
  var data = ranges.map(function(range) {
    var rangeValues = sheet.getRange(range).getValues();
    return rangeValues.flat(); // Flatten each range's values into a single array
  });

  return data;
}

/**
 * Function to set data in a target sheet and range.
 * @param {string} targetSheetName - Name of the target sheet.
 * @param {string[][] | any[][]} data - Data to be set in the target sheet.
 * @param {number} startRow - Starting row index in the target sheet.
 * @param {number} startColumn - Starting column index in the target sheet.
 */
function setData(targetSheetName, data, startRow, startColumn) {
  var targetSheet = getSheet(targetSheetName);

  // Determine the dimensions of the data
  var numRows = data.length;
  var numColumns = data[0].length;

  // Get the target range
  var targetRange = targetSheet.getRange(startRow, startColumn, numRows, numColumns);

  // Set values to the target range
  targetRange.setValues(data);
}

/**
 * Function to populate data from a sheet or multiple sheets to respective target sheets.
 * @param {string | string[]} sourceSheetNames - Name of the source sheet(s).
 * @param {string | string[][]} sourceRanges - Data ranges to retrieve from the source sheet(s).
 * @param {string | string[]} targetSheetsNames - Name of the target sheet(s).
 * @param {string | string[][]} targetRanges - Data ranges to set data to the target sheet(s).
 */
function saveData(sourceSheetNames, sourceRanges, targetSheetNames, targetRanges) {
  // Ensure that the number of source sheets and source ranges match
  if (sourceSheetNames.length !== sourceRanges.length || sourceSheetNames.length !== targetSheetNames.length) {
    throw new Error("Number of source sheet names does not match number of source ranges and target sheet names.");
  }

  if (targetRanges) {
    // Ensure that the number of target sheets and target ranges match
    if (targetSheetNames.length !== targetRanges.length) {
      throw new Error("Number of target sheet names does not match number of target ranges.");
    }
  }

  // Iterate through each source and target pair
  for (var i = 0; i < sourceSheetNames.length; i++) {
    var sourceSheetName = sourceSheetNames[i];
    var sourceRange = sourceRanges[i];
    var targetSheetName = targetSheetNames[i];

    // Get data from the source sheet
    var data = getData(sourceSheetName, sourceRange);
    Logger.log(data);

    // If target range is not specified, default to new row in target sheet
    if (!targetRange) {
      var lastRow = getSheet(targetSheetName).getLastRow();
      var startRow = lastRow + 1;
      var startColumn = 1; // Default to 1st column
      setData(targetSheetName, data, startRow, startColumn);
    } else {
      var targetRange = targetRanges[i];
      // Extract target range details
      var startRow = targetRange[0];
      var startColumn = targetRange[1];
      var numRows = data.length;
      var numColumns = data[0].length;

      // Set the data to the target sheet with the specified range
      setData(targetSheetName, data, startRow, startColumn, numRows, numColumns);
    }
  }
}

/**
 * Function to clear ranges, range lists, or individual cells on a sheet.
 * @param {Sheet} sheet - The sheet from which the ranges will be cleared.
 * @param {string[] | string[][]} rangesOrRangeLists - An array of range strings or range lists to clear.
 */
function clearRanges(sheetName, rangesOrRangeLists) {
  var sheet = getSheet(sheetName);

  if (!Array.isArray(rangesOrRangeLists)) {
    rangesOrRangeLists = [rangesOrRangeLists]; // Convert single range to array for consistency
  }

  rangesOrRangeLists.forEach(function(rangeOrRangeList) {
    if (Array.isArray(rangeOrRangeList)) {
      // If it's a range list, clear each range in the list
      rangeOrRangeList.forEach(function(range) {
        sheet.getRange(range).clear({ contentsOnly: true, skipFilteredRows: true });
      });
    } else {
      // If it's a single range, clear it
      sheet.getRange(rangeOrRangeList).clear({ contentsOnly: true, skipFilteredRows: true });
    }
  });
}

// Function to set formula in a specific range of a sheet
function setFormulaInRangeOfSheet(sheetName, range, formula) {
  var sheet = getSheet(sheetName);

  var cell = sheet.getRange(range);
  if (!cell) {
    throw new Error("Invalid range '" + range + "' in sheet '" + sheetName + "'.");
  }

  cell.setFormula(formula);
}

// Function to check of the specified range it blank or not
// Function to check if the specified range is blank or not
function checkBlankRange(sheetName, columnRangesToCheck, rowNo) {
  var sheet = getSheet(sheetName);

  // Fetch values of all specified ranges at once
  var values = sheet.getRangeList(columnRangesToCheck.map(range => range + rowNo)).getValues().flat();

  // Check if any blank values are found
  return !values.some(value => value !== "");
}

// Function to calculate duration between dates in month for loan
function calculateMonthsBetweenDates(startDate, endDate) {
  // Check if input dates are valid JavaScript Date objects
  if (!(startDate instanceof Date) || !(endDate instanceof Date) || isNaN(startDate) || isNaN(endDate)) {
    throw new Error("Invalid date input.");
  }

  var start = new Date(startDate);
  var end = new Date(endDate);
  var years = end.getFullYear() - start.getFullYear();
  var months = end.getMonth() - start.getMonth();
  var totalMonths = years * 12 + months;
  return totalMonths;
}

/** End of Utility Functions */

/** Main Functions */
// Function to clear monthly records
function clearMonthlyRecords() {
  var sheetsToClear = ["expenses", "incomes", "transactions"];
  var rangeToCLear = 'A2:D100';

  sheetsToClear.forEach(function(sheetName){
    clearRanges(sheetName, rangeToCLear);
  });
}

// Function to update report sheet
function updateReport() {
  // Get report data from prediction sheet
  var reportData = getData("prediction", "A8:N8");

  // Get total expenses and total incomes
  var totalIncomes = getSheet("report").getRange("B12").getValue();
  var totalExpenses = getSheet("report").getRange("B23").getValue();

  // Prepare data for setting values in report sheet
  var expenseIncomeData = [[totalIncomes, totalExpenses]];

  // Get the last row index in the report sheet
  var lastRow = getSheet("statement").getLastRow();

  // Set values in the statement sheet
  setData("statement", reportData, lastRow + 1, 1);
  setData("statement", expenseIncomeData, lastRow + 1, 17)
}

// Function to update the last month balance in Prediction sheet
function updateLastMonthBalance() {
  var data = getData("prediction", "A8:N8");
  setData("prediction", data, 3, 1);
  clearRanges("prediction", "A8:N8");
}

// Function to reset data monthly
function monthlyReset() {
  // Check if another instance of the script is already running
  if (isScriptAlreadyRunning()) {
    showMessage("The script is already running. Please wait!", ERROR_TITLE);
    return;
  }

  // Lock the script
  lock.waitLock(10000); // wait 10 seconds for others' use of the code section and lock to stop and then proceed

  try {
    // Save net total data from report to respective records
    saveData(
      ["report", "report", "report"],
      ["A12:Z12", "A23:Z23", "A34:Z34"],
      ["roi", "roe", "rot"]
    );

    // Call the update report function
    updateReport();

    // Call the clear functions
    clearMonthlyRecords();

    // update balance for the last month in the prediction sheet for new month prediction
    updateLastMonthBalance();

    showMessage("Reset successful! Data has been cleared and updated.", RESET_SUCCESSFUL_TITLE);
  } catch (e) {
    // Handle any errors that may occur during reset
    showMessage("An error occurred: " + e.message, ERROR_TITLE);
  } finally {
    // Release the lock
    lock.releaseLock();
  }
}

// // Trigger to calculate loan duration when status is set to clear (not practical lol but it was a good try)
// function onEdit(e) {
//   var sheet = e.source.getActiveSheet();
//   var editedCell = e.range;
//   var loanSheetName = "Loan";
//   var columnC = 3; // Column C
//   var clearStatus = "Clear";

//   // Check if the edited cell is in the "Loan" sheet, in column C, and the new value is "Clear"
//   if (sheet.getName() === loanSheetName && editedCell.getColumn() === columnC && editedCell.getValue() === clearStatus) {
//     calculateLoanDuration();
//   }
// }

// Reset Functionality for mobile
// function onEdit(e) {
//   var range = e.range;
//   var sheet = e.source.getActiveSheet();
//   var sheetName = sheet.getName();
//   var column = range.getColumn();
//   var row = range.getRow();

//   // Check if the edited cell is in the "Report" sheet and matches the desired cell
//   if (sheetName === "Report" && column === 10 && row === 20) {
//     reset();
//   }
// }

// Test function
function test() {
  // Implement as per your logic...
    clearMonthlyRecords();
}

// Custom menu
function createMenu() {
  SpreadsheetApp.getUi().createMenu("⚙️ Custom Menu")
    .addItem("Monthly Reset", "monthlyReset")
    .addItem("Test", "test")
    .addToUi();
}

function onOpen() {
  createMenu();
}
