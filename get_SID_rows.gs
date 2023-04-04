// This function reads how many rows are in EDC_SID_OPEN first column and returns the number to code sheet column B first empty row
function getRowCountSID() {
  
var edcSheet = SpreadsheetApp.getActive().getSheetByName("EDC_SID_OPEN");
var edcLastRow = edcSheet.getLastRow();  
var codeSheet = SpreadsheetApp.getActive().getSheetByName("code_SID");
var codeLastRow = codeSheet.getLastRow() + 1;

// Initialize HU's
var hu_sum = 0; 
var range_hu = edcSheet.getRange("M:M"); // Get the range of column M
var values = range_hu.getValues(); // Get the values in column M

// Sum the HU's
for (var i = 0; i < values.length; i++) {
  if (!isNaN(values[i][0])) {
    hu_sum += Number(values[i][0]); // Add the value to the sum if it's a number
  }
}
  var listRow_SID = codeSheet.getLastRow() + 1; 
  var counter = codeSheet.getLastRow();

codeSheet.getRange(listRow_SID, 1).setValue(counter);
codeSheet.getRange(codeLastRow, 2).setValue(edcLastRow-1);
codeSheet.getRange(listRow_SID, 3).setValue(hu_sum);
}

function getRowCountRMR() {
  var edcSheet = SpreadsheetApp.getActive().getSheetByName("EDC_RMR_OPEN");
  var edcLastRow = edcSheet.getLastRow();
  
  var codeSheet = SpreadsheetApp.getActive().getSheetByName("code_RMR");
  var codeLastRow = codeSheet.getLastRow() + 1;
  
  codeSheet.getRange(codeLastRow, 1).setValue(edcLastRow-1);
}
function getRowCountCLAIM() {
  var edcSheet = SpreadsheetApp.getActive().getSheetByName("EDC_CLAIM_OPEN");
  var edcLastRow = edcSheet.getLastRow();
  
  var codeSheet = SpreadsheetApp.getActive().getSheetByName("code_CLAIM");
  var codeLastRow = codeSheet.getLastRow() + 1;
  
  codeSheet.getRange(codeLastRow, 1).setValue(edcLastRow-1);
}




