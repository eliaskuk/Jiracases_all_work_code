function calculatePercentageChange() {
  // Get the needed sheets
  var sheet = SpreadsheetApp.getActive().getSheetByName("Summary");
  var sheet_import = SpreadsheetApp.getActive().getSheetByName("Change_percent");

  var lastRow = sheet.getLastRow();

  // SID
  var lastValueSID = sheet.getRange("B" + lastRow).getValue();
  var secondToLastRowSID = lastRow - 1;
  var secondLastValueSID = sheet.getRange("B" + secondToLastRowSID).getValue();
  var change_percentSID = (lastValueSID - secondLastValueSID) / secondLastValueSID
  var change_percent_2dSID = change_percentSID.toFixed(3)

  // HU
  var lastValueHU = sheet.getRange("C" + lastRow).getValue();
  var secondToLastRowHU = lastRow - 1;
  var secondLastValueHU = sheet.getRange("C" + secondToLastRowHU).getValue();
  var change_percentHU = (lastValueHU - secondLastValueHU) / secondLastValueHU
  var change_percent_2dHU = change_percentHU.toFixed(3)

  // RMR
  var lastValueRMR = sheet.getRange("D" + lastRow).getValue();
  var secondToLastRowRMR = lastRow - 1;
  var secondLastValueRMR = sheet.getRange("D" + secondToLastRowRMR).getValue();
  var change_percentRMR = (lastValueRMR - secondLastValueRMR) / secondLastValueRMR
  var change_percent_2dRMR = change_percentRMR.toFixed(3)

  // Claim
  var lastValueCLAIM = sheet.getRange("E" + lastRow).getValue();
  var secondToLastRowCLAIM = lastRow - 1;
  var secondLastValueCLAIM = sheet.getRange("E" + secondToLastRowCLAIM).getValue();
  var change_percentCLAIM = (lastValueCLAIM - secondLastValueCLAIM) / secondLastValueCLAIM
  var change_percent_2dCLAIM = change_percentCLAIM.toFixed(3)

  // importing the percents to sheet
  sheet_import.getRange("A2").setValue(change_percent_2dSID);
  sheet_import.getRange("B2").setValue(change_percent_2dHU);
  sheet_import.getRange("C2").setValue(change_percent_2dRMR);
  sheet_import.getRange("D2").setValue(change_percent_2dCLAIM);
}