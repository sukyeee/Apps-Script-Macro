function RESET() {
  let sheetFrom = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ezadmin");
  sheetFrom.getRange("A2:AQ").clearContent();
//서식맞추기 추가

}