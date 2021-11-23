function Paste_to_FINAL() {
  console.log("Paste_to_FINAL");

  let sheetFrom = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Sheet1");
  let sheetTo = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FINAL")
  let lastRow = sheetFrom.getLastRow(); //마지막 데이터가있는 행
  let lastRow_sheetTo = sheetTo.getMaxRows(); //마지막 최대 행 수

  if(lastRow_sheetTo < lastRow)  {
    Browser.msgBox('알림','Sheet1의 행의 수가 붙여넣기 할 FINAL보다 큽니다. FINAL행의 수를 Sheet1이상으로 늘려주세요. ',Browser.Buttons.OK)
    return 0;
  }

  sheetFrom.getRange(`A2:AQ${lastRow}`).activate();
  sheetTo.getRange('A2').activate();
  sheetFrom.getRange(`Sheet1!A2:AQ${lastRow}`).copyTo(sheetTo.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

}