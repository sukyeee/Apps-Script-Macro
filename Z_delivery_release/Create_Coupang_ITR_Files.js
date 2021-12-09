function Create_Coupang_ITR_Files(){

  //르엠마/엘이엠 저장할 제트배송 폴더 생성
  var todayMonth = new Date().getMonth()+1;
  var todayDate = new Date().getDate()
  var today = todayMonth + '/' + todayDate
  
  var folder_name =` ${today}제트배송_쿠팡입고요청`
  var folder_id = DriveApp.createFolder(folder_name).getId();

let sheet_RE = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("르엠마_제트배송")
let lastRow = sheet_RE.getLastRow(); //마지막 데이터가있는 행    

//숨겨진 행이 있는지 체크, 숨겨진행있으면 createFIlter 불가. ㅡ 있다면 필터 해제 (시간 너무 오래걸림)
// for(let i=4;i<lastRow;i++){ 
//  if(sheet_RE.isRowHiddenByFilter(i) ){
//    console.log(i)
//    var spreadsheet = SpreadsheetApp.getActive();
//       spreadsheet.getRange('A1:T1').activate();
//       spreadsheet.getActiveSheet().getFilter().remove();
//    break;
//  }
// }

if( sheet_RE.getFilter() == null) sheet_RE.getRange(`A1:T${lastRow}`).createFilter();

//I4~ 데이터 내용 지우기
sheet_RE.getRange(`I4:I${lastRow}`).activate();
sheet_RE.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});

sheet_RE.getRange('I4').activate();
sheet_RE.getCurrentCell().setFormula('=VLOOKUP(E4,FINAL!A:F,6,false)');
sheet_RE.getActiveRange().autoFillToNeighbor(SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES); //자동채우기

sheet_RE.getRange('A1:T1').activate();
var criteria = SpreadsheetApp.newFilterCriteria()
.setHiddenValues(['#N/A'])
.build();
sheet_RE.getFilter().setColumnFilterCriteria(9, criteria);

//다른 시트로 복사 후, 저장해야함. (업로드파일은 필터 걸려있는 상태 x, 값만 유지)
var spreadsheet = SpreadsheetApp.getActive();
spreadsheet.insertSheet(`르엠마쿠팡입고요청`);
// var currentCell = spreadsheet.getCurrentCell();
// spreadsheet.getActiveRange().getDataRegion().activate();
// currentCell.activateAsCurrentCell();
spreadsheet.setActiveSheet(spreadsheet.getSheetByName('르엠마쿠팡입고요청'), true);
sheet_RE.getRange(`\'르엠마_제트배송\'!A1:T${lastRow}`).copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

//xlsx로 저장

  var blob,exportUrl,options,pdfFile,response,sheetTabNameToGet,sheetTabId,ss,ssID,url_base; 
        lastRow = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('르엠마쿠팡입고요청').getLastRow(); //마지막 데이터가있는 행    
        var range = range ? range : `A1:T${lastRow}`; //저장할 범위
        console.log(lastRow);
        sheetTabNameToGet = `르엠마쿠팡입고요청`;//Replace the name with the sheet tab name for your situation
        ss = SpreadsheetApp.getActiveSpreadsheet() ;//This assumes that the Apps Script project is bound to a G-Sheet
        ssID = ss.getId();
        sh = ss.getSheetByName(sheetTabNameToGet);
        sheetTabId = sh.getSheetId();
      
        exportUrl = 'https://docs.google.com/spreadsheets/d/' +ssID+ '/export?exportFormat=xlsx&format=xlsx' +
          '&gid=' + sheetTabId + '&id=' + ssID +
          '&range=' + range ;      // do not repeat row headers (frozen rows) on each page

        options = {
          headers: {
            'Authorization': 'Bearer ' +  ScriptApp.getOAuthToken(),
          }
        }

        options.muteHttpExceptions = true;//Make sure this is always set
        response = UrlFetchApp.fetch(exportUrl, options);

        if (response.getResponseCode() !== 200) {
          console.log("Error exporting Sheet to PDF!  Response Code: " + response.getResponseCode());
          return;
        }
        
        blob = response.getBlob();
        blob.setName(`${today}르엠마_제트배송.xlsx`)
        DriveApp.getFolderById(folder_id).createFile(blob).getId(); //파일 이동

//필터 해제
var criteria = SpreadsheetApp.newFilterCriteria()
.build();
sheet_RE.getFilter().setColumnFilterCriteria(9, criteria); //I열 필터해제

SpreadsheetApp.getActive().deleteSheet(SpreadsheetApp.getActive().getSheetByName(`르엠마쿠팡입고요청`)); //추가된 시트 삭제
       

//-----------------------------------------엘이엠-----------------------------------------------------------------------------------

let sheet_LE = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("엘이엠_제트배송")
  lastRow = sheet_LE.getLastRow(); //마지막 데이터가있는 행    

//숨겨진 행이 있는지 체크, 숨겨진행있으면 createFIlter 불가. ㅡ 있다면 필터 해제 (시간 너무 오래걸림)
// for(let i=4;i<lastRow;i++){ 
//  if(sheet_LE.isRowHiddenByFilter(i) ){
//    console.log(i)
//    var spreadsheet = SpreadsheetApp.getActive();
//       spreadsheet.getRange('I1').activate();
//       spreadsheet.getActiveSheet().getFilter().remove();
//    break;
//  }
// }

if( sheet_LE.getFilter() == null)sheet_LE.getRange(`A1:T${lastRow}`).createFilter();

//I4~ 데이터 내용 지우기
sheet_LE.getRange(`I4:I${lastRow}`).activate();
sheet_LE.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});

sheet_LE.getRange('I4').activate();
sheet_LE.getCurrentCell().setFormula('=VLOOKUP(E4,FINAL!A:F,6,false)');
sheet_LE.getActiveRange().autoFillToNeighbor(SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES); //자동채우기

sheet_LE.getRange('A1:T1').activate();
var criteria = SpreadsheetApp.newFilterCriteria()
.setHiddenValues(['#N/A'])
.build();
sheet_LE.getFilter().setColumnFilterCriteria(9, criteria);


//다른 시트로 복사 후, 저장해야함. (업로드파일은 필터 걸려있는 상태 x, 값만 유지)
var spreadsheet = SpreadsheetApp.getActive();
spreadsheet.insertSheet(`엘이엠쿠팡입고요청`);
spreadsheet.setActiveSheet(spreadsheet.getSheetByName('엘이엠쿠팡입고요청'), true);
sheet_RE.getRange(`\'엘이엠_제트배송\'!A1:T${lastRow}`).copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_VALUES, false);

//xlsx로 저장
  var blob,exportUrl,options,pdfFile,response,sheetTabNameToGet,sheetTabId,ss,ssID,url_base; 
        lastRow = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('엘이엠쿠팡입고요청').getLastRow(); //마지막 데이터가있는 행    
        var range = range ? range : `A1:T${lastRow}`; //저장할 범위
  
        sheetTabNameToGet = `엘이엠쿠팡입고요청`;//Replace the name with the sheet tab name for your situation
        ss = SpreadsheetApp.getActiveSpreadsheet() ;//This assumes that the Apps Script project is bound to a G-Sheet
        ssID = ss.getId();
        sh = ss.getSheetByName(sheetTabNameToGet);
        sheetTabId = sh.getSheetId();
      
        exportUrl = 'https://docs.google.com/spreadsheets/d/' +ssID+ '/export?exportFormat=xlsx&format=xlsx' +
          '&gid=' + sheetTabId + '&id=' + ssID +
          '&range=' + range ;      // do not repeat row headers (frozen rows) on each page

        options = {
          headers: {
            'Authorization': 'Bearer ' +  ScriptApp.getOAuthToken(),
          }
        }

        options.muteHttpExceptions = true;//Make sure this is always set
        response = UrlFetchApp.fetch(exportUrl, options);

        if (response.getResponseCode() !== 200) {
          console.log("Error exporting Sheet to PDF!  Response Code: " + response.getResponseCode());
          return;
        }
        
        blob = response.getBlob();
        blob.setName(`${today}엘이엠_제트배송.xlsx`)
        DriveApp.getFolderById(folder_id).createFile(blob).getId(); //파일 이동

//필터 해제
var criteria = SpreadsheetApp.newFilterCriteria()
.build();
sheet_LE.getFilter().setColumnFilterCriteria(9, criteria); //I열 필터해제

SpreadsheetApp.getActive().deleteSheet(SpreadsheetApp.getActive().getSheetByName(`엘이엠쿠팡입고요청`)); //추가된 시트 삭제

}
