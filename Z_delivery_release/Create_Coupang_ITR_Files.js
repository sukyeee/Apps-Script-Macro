function Create_Coupang_ITR_Files(){

  // 쿠팡입고요청 제트배송 폴더 생성
  var todayMonth = new Date().getMonth()+1;
  var todayDate = new Date().getDate()
  var today = todayMonth + '/' + todayDate
  
  var folder_name =` ${today}제트배송_쿠팡입고요청`
  var folder_id = DriveApp.createFolder(folder_name).getId();

// 엘이엠 시트 / 데이터 없으면 오류메시지
if(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("엘이엠_제트배송") == null) {
      Browser.msgBox('알림',' 엘이엠_제트배송  시트가 존재하지 않습니다. ',Browser.Buttons.OK)
      return 0;
}
let sheet_LE = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("엘이엠_제트배송")
let lastRow_LE = sheet_LE.getLastRow(); //마지막 데이터가있는 행    
let sheet_RE = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("르엠마_제트배송")
let lastRow_RE = sheet_RE.getLastRow(); //마지막 데이터가있는 행    
let sheet_FINAL = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FINAL")
let lastRow = sheet_FINAL.getLastRow();
let maxRow = sheet_FINAL.getMaxRows(); //데이터와 상관없는 마지막행

//숨겨진 행이 있는지 체크, 숨겨진행있으면 createFIlter 불가. ㅡ 있다면 필터 해제 (시간 너무 오래걸림)
if(sheet_RE.isRowHiddenByFilter(4) ||sheet_RE.getFilter() != null ){
    sheet_RE.getRange('A3:U3').activate();
    sheet_RE.getFilter().remove();
}
if( sheet_RE.getFilter() == null) sheet_RE.getRange(`A3:U${lastRow_RE}`).createFilter();

//sheet_FINAL 필터유무 확인 

if( sheet_FINAL.getFilter() == null) sheet_FINAL.getRange(`A1:AQ${lastRow}`).createFilter();
var criteria = SpreadsheetApp.newFilterCriteria()
.build();
sheet_FINAL.getFilter().setColumnFilterCriteria(12, criteria); //L열 필터해제

// 1. 엘이엠_제트배송 -> 르엠마_제트배송 데이터 합치기  
  let maxRow_RE = sheet_RE.getMaxRows();
  sheet_RE.getRange(`${maxRow_RE}:${maxRow_RE}`).activate();
  sheet_RE.insertRowsAfter(sheet_RE.getActiveRange().getLastRow(), 1);
  //maxRow에서 아래 행 1개 추가하기
  sheet_RE.getRange(`A${lastRow_RE+1}`).activate();
  sheet_RE.getRange(`\'엘이엠_제트배송\'!A4:T${lastRow_LE}`).copyTo(sheet_RE.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  lastRow_RE = sheet_RE.getLastRow(); //르엠마+엘이엠 합친 lastRow로 업데이트


// 2. "A" (특정)그룹만 필터 건 후 Vlookup 
//초기작업 -> 존재하는 그룹 배열에 넣기
//A에 데이터 있는거 빼고 아래 데이터 다 지우기

if(sheet_FINAL.getRange(`A${maxRow}`).getValue() == '') {
  sheet_FINAL.getRange(`${maxRow}:${maxRow}`).activate();
  let lastDataRow = sheet_FINAL.getSelection().getNextDataRange(SpreadsheetApp.Direction.UP).getRow();
  sheet_FINAL.getRange(`${lastDataRow + 1}:${maxRow}`).activate();
  sheet_FINAL.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
} 
  //빈 행 제거하기
  // console.log(lastRow)
  spreadsheet = SpreadsheetApp.getActive();
  for(let i=2;i<=lastRow;i++){
    if(sheet_FINAL.getRange(`A${i}`).getValue() == ''){
         sheet_FINAL.getRange(`${i}:${i}`).activate();
         SpreadsheetApp.getActive().getActiveSheet().deleteRows(spreadsheet.getActiveRange().getRow(), spreadsheet.getActiveRange().getNumRows());
    }
  }

spreadsheet = SpreadsheetApp.getActive();
spreadsheet.getRange('AQ:AQ').activate();
spreadsheet.getRange('L:L').copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
spreadsheet.getRange(`AQ2:AQ${lastRow}`).activate();
//필터 걸지 않은 상태에서 중복 제거하기!!
spreadsheet.getActiveRange().removeDuplicates().activate();

//오름차순 정렬 (행 공백때문에 필요함)
spreadsheet = SpreadsheetApp.getActive();
spreadsheet.getActiveSheet().getFilter().sort(43, true);

// 1. 박스 그룹별로 배열에 저장
let box_group = new Array;
let box_all = new Array;
let box = 2; // A:2 , B:3 , C:4
for(let i = 0; i>=0; i++){

      if(SpreadsheetApp.getActiveSpreadsheet().getRange(`AQ${box}`).getValue() != '') {
          SpreadsheetApp.getActiveSpreadsheet().getRange(`AR${box}`).setValue(`=LEFT(AQ${box}, 1)`); 
      }
      else break;
        box_all[i] = SpreadsheetApp.getActiveSpreadsheet().getRange(`AQ${box}`).getValue();
        box_group[i] = SpreadsheetApp.getActiveSpreadsheet().getRange(`AR${box}`).getValue();
     box++;

}

for(let i=0;i<box_group.length;i++){
      //0:A, 1:B, 2:C ... 순서
       console.log(box_group[i]);


//I4, U4 데이터 내용 지우기
sheet_RE.getRange(`I4:I${lastRow_RE}`).activate();
sheet_RE.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
sheet_RE.getRange(`U4:U${lastRow_RE}`).activate();
sheet_RE.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});

sheet_RE.getRange('I4').activate();
sheet_RE.getCurrentCell().setFormula('=VLOOKUP(E4,FINAL!A:F,6,false)');
sheet_RE.getActiveRange().autoFillToNeighbor(SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES); //자동채우기
sheet_RE.getRange('U4').activate();
sheet_RE.getCurrentCell().setFormula('=VLOOKUP(E2,FINAL!A:L,12,false)');
sheet_RE.getActiveRange().autoFillToNeighbor(SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES); //자동채우기

var criteria = SpreadsheetApp.newFilterCriteria()
.setHiddenValues(['#N/A'])
.build();
sheet_RE.getFilter().setColumnFilterCriteria(9, criteria); //I열

criteria = SpreadsheetApp.newFilterCriteria()
.whenTextContains(`${box_group[i]}`)
.build();
sheet_RE.getFilter().setColumnFilterCriteria(21, criteria); // U열


//필터 건 후 loastRow, maxRow출력
console.log(box_group[i] , sheet_RE.getLastRow(), sheet_RE.getMaxRows());

//xlsx 파일 생성
    spreadsheet.insertSheet(`${box_group[i]} group`); //A group
    // spreadsheet.setActiveSheet(spreadsheet.getSheetByName(`${box_group[i]} group`), true);
    spreadsheet.getRange('A1').activate();
    lastRow_RE = sheet_RE.getLastRow(); //마지막 데이터가있는 행
    console.log('lastRow_RE ', lastRow_RE);
    spreadsheet.getRange(`르엠마_제트배송!A1:T${lastRow_RE}`).copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
  
      
      //folder_name안에 xlsx 파일로 저장
        let sheet_group = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(`${box_group[i]} group`);
        let lastRow_sheet_group = sheet_group.getLastRow(); //마지막 데이터가있는 행    
        if(lastRow_sheet_group == 3) { //데이터가 아무것도 없으면 (헤더만 있으면)
              //A, B, C.. 시트 삭제
            SpreadsheetApp.getActive().deleteSheet(SpreadsheetApp.getActive().getSheetByName(`${box_group[i]} group`));

            //필터 해제 하기
            criteria = SpreadsheetApp.newFilterCriteria()
            .build();
            sheet_RE.getFilter().setColumnFilterCriteria(9, criteria); //I열 필터해제
            sheet_RE.getFilter().setColumnFilterCriteria(21, criteria); //U열 필터해제
            continue;
        }
        var blob,exportUrl,options,pdfFile,response,sheetTabNameToGet,sheetTabId,ss,ssID,url_base; 
        var range = range ? range : `A1:T${lastRow_sheet_group}`; //저장할 범위
  
        sheetTabNameToGet = `${box_group[i]} group`;//Replace the name with the sheet tab name for your situation
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
        blob.setName(`${box_group[i]} group.xlsx`)
        DriveApp.getFolderById(folder_id).createFile(blob).getId(); //파일 이동

        //A, B, C.. 시트 삭제
       SpreadsheetApp.getActive().deleteSheet(SpreadsheetApp.getActive().getSheetByName(`${box_group[i]} group`));

      //필터 해제 하기
      criteria = SpreadsheetApp.newFilterCriteria()
      .build();
      sheet_RE.getFilter().setColumnFilterCriteria(9, criteria); //I열 필터해제
      sheet_RE.getFilter().setColumnFilterCriteria(21, criteria); //U열 필터해제
    
      //------------------------
  }

  //모두 필터 해제 하기
criteria = SpreadsheetApp.newFilterCriteria()
.build();
sheet_RE.getFilter().setColumnFilterCriteria(9, criteria); //J열 필터해제
sheet_RE.getFilter().setColumnFilterCriteria(21, criteria); //L열 필터해제
sheet_RE.getRange('A1').activate();

}