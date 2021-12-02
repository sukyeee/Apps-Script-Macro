function Create_Barcode_Files() {
  var todayMonth = new Date().getMonth()+1;
  var todayDate = new Date().getDate()
  var today = todayMonth + '/' + todayDate
       
  //folder 생성
    var folder_name =`test_` + today + `_제트배송_출고_바코드`
    var folder_id = DriveApp.createFolder(folder_name).getId();
 
  //0. 1번 초기작업
   let sheet_FINAL = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FINAL")
    //필터 해제
   var spreadsheet = SpreadsheetApp.getActive();
   var criteria = SpreadsheetApp.newFilterCriteria()
  .build();
  sheet_FINAL.getFilter().setColumnFilterCriteria(10, criteria); //J열 필터해제
  sheet_FINAL.getFilter().setColumnFilterCriteria(12, criteria); //L열 필터해제
  

        //A에 데이터 있는거 빼고 아래 데이터 다 지우기*****
   let maxRow = sheet_FINAL.getMaxRows(); //데이터와 상관없는 마지막행

  if(sheet_FINAL.getRange(`A${maxRow}`).getValue() == '') {
    sheet_FINAL.getRange(`${maxRow}:${maxRow}`).activate();
    let lastDataRow = sheet_FINAL.getSelection().getNextDataRange(SpreadsheetApp.Direction.UP).getRow();
    sheet_FINAL.getRange(`${lastDataRow + 1}:${maxRow}`).activate();
    sheet_FINAL.getActiveRangeList().clear({contentsOnly: true, skipFilteredRows: true});
    // console.log(maxRow, lastDataRow); 
  } 

    //빈 행 제거하기
    let lastRow = sheet_FINAL.getLastRow(); //마지막 데이터가있는 행    
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
  
       console.log(box_group[i]);
       box++;

  }

// 1. J열 필터로 엑셀 파일 사용값 항목만 선택
  criteria = SpreadsheetApp.newFilterCriteria()
    .whenTextContains('엑셀파일사용')
    .build();
    sheet_FINAL.getFilter().setColumnFilterCriteria(10, criteria);
 
 for(let i=0;i<box_group.length;i++){
        //0:A, 1:B, 2:C ... 순서
        // ----------------------
        // ------** first가 "엑셀 파일 사용" 행부터 시작해야함, 현재 필터 되어있지 않은상태로 출력됨 
  
  //필터 된 항목 중 데이터 범위 지정
  // let group_end = new Array; //A~A, B~B, C~C 범위
  // box = 0;
  // for(let i=2;i<=lastRow;i++){
  //     if(box_all[box] != SpreadsheetApp.getActiveSpreadsheet().getRange(`L${i}`).getValue()){
  //         group_end[box] = i-1;
  //         box++;
  //     }
  //     group_end[box] = lastRow;
  // }
  //       console.log(group_end)
  
    //2. B,F 열의 값만 박스 그룹별 엑셀 파일에 복붙
    criteria = SpreadsheetApp.newFilterCriteria()
    .whenTextContains(`${box_group[i]}`) //'A' or 'B' .... 포함된 것만 필터 
    .build();
    sheet_FINAL.getFilter().setColumnFilterCriteria(12, criteria);
    console.log(box_group[i]);


  lastRow = sheet_FINAL.getLastRow(); //마지막 데이터가있는 행    

// 시트 그룹별 ㅡ B,F 열의 값만 박스 그룹별 엑셀 파일에 복붙 

    // first = 2;
       
      // sheet_FINAL.getRangeList(['B:B', 'F:F']).activate();
      spreadsheet.insertSheet(`${box_group[i]} group`); //A group
      // spreadsheet.getRange('A1').setValue('상품코드');
      // spreadsheet.getRange('B1').setValue('수량');
      spreadsheet.getRange('A1').activate();
      spreadsheet.getRange(`FINAL!B:B`).copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
      // spreadsheet.setActiveSheet(spreadsheet.getSheetByName('FINAL'), true);
      // spreadsheet.getRange('F:F').activate();
      spreadsheet.setActiveSheet(spreadsheet.getSheetByName(`${box_group[i]} group`), true);
      spreadsheet.getRange('B1').activate();
      spreadsheet.getRange(`FINAL!F:F`).copyTo(spreadsheet.getActiveRange(), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
      // first = group_end[i]+1;

        //folder_name안에 xlsx 파일로 저장
          // let sheet_id = new Array;
          let sheet_group = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(`${box_group[i]} group`);
          // sheet_id[i] = SpreadsheetApp.getActiveSpreadsheet().getSheetId();
          let lastRow_sheet_group = sheet_group.getLastRow(); //마지막 데이터가있는 행    

          var blob,exportUrl,options,pdfFile,response,sheetTabNameToGet,sheetTabId,ss,ssID,url_base; 
          var range = range ? range : `A1:B${lastRow_sheet_group}`; //저장할 범위
    
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
         
        //------------------------
    }

    //모두 필터 해제 하기
  criteria = SpreadsheetApp.newFilterCriteria()
  .build();
  sheet_FINAL.getFilter().setColumnFilterCriteria(10, criteria); //J열 필터해제
  sheet_FINAL.getFilter().setColumnFilterCriteria(12, criteria); //L열 필터해제
  sheet_FINAL.getRange('A1').activate();
}