function RESET() {
  let sheetFrom = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("ezadmin");
  sheetFrom.getRange("A2:AQ").clearContent();

//FINAL 필터 해제 후 빈 행 제거하기
    var spreadsheet = SpreadsheetApp.getActive();
    let sheet_FINAL = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FINAL");
    let lastRow = sheet_FINAL.getLastRow(); //마지막 데이터가있는 행
//필터 해제
   var criteria = SpreadsheetApp.newFilterCriteria()
  .build();
  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(10, criteria); //J열 필터해제
  spreadsheet.getActiveSheet().getFilter().setColumnFilterCriteria(12, criteria); //L열 필터해제
  
    for(let i=2;i<=lastRow;i++){
      if(sheet_FINAL.getRange(`A${i}`).getValue() == ''){
           sheet_FINAL.getRange(`${i}:${i}`).activate();
           SpreadsheetApp.getActive().getActiveSheet().deleteRows(spreadsheet.getActiveRange().getRow(), spreadsheet.getActiveRange().getNumRows());
           lastRow = lastRow - 1;
      }
    }
}

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

function Print_FINAL(){
    // 1. 데이터 재정렬
    let sheet_FINAL = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("FINAL")
  
    sheet_FINAL.sort(10, true);
    sheet_FINAL.sort(11, true);
    sheet_FINAL.sort(12, true);

    // 2. 서식 정리
    let lastRow = sheet_FINAL.getLastRow(); //마지막 데이터가있는 행    
    // console.log(sheet_FINAL.getRange(`L2`).getValue() )
    for(let i=2;i<lastRow;i++){
      
      if(sheet_FINAL.getRange(`L${i}`).getValue() != sheet_FINAL.getRange(`L${i+1}`).getValue() ){
           sheet_FINAL.insertRowsAfter(sheet_FINAL.getRange(`L${i}`).getLastRow(), 1);
             i++;
           sheet_FINAL.getRange(`A${i}:AQ${i}`).setBackground('BACKGROUND');
         
      }
    }

    // 3. FINAL 시트내용을 PDF로 저장/다운
      var blob,exportUrl,options,pdfFile,response,sheetTabNameToGet,sheetTabId,ss,ssID,url_base; 
      var range = range ? range : `B1:L${lastRow}`; //저장할 범위
      
      var todayMonth = new Date().getMonth()+1;
      var todayDate = new Date().getDate()
      var today = todayMonth + '/' + todayDate
      
      sheetTabNameToGet = "FINAL";//Replace the name with the sheet tab name for your situation
      ss = SpreadsheetApp.getActiveSpreadsheet() ;//This assumes that the Apps Script project is bound to a G-Sheet
      ssID = ss.getId();
      sh = ss.getSheetByName(sheetTabNameToGet);
      sheetTabId = sh.getSheetId();
      // url_base = ss.getUrl().replace(/edit$/,'');

      exportUrl = 'https://docs.google.com/spreadsheets/d/' +ssID+ '/export?exportFormat=pdf&format=pdf' +
        '&gid=' + sheetTabId + '&id=' + ssID +
        '&range=' + range + 
        //'&range=NamedRange +
        '&size=A4' +     // paper size
        '&portrait=true' +   // orientation, false for landscape
        '&fitw=true' +       // fit to width, false for actual size
        `&sheetnames=false&printtitle=true&pagenumbers=true$dates=true` + //hide optional headers and footers
        '&gridlines=true' + // hide gridlines
        '&fzr=true';       // do not repeat row headers (frozen rows) on each page

      //Logger.log('exportUrl: ' + exportUrl)
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
      blob.setName(`${today} 로켓제휴 제트배송 출고.pdf`)
      pdfFile = DriveApp.createFile(blob);//Create the PDF file
     
      //Logger.log('pdfFile ID: ' +pdfFile.getId())
}


//--------------------------------------------------------------------------------------------------



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

//---------------------------------------------------------------------------------------------------------------------------------------

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
  sheet_RE.getRange(`I4`).autoFill(sheet_RE.getRange(`I4:I${lastRow_RE}`),SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);
  sheet_RE.getRange(`U4`).activate();
  sheet_RE.getCurrentCell().setFormula('=VLOOKUP(E4,FINAL!A:L,12,false)');
  sheet_RE.getRange(`U4`).autoFill(sheet_RE.getRange(`U4:U${lastRow_RE}`),SpreadsheetApp.AutoFillSeries.DEFAULT_SERIES);


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

