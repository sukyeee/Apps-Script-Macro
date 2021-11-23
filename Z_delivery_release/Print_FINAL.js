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
   
}