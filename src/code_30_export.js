// 製作weekpdf
function saveWeekPDFs(grade, sht, week){

  let pdfFileName, secCText, secText, eGrade, range, weekNum, semiStartDay, gradeTxt;
  let gradeArr=[];
  let shtArr=[];
  let exportSheetName='045_dosa_week_report';
  let receivePdfArr=[]; // 接收pdf參數
  let pdfDetailArr=[];
  let tempArr=[]

  let datetime=getDatetime()[0];
  let dateTxt = getDatetime()[2]; // yyMMdd
  // 左側補零
  const zeroPadLeft=(num, places)=>String(num).padStart(places,0);

  // todo:單頁或全部pdf
  // console.log('g1',grade);
  // console.log('g2',sht);
  // console.log('g3',week);

  // 年級 沒傳入值代表全部年級
  // if (grade){
  //   gradeArr.push(grade)
  // }else{
  //   gradeArr=["H1","H2","H3","J1","J2","J3"]
  // }
  gradeArr=["H1","H2"]
  // console.log(gradeArr);

  // 時段，沒傳入值代早、午都處理
  if(sht){
    shtArr.push(sht)
  }else{
    shtArr=["021_data_morning","022_data_noon"];
  }
  // console.log('sht:',sht);
  // console.log('shtArr:',shtArr);

  // 周別
  weekNum=(week)?week: getSchWeek()[2];

  for (let i=0;i<gradeArr.length;i++){
        eGrade=gradeArr[i];
        ss = SpreadsheetApp.getActiveSpreadsheet();
        sh = ss.getSheetByName(exportSheetName);
        sh.getRange('c1').setValue(eGrade)
        sh.getRange('e1').setValue(weekNum)
        // 決定列印範圍，j1、j2只有7班
        if (eGrade =='J1' ||  eGrade=='J2'){
          range = "C4:O12";
        }else{
          range = "C4:O13";
        }
    for (let j=0;j<shtArr.length;j++){
      // 取得時段文字
      secCText=getSecTextFromSheet(shtArr[j])[1]
      secText=getSecTextFromSheet(shtArr[j])[0]
    }
    // console.log(range);

    sh.getRange('h1').setValue(secCText)

    Utilities.sleep(3000);

    pdfFileName=dateTxt + '_week' + zeroPadLeft(weekNum,2) + '_' + secText + '_' + eGrade

    // console.log(sh.getRange('c1').getValue())
    console.log(eGrade);
    // console.log(weekNum);
    // console.log(secCText);
    // console.log(exportSheetName);
    // console.log(range);
    console.log(pdfFileName);

    // saveToPdf(sheetName,exportRange, exportFileName)
    tempArr= saveToPdf(exportSheetName, range, pdfFileName);
    // console.log(tempArr);
    receivePdfArr={
      grade:eGrade,
      date:dateTxt,
      secText:secText,
      pdfFileName:tempArr[0],
      pdfUrl:tempArr[1],
    }
    // console.log((receivePdfArr));
    Utilities.sleep(2500);
    console.log('aaaa');
    pdfDetailArr.push(receivePdfArr);
  }
  // todo:傳訊息到line，國、高中部，整合完只傳一次就好

    // Utilities.sleep(10000);
    // 發佈到國高中導師群組
    // sendMessageToLine(dept,pdfUrl,dateTxt2, secTxt)

  console.log(pdfDetailArr);
}


// 製作並儲存pdf
function saveToPdf(sheetName, exportRange, exportFileName){

  let blob,exportUrl,options,pdfFile,response,sheetTabNameToGet,sheetTabId,ss,ssID,url_base;
  let fileName, secCText, eGrade, range, weekNum, semiStartDay, sheetLink;
  let pdfID,sheetTabNamePdfLink;

  ss = SpreadsheetApp.getActiveSpreadsheet();
  ssID = ss.getId();
  sh = ss.getSheetByName(sheetName);
  sheetTabId = sh.getSheetId();

  // 設定pdf links 存放 sheet
  sheetTabNamePdfLink='055_download_links';
  sheetLink=ss.getSheetByName(sheetTabNamePdfLink);

  // pdf create datetime
  pdfCreateTime=getDatetime()[0];

  // console.log(ssID,sh)

  url_base = ss.getUrl().replace(/edit$/,'');

  // Logger.log('url_base: ' + url_base)

  exportUrl = url_base + 'export?exportFormat=pdf&format=pdf' +
    '&gid=' + sheetTabId + '&id=' + ssID +
    '&range=' + exportRange +
    '&size=A4' +     // paper size
    '&portrait=true' +   // orientation, false for landscape
    '&fitw=true' +       // fit to width, false for actual size
    '&sheetnames=false&printtitle=false&pagenumbers=true' + //hide optional headers and footers
    '&gridlines=false' + // hide gridlines
    '&fzr=false';       // do not repeat row headers (frozen rows) on each page

  // Logger.log('exportUrl: ' + exportUrl)

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
  let driverFolder=DriveApp.getFoldersByName("800_dosa_rounds").next()
  fileName=exportFileName +'.pdf'
  blob.setName(fileName)
  pdfFile = driverFolder.createFile(blob);//Create the PDF file
  // pdfID=pdfFile.getId()
  pdfUrl='https://drive.google.com/file/d/'+ pdfID +'/view?usp=share_link'


  // Logger.log('pdfFile ID: ' +pdfID)

  // console.log(exportFileName,exportRange)
  // console.log(pdfUrl,fileName)
  Utilities.sleep(1000);
  return [exportFileName, pdfUrl];
}

function test2(){

  // const zeroPadLeft=(num, places)=>String(num).padStart(places,0);

  // console.log(zeroPadLeft(2,2));
  // console.log(zeroPadLeft(22,2));
  let sheetName='021_data_morning';
  let grade="H1";
  let week=4;
  let tempArr=[]
  let tempArr2=[]

  // saveToPdf(sheetName,exportRange, exportFileName  )
  // saveWeekPDFs(grade, sheetName, week);
  saveWeekPDFs(null,null,4);
  console.log('finish!!');

  // tempArr.q1=grade;
  // tempArr.q2=week;
  // tempArr.q3=sheetName;
  // tempArr={
  //   q1:grade,
  //   q2:week,
  //   q3:sheetName,
  // }

  // tempArr2.push(tempArr);
  // tempArr2.push(tempArr);
  // console.log(tempArr);
  // console.log(tempArr2);



}




