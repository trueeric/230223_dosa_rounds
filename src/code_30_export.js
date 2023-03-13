// 製作weekpdf
function saveWeekPDFs(grade=null,sht=null,week ){

  // let blob,exportUrl,options,pdfFile,response,sheetTabNameToGet,sheetTabId,ss,ssID,url_base;
  // let fileName, secCtext, eGrade, range, weekNum, semiStartDay;
  // let pdfID,sheetTabNamePdfLink;
  let gradeArr=[];
  let shtArr=[];

  let datetime=getDatetime(date)[0];
  let dateTxt = getDatetime(date)[2]; // yyMMdd
  let dateTxt2 = datetime;

  // todo:單頁 全部pdf

  // 年級 沒傳入值代表全部年級
  if (grade){
    gradeArr.push(grade)
  }else{
    gradeArr=["H1","H2","H3","J1","J2","J3"]
  }

  // 時段，沒傳入值代早、午都處理
  if(sht){

    shtArr.push(sht)
  }else{

    shtArr=["021_data_morning","022_data_noon"];
  }

  // 周別
  weekNum=getSchWeek(week)

  for (let j=0;j<shtArr.length;j++){
    // 取得時段文字
    secCtext=getSecText(shtArr[j][1])

    for (let i=0;i<gradeArr.length;i++){

      eGrade=gradeArr[i];

      // 決定列印範圍，j1、j2只有7班
      if (eGrade =='J1' ||  eGrade=='J2'){
        range = "C4:O12"; // Replace with your desired range
      }else{
        range = "C4:O13"; // Replace with your desired range
      }
      // console.log(range);


      // sheetTabNameToGet = "031_morning_form";
      ss = SpreadsheetApp.getActiveSpreadsheet();
      ssID = ss.getId();
      sh = ss.getSheetByName(shtArr[j]);
      sheetTabId = sh.getSheetId();
      sh.getRange('c1').setValue(gradeArr[i])
      sh.getRange('h1').setValue(secCtext)

      // console.log(sh.getRange('c1').getValue())

      let secTxt=sh.getRange('s1').getValue()

      // 設定pdf links sheet
      sheetTabNamePdfLink='055_download_links';
      sheetLink=ss.getSheetByName(sheetTabNamePdfLink);

      // todo:傳訊息到line
      // 發佈到國高中導師群組
      // sendMessgeToLine(dept,pdfUrl,dateTxt2, secTxt)
    }
  }

}


// 製作並儲存pdf
function saveToPdf(sheetName,exportRange, exportFileName){

  let blob,exportUrl,options,pdfFile,response,sheetTabNameToGet,sheetTabId,ss,ssID,url_base;
  let fileName, secCtext, eGrade, range, weekNum, semiStartDay, sheetLink;
  let pdfID,sheetTabNamePdfLink;

  ss = SpreadsheetApp.getActiveSpreadsheet();
  ssID = ss.getId();
  sh = ss.getSheetByName(sheetName);
  sheetTabId = sh.getSheetId();

  // 設定pdf links 存放 sheet
  sheetTabNamePdfLink='055_download_links';
  sheetLink=ss.getSheetByName(sheetTabNamePdfLink);

  // pdf create datetime
  pdfCreateTime=getDatetime()

  // console.log(ssID,sh)

  url_base = ss.getUrl().replace(/edit$/,'');

  // Logger.log('url_base: ' + url_base)

  exportUrl = url_base + 'export?exportFormat=pdf&format=pdf' +

    '&gid=' + sheetTabId + '&id=' + ssID +
    '&range=' + exportRange +
    //'&range=NamedRange +
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

  // Logger.log(response)
  // Logger.log(response.getResponseCode())

  if (response.getResponseCode() !== 200) {
    console.log("Error exporting Sheet to PDF!  Response Code: " + response.getResponseCode());
    return;
  }

  blob = response.getBlob();
  let driverFolder=DriveApp.getFoldersByName("800_dosa_rounds").next()

  fileName=exportFileName +'.pdf'

  blob.setName(fileName)

  pdfFile = driverFolder.createFile(blob);//Create the PDF file
  pdfID=pdfFile.getId()
  pdfUrl='https://drive.google.com/file/d/'+ pdfID +'/view?usp=share_link'
  Logger.log('pdfFile ID: ' +pdfID)

  // 發佈到國高中導師群組
  // sendMessgeToLine(dept,pdfUrl,dateTxt2, secTxt)
  console.log(exportFileName,exportRange)
  console.log(pdfUrl,fileName)
  return [dept,pdfUrl,dateTxt2, secTxt];


}

function test2(){

  let sheetName='045_dosa_week_report'
  let exportRange='C4:O13'
  let exportFileName='test126'

  saveToPdf(sheetName,exportRange, exportFileName  )

}




