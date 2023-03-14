// 製作weekpdf
function saveWeekPDFs(grade, sht, week){

  let pdfFileName, secCText, secText, eGrade, range, weekNum, cGrade;
  let shtLink, linkSheetName, ss, sh, exportSheetName, dept, message, messageDesc;
  let gradeArr=[];
  let shtArr=[];
  let receivePdfArr=[]; // 接收pdf參數
  let pdfDetailArr=[];
  let tempArr=[];

  let datetime=getDatetime()[0];
  let dateTxt = getDatetime()[2]; // yyMMdd
  let date= getDatetime()[1];

  exportSheetName='045_dosa_week_report';
  linkSheetName='055_download_links';
  ss = SpreadsheetApp.getActiveSpreadsheet();
  sh = ss.getSheetByName(exportSheetName);
  shtLink = ss.getSheetByName(linkSheetName);

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
  gradeArr=["H1","J2"];
  // gradeArr=["H1"];
  gradeArr.sort();
  const deptArr=Array.from(new Set(gradeArr.map(s=>s[0])));
  console.log('gradeArr:',gradeArr);
  console.log('deptArr:',deptArr);

  // 時段，沒傳入值代早、午都處理
  if(sht){
    shtArr.push(sht)
  }else{
    shtArr=["021_data_morning","022_data_noon"];
  }
  // console.log('sht:',sht);
  console.log('shtArr:',shtArr);

  // 周別
  weekNum=(week)?week: getSchWeek()[2];
  console.log('weekNum:',weekNum);

  for (let k=0;k<deptArr.length;k++){

    // 表頭不要重覆
    message = "~測試中~_學務處巡視第"+ weekNum + "周彙整表" + "\n";
    message +="相關資料已上傳，請按下方連結，點閱pdf檔資料。如出現驗證畫面時，煩請同仁以「瀛海中學核發的gmail」(例如 ooo@gm.yhsh.tn.edu.tw )身份登入。" + "\n"+ "\n";

    for (let i=0;i<gradeArr.length;i++){
      eGrade=gradeArr[i];
      dept=eGrade.slice(0,1);
      messageDesc='';
      // 確認部別(deptArr[k])是否相同
      if(dept!=deptArr[k]){
        continue;
      }
      cGrade=getCGrade(eGrade);
      sh.getRange('c1').setValue(eGrade)
      sh.getRange('e1').setValue(weekNum)
      // 決定列印範圍，j1、j2只有7班
      if (eGrade =='J1' ||  eGrade=='J2'){
        range = "C4:O12";
      }else{
        range = "C4:O13";
      }
      message +="年級: 【" +  cGrade + " 】" "\n";

      for (let j=0;j<shtArr.length;j++){
        // 取得時段文字
        secCText=getSecTextFromSheet(shtArr[j])[1]
        secText=getSecTextFromSheet(shtArr[j])[0]
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
          secText:secText,
          datetime:tempArr[2],
          pdfFileName:tempArr[0],
          pdfUrl:tempArr[1],
        }

        // console.log((receivePdfArr));
        Utilities.sleep(2500);

        messageDesc +="檔案說明: " +"第" + weekNum + "周" + secCText + " 巡視彙整" + "\n";
        messageDesc +="檔案連結網址: " + pdfUrl + " \n"+ "\n"

        pdfDetailArr.push(receivePdfArr);
      }

      message += messageDesc;
    }
    message +="上傳日期: " +  date + "\n"+ "\n";
    console.log(message);
  }

  // todo:存入本次的連結資訊至 055_download_links

  // todo:傳訊息到line，國、高中部，整合完只傳一次就好

    // Utilities.sleep(10000);
    // 發佈到國高中導師群組
  console.log(pdfDetailArr);
}


// 製作並儲存pdf
function saveToPdf(sheetName, exportRange, exportFileName){

  let blob, exportUrl, options, pdfFile, response, sheetTabId, ss, sh, ssID, url_base;
  let fileName,  sheetLink;
  let pdfID, sheetTabNamePdfLink, pdfCreateTime;

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
  return [exportFileName, pdfUrl, pdfCreateTime];
}

function test2(){

  // const zeroPadLeft=(num, places)=>String(num).padStart(places,0);

  // console.log(zeroPadLeft(2,2));
  // console.log(zeroPadLeft(22,2));
  // let sheetName='021_data_morning';
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


function test3(){
  let ss, sh, sheetName;

  let tempArr=[ { grade: 'H1',
  date: '230313',
  secText: 'morning',
  pdfFileName: '230313_week04_morning_H1',
  pdfUrl: 'https://drive.google.com/file/d/undefined/view?usp=share_link' },
{ grade: 'H1',
  date: '230313',
  secText: 'noon',
  pdfFileName: '230313_week04_noon_H1',
  pdfUrl: 'https://drive.google.com/file/d/undefined/view?usp=share_link' },
{ grade: 'H2',
  date: '230313',
  secText: 'morning',
  pdfFileName: '230313_week04_morning_H2',
  pdfUrl: 'https://drive.google.com/file/d/undefined/view?usp=share_link' },
{ grade: 'H2',
  date: '230313',
  secText: 'noon',
  pdfFileName: '230313_week04_noon_H2',
  pdfUrl: 'https://drive.google.com/file/d/undefined/view?usp=share_link' },
{ grade: 'J2',
  date: '230313',
  secText: 'morning',
  pdfFileName: '230313_week04_morning_J2',
  pdfUrl: 'https://drive.google.com/file/d/undefined/view?usp=share_link' },
{ grade: 'J2',
  date: '230313',
  secText: 'noon',
  pdfFileName: '230313_week04_noon_J2',
  pdfUrl: 'https://drive.google.com/file/d/undefined/view?usp=share_link' } ]

  sheetName='055_download_links';
  ss = SpreadsheetApp.getActiveSpreadsheet();
  sh = ss.getSheetByName(sheetName);

  // sh.getRange(2,6,20,20).setValues(tempArr);
  // tempArr.keys
  // console.log(Object.values(tempArr)[1]);
  for (let row=0;row<tempArr.length;row++){
    for(let col=0;col<5;col++){
      sh.getRange(row+2,col+6,1,1).setValue(Object.values(tempArr[row])[col])
      // console.log(tempArr[i].grade);
      // console.log(tempArr[i].pdfUrl);
    }
  }
}

function test4() {
  let tempArr=['H1','H2','H2'];
  // let temp2=tempArr.map(s=>s[0]);
  const noDupArray=Array.from(new Set(tempArr.map(s=>s[0])));

  console.log(noDupArray);
}




