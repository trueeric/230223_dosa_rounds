/*

這是專為瀛海中學學務處開發的小小系統，非經本人同意，請勿外流給瀛海中學學務處以外人員
作者:溫孝文
日期:2023/2/23

*/
//清空012每日午休工作表資料
function clearData_012_daily_noon() {
  let ss= SpreadsheetApp.getActive();
  let sheetName=ss.getSheetByName('012_daily_noon');
  let range=sheetName.getRange('c3:q48');
  range.setValue(null);
  // SpreadsheetApp.getUi().alert("012每日午休工作表資料清空完成!");

}

//當日午休資料歸檔至022_data_noon
function copyDailyNoonTo022_Data_noon(){
  // 設定來源檔資料範圍
  let spreadSheet = SpreadsheetApp.getActive();
  let sourceSheet=spreadSheet.getSheetByName('012_daily_noon');
  // 來源範圍要一直到email欄
  let sourceRange=sourceSheet.getRange('a3:q48');
  let sourceValues=sourceRange.getValues();
  let sourceRowCount=sourceValues.length;
  // let columnCount=sourceValue.length


  let targetSheet=spreadSheet.getSheetByName('022_data_noon');
  let targetLastRow=targetSheet.getLastRow();
  // console.log(targetLastRow)
  sourceRange.copyTo(targetSheet.getRange(targetLastRow+1,3));

  //插入第1欄_index函數 要改成 r1c1模式
  let targetIndexColumnRange=targetSheet.getRange(targetLastRow+1,1,sourceRowCount,1);
  let formula_0='=if(ISBLANK(INDIRECT("R[0]C[2]",false)),"",INDIRECT("R[0]C[2]",false)&if(isblank(INDIRECT("R[0]C[17]",false)),"","_"&text(INDIRECT("R[0]C[17]",false),"yymmdd")))';
  targetIndexColumnRange.setFormula(formula_0);

  //插入第2欄_id函數
  let targetDateColumnRange=targetSheet.getRange(targetLastRow+1,2,sourceRowCount,1);
  let formula_1='=row()-2';
  targetDateColumnRange.setFormula(formula_1);

  // 插入最後col_20姓名函數 要改成 r1c1模式
  let targetNameColumnRange=targetSheet.getRange(targetLastRow+1,20,sourceRowCount,1);
  let formula_20='=if(isblank(INDIRECT("R[0]C[-1]",false)),"",ifna(vlookup(INDIRECT("R[0]C[-1]",false),\'052_dosa_staff\'!$A$1:$B,2,0),""))';
  targetNameColumnRange.setFormula(formula_20);



   // 插入column_21 :=if(ISBLANK($C3),"",if(n3,1,0)) 導師是否隨班 沒有:0 有:1 無法確定:2 要改成 r1c1模式
  let targetColumn_21_range=targetSheet.getRange(targetLastRow+1,21,sourceRowCount,1);
  let formula_21='=if(INDIRECT("R[0]C[-7]",false)+INDIRECT("R[0]C[-6]",false),if(INDIRECT("R[0]C[-7]",false),0,1),2)';
  targetColumn_21_range.setFormula(formula_21);

  // 插入最後col_22 周別 函數 要改成 r1c1模式
  let targetWeekColumnRange=targetSheet.getRange(targetLastRow+1,22,sourceRowCount,1);
  let formula_22='=if(isblank(INDIRECT("R[0]C[-4]",false)),"",rounddown(DATEDIF(\'002_params\'!$B$2,INDIRECT("R[0]C[-4]",false),"D")/7)+1)';
  targetWeekColumnRange.setFormula(formula_22);


  //清空012每日午休工作表資料
  clearData_012_daily_noon() ;

}


// 製作當日或指定日的午休報表pdf
function saveDailyNoonPDF(date){

  let blob,exportUrl,options,pdfFile,response,sheetTabNameToGet,sheetTabId,ss,ssID,url_base;
  let printDate, fileName, cellPrintDate;
  let pdfID,sheetTabNamePdfLink;
  let deptArr=[];
  let dept='';
  let secTxt;

  let datetime=getDatetime(date)[0];
  let dateTxt = getDatetime(date)[2]; //yyMMdd
  let dateTxt2 = datetime;


  // 部別
  deptArr=["H","J"]
  for (let i=0;i<deptArr.length;i++){

    dept=deptArr[i];
    console.log(dept)
    if (dept=='H'){
      range = "C4:P29"; // Replace with your desired range
    }else if(dept=='J'){
      range = "C4:P27"; // Replace with your desired range
    }
    // console.log(range);


    sheetTabNameToGet = "032_noon_form";//Replace the name with the sheet tab name for your situation
    ss = SpreadsheetApp.getActiveSpreadsheet();//This assumes that the Apps Script project is bound to a G-Sheet
    ssID = ss.getId();
    sh = ss.getSheetByName(sheetTabNameToGet);
    sheetTabId = sh.getSheetId();
    sh.getRange('c1').setValue(dept);
    console.log(sh.getRange('c1').getValue());

    secTxt=sh.getRange('s1').getValue()

    // 設定pdf links sheet
    sheetTabNamePdfLink='055_download_links';
    shLink=ss.getSheetByName(sheetTabNamePdfLink);

    // console.log(ssID,sh)

    url_base = ss.getUrl().replace(/edit$/,'');

    // Logger.log('url_base: ' + url_base)

    exportUrl = url_base + 'export?exportFormat=pdf&format=pdf' +

      '&gid=' + sheetTabId + '&id=' + ssID +
      '&range=' + range +
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

    //Logger.log(response.getResponseCode())

    if (response.getResponseCode() !== 200) {
      console.log("Error exporting Sheet to PDF!  Response Code: " + response.getResponseCode());
      return;
    }

    blob = response.getBlob();
    let driverFolder=DriveApp.getFoldersByName("800_dosa_rounds").next()

    fileName=dateTxt+'_'+secTxt+'_'+dept+'.pdf'

    blob.setName(fileName)

    pdfFile = driverFolder.createFile(blob);//Create the PDF file
    pdfID=pdfFile.getId()
    pdfUrl='https://drive.google.com/file/d/'+ pdfID +'/view?usp=share_link'
    // Logger.log('pdfFile ID: ' +pdfID)
    // dept	date	secTxt	pdfUrl
    shLink.getRange(4+i,1,1,1).setValue(dept);
    shLink.getRange(4+i,2,1,1).setValue(dateTxt2);
    shLink.getRange(4+i,3,1,1).setValue(secTxt);
    shLink.getRange(4+i,4,1,1).setValue(pdfUrl);

    // 發佈到國、高中導師群組
    sendMessageToLine(dept,pdfUrl,dateTxt2, secTxt)
  }
}


function sendNoonClassNoData(){

  let sheetTabNameMorning = "012_daily_noon";
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(sheetTabNameMorning);
  let range=sh.getRange(3,1,46,16).getValues();
  let datetime=getDatetime()[1];
  let todayDate=datetime.slice(8,10);
  let eclassArr=[];
  let eclassTxt='';

  // let message='~學務巡視測試模擬 *周一至周五1400-1500* 自動發通知至學務相關群組提醒學務夥伴~'+ '\n' ;
  let message='❖學務班級巡視❖'+ '\n' ;

  let lineTokens=[]
  // eric_temp
  let token = "i4npBEKuCBH5sUQ6LxuaOIDhsev3q5VzpnYs97wZW0u";
  // 2021學務網路測試_仁者無敵
  // let token2 = "gzyQg3S3QzZltehQKTMWl5BEkJ6iczB7kAynCJ0pDp3";
  // 2023學務網路測試_學務測試
  // let token3 = "3EiMf2mx0LWiSU070a771Q3YMTcECV6uxbfQmBGvrSt";
  // 2023學務處_110學務處
  let token4 = "eoxQYuy5mr9qeh4WsS0yXm1BvzvBWKiJvxz0bSMkLEb";
  lineTokens.push(token)
  // lineTokens.push(token3)
  lineTokens.push(token4)

  let noDataClass=range.filter(function(e){return e[15] ==''})
  // console.log(function(e))
  // console.log(noDataClass.length)

  let grade=''

  if (noDataClass.length<=0) {
    message+='本('+ todayDate +')日午休巡視結果，各班資料 *皆已入檔* ，感謝好夥伴的協助!!' + '\n';

  }else{

    message+='本('+ todayDate +')日 *午休* 巡視，下列班級紀錄目前尚未入檔，煩請同仁協助處理，感謝您的幫忙!! ' + '\n';

      for (i=0;i<noDataClass.length;i++){

        eclassArr.push(noDataClass[i][0])
      }

      eclassTxt=getSortedGroupEclass(eclassArr);
      message+=eclassTxt + '\n' + '\n';
      message+='◉ 若已完成彙整作業或本日非巡視日，請忽略本次訊息 ◉';
      // console.log(eclassArr)
  }
  // console.log(message)
  sendLineNotify(message,lineTokens)

}
