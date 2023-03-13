/*

這是專為瀛海中學學務處開發的小小系統，非經本人同意，請勿外流給瀛海中學學務處以外人員
作者:溫孝文
日期:2023/2/23

*/

function GetAllSheetNames() {
let sheetNames = new Array()
let sheets = SpreadsheetApp.getActiveSpreadsheet().getSheets();
// console.log(sheets.length)
for (let i=0 ; i<sheets.length ; i++) sheetNames.push( [ sheets[i].getName() ] )
return sheetNames

}

// send to line notify
function sendLineNotify(message, lineTokens){
  let options;
  for(let i=0;i<lineTokens.length;i++){

    options =
    {
      "method"  : "post",
      "payload" : {"message" : message},
      "headers" : {"Authorization" : "Bearer " + lineTokens[i]}
    };
    UrlFetchApp.fetch("https://notify-api.line.me/api/notify", options);
  }
}


// 資料來源: https://stackoverflow.com/questions/19223823/google-script-trigger-weekdays-only
// 工作日10:00提醒學務同仁早自修沒資料
function createTriggerMorningNoDataClass() {
  var days = [ScriptApp.WeekDay.MONDAY, ScriptApp.WeekDay.TUESDAY, ScriptApp.WeekDay.WEDNESDAY, ScriptApp.WeekDay.THURSDAY, ScriptApp.WeekDay.FRIDAY];
  for (let i=0; i<days.length; i++) {
    ScriptApp.newTrigger("sendMorningClassNoData").timeBased().onWeekDay(days[i]).atHour(9).create();
   }
}

// 工作日15:00提醒學務同仁午休沒資料
function createTriggerNoonNoDataClass() {
  var days = [ScriptApp.WeekDay.MONDAY, ScriptApp.WeekDay.TUESDAY, ScriptApp.WeekDay.WEDNESDAY, ScriptApp.WeekDay.THURSDAY, ScriptApp.WeekDay.FRIDAY];
  for (let i=0; i<days.length; i++) {
    ScriptApp.newTrigger("sendNoonClassNoData").timeBased().onWeekDay(days[i]).atHour(14).create();
   }
}

// *datetime的日期格式，沒傳入日期，則傳入今天日期
function getDatetime(importDatetime){

  let datetime, date,dateShort;
  let timezone = SpreadsheetApp.getActive().getSpreadsheetTimeZone();

  // 如果無傳入值，預設為目前時間
  if (importDatetime !=null){
    dateHourMin = Utilities.formatDate(new Date(importDatetime), timezone, "yyyy/MM/dd HH:mm");
    date = Utilities.formatDate(new Date(importDatetime), timezone, "yyyy-MM-dd");
    datetime=Utilities.parseDate(importDatetime, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }else{
    dateHourMin = Utilities.formatDate(new Date(), timezone, "yyyy/MM/dd HH:mm");
    date = Utilities.formatDate(new Date(), timezone, "yyyy-MM-dd");
    dateShort= Utilities.formatDate(new Date(), timezone, "yyMMdd");
    datetime=new Date();
  }

  return [dateHourMin, date, dateShort,datetime]
}


// *發每日pdf訊息至國高中導師群組
function sendMessageToLine(dept,pdfUrl,dateTxt2,secTxt) {

  let message, deptTxt, secCTxt;
  let lineTokens=[]
  // eric_temp
  // let token = "i4npBEKuCBH5sUQ6LxuaOIDhsev3q5VzpnYs97wZW0u";
  // 2021學務網路測試_仁者無敵
  let token2 = "gzyQg3S3QzZltehQKTMWl5BEkJ6iczB7kAynCJ0pDp3";
  // 2023學務網路測試_學務測試
  let token3 = "3EiMf2mx0LWiSU070a771Q3YMTcECV6uxbfQmBGvrSt";
  // 2023學務處_110學務處
  let token4 = "eoxQYuy5mr9qeh4WsS0yXm1BvzvBWKiJvxz0bSMkLEb";

  // 高中導師群組line token 先用eric_temp代替
  let tokenH = "i4npBEKuCBH5sUQ6LxuaOIDhsev3q5VzpnYs97wZW0u";

  // 國中導師line token  先用eric_temp代替
  let tokenJ = "i4npBEKuCBH5sUQ6LxuaOIDhsev3q5VzpnYs97wZW0u";


  // 時段
  if (secTxt=='morning' || secTxt == null){
    secCTxt="早自修";

  }else if(secTxt=='noon'){
    secCTxt="午休";

  }

  // console.log(dept,pdfUrl,dateTxt2)

  if (dept!=null && pdfUrl !=null) {

    if(dept=='H'){

      lineTokens.push(tokenH);
      lineTokens.push(token4);
      deptTxt='高中部'

    }else{

      lineTokens.push(tokenJ);
      lineTokens.push(token4);
      deptTxt='國中部'

    }
  }

    message = "~測試中~_學務處巡視" + "\n";
    // message+="------------------------------------"+ "\n"+ "\n"+ "\n";
    message +="相關彙整表已上傳，請按下方連結，點閱pdf檔資料。如出現驗證畫面時，煩請同仁以「瀛海中學核發的gmail」(例如 ooo@gm.yhsh.tn.edu.tw )身份登入。" + "\n"+ "\n";
    message +="上傳時間: " +  dateTxt2 + "\n"+ "\n";
    message +="檔案說明: " + deptTxt+' '+  dateTxt2.substr(5,5) + secCTxt + " 巡視彙整表" + "\n"+ "\n";
    message +="檔案連結網址: " + pdfUrl + " \n"+ "\n"+ "\n";
    // message+="------------------------------------"+ "\n";
    // Logger.log(message);

  // 延遲送出
  // Utilities.sleep(8000);

  sendLineNotify(message, lineTokens);

}


// *回傳以年級分群的班級並排序
function  getSortedGroupEclass(eclassArr){

  let eclassTxt='';
  let grade='';

  // 排序eclass
  eclassArr=eclassArr.sort();

  for (i=0;i<eclassArr.length;i++){

      // 年級跳行
      if (grade !=eclassArr[i].slice(0,2)){
        eclassTxt+= '\n'
        eclassTxt+= eclassArr[i].slice(0,2) + '\n'
      }

      eclassTxt +=eclassArr[i]
      grade= eclassArr[i].slice(0,2)

      // 防止最後一個值沒辦法還有下一個可以比較
      if (i>=eclassArr.length-1){
        break;
      }

      // 結尾年級跳行
      if(grade !=eclassArr[i+1].slice(0,2)) {
        eclassTxt+= '\n'
      }else{
        eclassTxt+= ', '
      }
  }

  // console.log(eclassTxt);
  return eclassTxt;
}

// !以下未完成，先放著
// 確認是否已完成全部匯入該時段_temp
function checkSecImportNum(secText){

  // 抓第一頁的班級數
  // let firstShtName='011_daily_morning';
  let firstShtName='012_daily_noon';

  secText='morning';

  // 依時段決定開哪個匯入檔
  let importedShtName=(secText=='morning')?'021_data_morning' :'022_data_noon';

  let curDate=getDatetime()[1];
  // console.log(curDate);

  let spreadSheet = SpreadsheetApp.getActive();
  let firstSheet=spreadSheet.getSheetByName(firstShtName);
  let importedSheet=spreadSheet.getSheetByName(importedShtName);


  let firstShtRange=firstSheet.getRange(3,1,firstSheet.getLastRow(),1).getValues();
  let importedShtRange=importedSheet.getRange(3,18,importedSheet.getLastRow(),1).getValues();

  let firstShtCount=firstShtRange.filter(String).length;
  // let importedFilterCount=importedShtRange.filter(function(e){return e.getDay()==curDate}).length;
  // let aa=importedShtRange.filter(function(e){return e[1].getDay()==curDate}).length;

  // console.log(firstShtCount);

  console.log(importedShtRange[220])


}



// *取得上下午的文字
function getSecTextFromSheet(shtName){

  let secText,secCtext;
  if(shtName=='021_data_morning'){
    secText='morning';
    secCtext='早自修'
  }else if(shtName=='022_data_noon'){
    secText='noon';
    secCtext='午休'
  }

  return [secText, secCtext];
}

// *取得學期第一個周一
function getSemiMonday(){
  let sheetTabNameToGet = "002_params";
  let ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(sheetTabNameToGet);
  let semiMonday=sh.getRange('b2').getValue();
  return semiMonday;
}

// *計算周別，沒參數則算到目前日期的周別
function getSchWeek(date){

  let endDate,weekNum, dayNum, weekDay;
  let startMonday=getSemiMonday();

  endDate=getDatetime(date)[3];

  if(startMonday>endDate){
    console.log("ERROR!!")
  }else{

    weekDay=endDate.getDay();
    dayNum=Number(Math.floor((endDate-startMonday)/(24*3600*1000)).toFixed(0));
    weekNum=Number((Math.floor(dayNum-weekDay+1)/7).toFixed(0))+1;
  }
  return [weekDay, dayNum ,weekNum];
}





