/*

這是專為瀛海中學學務處開發的小小系統，非經本人同意，請勿外流給瀛海中學學務處以外人員
作者:溫孝文
日期:2023/2/23

*/


function myFunction() {
  let ss=SpreadsheetApp.openById('1ZOcY3rh9eZPsG7p7zDUIEuJksyhMgF6ISqGo1swnBco');
  let spreadSheet = SpreadsheetApp.getActive();
  let sheetName=ss.getSheetByName('032_noon_form');
  let v1=sheetName.getRange('s1').getValue();
  console.log(v1);


  // if (spreadSheet.getActiveSheet().getName()=='012_daily_noon') {
  //   Logger.log("aaa")
  // }else{
  //   Logger.log('bbb')
  // }
  // console.log(spreadSheet.getActiveSheet().getName())
  // let range=sheetName.getRange('b3:f48')
  // range.setValue(null)
  
  // let timezone = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
  // let date = Utilities.formatDate(new Date(), timezone, "yyyy/MM/dd HH:mm");
  // let datetxt1=Utilities.formatDate(new Date(date), timezone, "yyyy/MM/dd");
  // let range=sheetName.getRange(3,7);
  // range.setValue(date);
  // let currentTime=date;
  // console.log(datetxt1);
  // console.log(currentTime)
  // console.log('active',Session.getActiveUser().getEmail())
  // console.log('effect',Session.getEffectiveUser().getEmail())
  // console.log("Done!")
  
}



// 舊
// // 異動資料時在該列尾巴加入更新時間及user_gmail
// function onEdit(e){  
//   //check edit range onEdit(e)需要限定作用範圍
//   let editRange={
//     top:3,
//     bottom:48,
//     left:3,
//     right:14,
//   }
  
//   let thisRow=e.range.getRow();
//   if(thisRow<editRange.top || thisRow>editRange.bottom) return;

//   let thisCol=e.range.getColumn();
//   if(thisCol<editRange.left || thisCol>editRange.right) return;

//   let timezone = SpreadsheetApp.getActive().getSpreadsheetTimeZone();
//   let date = Utilities.formatDate(new Date(), timezone, "yyyy/MM/dd HH:mm");
//   let staffEmail=Session.getActiveUser().getEmail(); //getActiveUser才正確
//   console.log(Session.getActiveUser());
//   console.log(staffEmail);


//   let spreadSheet = SpreadsheetApp.getActive();
//   let sheetName=spreadSheet.getSheetByName('012_daily_noon');
//   sheetName.getRange(thisRow,16)
//            .setValue(date);
//   sheetName.getRange(thisRow,17)
//            .setValue(staffEmail);
//}


// function sendFileToLine(file_id,line_token) {
//   let fileId = "136eFZSExxjuk_GdgGgJUmXtdzLCMe9HZ"; // Replace with the ID of the file you want to send
//   let lineToken = "i4npBEKuCBH5sUQ6LxuaOIDhsev3q5VzpnYs97wZW0u";; // Replace with your Line access token

//   let file = DriveApp.getFileById(fileId);
//   let blob = file.getBlob();
//   let fileName = file.getName();

//   console.log(file)
  
//   let options = {
//     "method" : "post",
//     "headers" : {
//       "Authorization" : "Bearer " + lineToken,
//       "Content-Type" : "application/octet-stream",
//       "Content-Disposition" : "attachment; filename=\"" + fileName + "\""
//     },
//     "payload" : blob.getBytes()
//   };
  
//   UrlFetchApp.fetch("https://api-data.line.me/v2/bot/message/push", options);
// }



// function sendMessgeMorning(dept,pdfUrl) {
// // eric_temp
//   let token = "i4npBEKuCBH5sUQ6LxuaOIDhsev3q5VzpnYs97wZW0u";
// // 2021學務網路測試_仁者無敵
//  let token2 = "gzyQg3S3QzZltehQKTMWl5BEkJ6iczB7kAynCJ0pDp3";

// // 2021學務網路測試_仁者無敵
//   // let token3 = "s5qDJreFqU8MrGmNGdxF2QsDMeoD4D1Axhuu8I4FX5w";

// // 2022 功能測試_東霖
//   let token4 = "KzRYV5bKDdoNrIIe15LZ430LFBA7dFgFHUOo1mkoYic";

//   // 延遲時間，取最新資料
//   // Utilities.sleep(3000);

//   let pdf_morning_h_id= "136eFZSExxjuk_GdgGgJUmXtdzLCMe9HZ"

//   let message = "~測試中~_學務處巡視" + "\n";
//   message+="------------------------------------"+ "\n"; 
//   message += Utilities.formatDate(new Date(),'GMT+8','MM/dd')+"_學務處早自修巡視彙整表已上傳，請以瀛海中學gmail登入下方連結，點閱相關pdf檔" + "\n";
//   message += "https://drive.google.com/drive/folders/1wDmLPmoJdQjioHNiz2hoErvQ1gfypd-K" + "\n";
//   message+="------------------------------------"+ "\n"; 
    

//     Logger.log(message); 

//   // 延遲送出
//   // Utilities.sleep(8000);

//   // eric_temp
//   sendLineNotify(message, token);
  
//   // sendLineNotify2(message, token2);
//   // sendLineNotify3(message, token3);
//   // sendLineNotify4(message, token4);

//   // sendFileToLine(pdf_morning_h_id,token)

// }


// function sendLineNotify(message, token){
//   let options =
//    {
//      "method"  : "post",
//      "payload" : {"message" : message},
//      "headers" : {"Authorization" : "Bearer " + token}
//    };
//    UrlFetchApp.fetch("https://notify-api.line.me/api/notify", options);
// }

// function sendLineNotify2(message, token2){
//   let options =
//    {
//      "method"  : "post",
//      "payload" : {"message" : message},
//      "headers" : {"Authorization" : "Bearer " + token2}
//    };
//    UrlFetchApp.fetch("https://notify-api.line.me/api/notify", options);
// }

// function sendLineNotify3(message, token3){
//   let options =
//    {
//      "method"  : "post",
//      "payload" : {"message" : message},
//      "headers" : {"Authorization" : "Bearer " + token3}
//    };
//    UrlFetchApp.fetch("https://notify-api.line.me/api/notify", options);
// }

// function sendLineNotify4(message, token4){
//   let options =
//    {
//      "method"  : "post",
//      "payload" : {"message" : message},
//      "headers" : {"Authorization" : "Bearer " + token4}
//    };
//    UrlFetchApp.fetch("https://notify-api.line.me/api/notify", options);
// }

// function sendMessgeNoon() {
// // eric_temp
//   let token = "i4npBEKuCBH5sUQ6LxuaOIDhsev3q5VzpnYs97wZW0u";
// // 2021學務網路測試_仁者無敵
//  let token2 = "gzyQg3S3QzZltehQKTMWl5BEkJ6iczB7kAynCJ0pDp3";

// // 2021學務網路測試_仁者無敵
//   // let token3 = "s5qDJreFqU8MrGmNGdxF2QsDMeoD4D1Axhuu8I4FX5w";

// // 2022 功能測試_東霖
//   let token4 = "KzRYV5bKDdoNrIIe15LZ430LFBA7dFgFHUOo1mkoYic";

//   // 延遲時間，取最新資料
//   // Utilities.sleep(3000);

//   let pdf_morning_h_id= "136eFZSExxjuk_GdgGgJUmXtdzLCMe9HZ"

//   let message = "~測試中~_學務處巡視" + "\n";
//   message+="------------------------------------"+ "\n"; 
//   message += Utilities.formatDate(new Date(),'GMT+8','MM/dd')+"_學務處午休巡視彙整表已上傳，請以瀛海中學gmail登入下方連結，點閱相關pdf檔" + "\n";
//   message += "https://drive.google.com/drive/folders/1wDmLPmoJdQjioHNiz2hoErvQ1gfypd-K" + "\n";
//   message+="------------------------------------"+ "\n"; 
    

//     Logger.log(message); 

//   // 延遲送出
//   // Utilities.sleep(8000);

//   // eric_temp
//   // sendLineNotify(message, token);
  
//   // sendLineNotify2(message, token2);
//   // sendLineNotify3(message, token3);
//   // sendLineNotify4(message, token4);

//   // sendFileToLine(pdf_morning_h_id,token)

// }


// function sendLineNotify(message, token){
//   let options =
//    {
//      "method"  : "post",
//      "payload" : {"message" : message},
//      "headers" : {"Authorization" : "Bearer " + token}
//    };
//    UrlFetchApp.fetch("https://notify-api.line.me/api/notify", options);
// }


// function sendNoonClassNoData(){

//   let sheetTabNameMorning = "012_daily_noon";
//   let ss = SpreadsheetApp.getActiveSpreadsheet();
//   let sh = ss.getSheetByName(sheetTabNameMorning);
//   let range=sh.getRange(3,1,46,16).getValues();
//   let datetime=getTodayDatetime();
//   let todayDate=datetime.slice(8,10);
  
//   let message='~學務巡視測試模擬 *周一至周五1500-1530* 自動發通知至學務相關群組提醒學務夥伴~'+ '\n' ;

//   let lineToken=[]
// // eric_temp
// let token = "i4npBEKuCBH5sUQ6LxuaOIDhsev3q5VzpnYs97wZW0u";
// // 2021學務網路測試_仁者無敵
// // let token2 = "gzyQg3S3QzZltehQKTMWl5BEkJ6iczB7kAynCJ0pDp3";
// // 2023學務網路測試_學務測試
// let token3 = "3EiMf2mx0LWiSU070a771Q3YMTcECV6uxbfQmBGvrSt";
// lineToken.push(token)
// lineToken.push(token3)

//   let noDataClass=range.filter(function(e){return e[16] ==''})
//   // console.log(function(e))
//   // console.log(noDataClass.length)

//   let grade=''

//   if (noDataClass.length<=0) {
//     message+='目前本('+ todayDate +')日各班午休巡視結果資料 *皆已入檔* ，感謝好夥伴的協助!!' + '\n';
  
//   }else{

//     message+='本('+ todayDate +')日下列班級目前 *午修* 巡視紀錄尚未入檔，煩請同仁協助處理，感謝您的幫忙!!' + '\n';
  
//     for (i=0;i<noDataClass.length;i++){

//       // 年級跳行
//       if (grade !=noDataClass[i][0].slice(0,2)){
//         message+= '\n'
//         message+= noDataClass[i][0].slice(0,2) + '\n'
//       }
      
//       message +=noDataClass[i][0] 
//       grade= noDataClass[i][0].slice(0,2)
      
//       // 防止最後一個值沒辦法還有下一個可以比較 
//       if (i>=noDataClass.length-1){
//         break;
//       }
    
//       // 結尾年級跳行
//       if(grade !=noDataClass[i+1][0].slice(0,2)) {
//         message+= '\n'
//       }else{
//         message+= ', '
//       }  
//     }
//   }
//   // console.log(message)
//   sendLineNotify(message,lineToken)  

// }
