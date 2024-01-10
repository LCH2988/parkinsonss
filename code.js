// Code.gs

// 連結HTML檔案
function doGet() {
  var html = HtmlService.createTemplateFromFile("index6");
  var check = html.evaluate();
  var show = check.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return show;
}

// 取得試算表中 B2 和 H2 的值
function getItemDetails() {
  //var ss = SpreadsheetApp.getActiveSpreadsheet();
  //var sheet = ss.getSheetByName("#A_參數設定"); // 請替換成實際的工作表名稱
  var sheet = SpreadsheetApp.openById("1len_8cGvtTID7jlWo-6zCUnlf7mGnka1jgjMtXtGkGiHhh0hWOdjscS6").getSheetByName("#A_參數設定");

  var itemD2 = sheet.getRange("D2").getValue();
  var itemD3 = sheet.getRange("D3").getValue();
  var itemE3 = sheet.getRange("E3").getValue();
  var itemF3 = sheet.getRange("F3").getValue();
  var itemI3 = sheet.getRange("I3").getValue();
  var itemJ3 = sheet.getRange("J3").getValue();  
  return {
    itemD2: itemD2,//報名清單
    itemD3: itemD3,//標題
    itemE3: itemE3,//代號
    itemF3: itemF3,//項目名稱
    itemI3: itemI3,//建檔人
    itemJ3: itemJ3,//統計
  };
}

// 取得報名清單
function getRegistrationList() {
  //var ss = SpreadsheetApp.getActiveSpreadsheet();
  //var sheet = ss.getSheetByName("#B_活動報名清單"); // 請替換成實際的工作表名稱
  var sheet = SpreadsheetApp.openById("1len_8cGvtTID7jlWo-6zCUnlf7mGnka1jgjMtXtGkGiHhh0hWOdjscS6").getSheetByName("#B_活動報名清單");

  var data = sheet.getDataRange().getValues();

  var registrations = [];

  // 從第二列開始讀取資料
  for (var i = 2; i < data.length; i++) {
    //編號	代號	項目名稱	姓名	人數	備註	類別	主題	回覆代碼
    var registration = {
      k1: data[i][0],
      name: data[i][3], // 假設姓名在第三欄
      count: data[i][4], // 假設人數在第五欄
      remarks: data[i][5] // 假設備註在第六欄
    };

    registrations.push(registration);
  }

  return registrations;
}




function addData(rowData) {
  // 抓時間
  var currentDate = new Date();
  var //ss2 = SpreadsheetApp.getActiveSpreadsheet();
  var //sheet2 = ss2.getSheetByName("#A_參數設定"); // 請替換成實際的工作表名稱
  var sheet2 = SpreadsheetApp.openById("1len_8cGvtTID7jlWo-6zCUnlf7mGnka1jgjMtXtGkGiHhh0hWOdjscS6").getSheetByName("#A_參數設定");

  //編號	序號	類別	標題	代號	項目名稱	項目內容	建檔時間	建檔人
  var itemA3 = sheet2.getRange("A3").getValue();
  var itemB3 = sheet2.getRange("B3").getValue();
  var itemC3 = sheet2.getRange("C3").getValue();
  var itemD3 = sheet2.getRange("D3").getValue();
  var itemE3 = sheet2.getRange("E3").getValue();
  var itemF3 = sheet2.getRange("F3").getValue();
  var itemI3 = sheet2.getRange("I3").getValue();
  var itemJ3 = sheet2.getRange("J3").getValue();

  //var ss = SpreadsheetApp.getActiveSpreadsheet();
  // 抓試算表名稱
  // 請輸入您的試算表名稱在這裡
  //var ws = ss.getSheetByName("#B_活動報名清單");
  var ws = SpreadsheetApp.openById("1len_8cGvtTID7jlWo-6zCUnlf7mGnka1jgjMtXtGkGiHhh0hWOdjscS6").getSheetByName("#B_活動報名清單");

  //編號	代號	項目名稱	姓名	人數	備註	類別	主題	時間
  // 提取 rowData 中的值
  var name = rowData.name;
  var count = rowData.count;
  var phone = rowData.phone;
  var remarks = rowData.remarks;
  var Lrow = ws.getLastRow()-1;
  // 插入資料
  if( name !="" && count != "")
  ws.appendRow([Lrow, itemE3, itemF3, name, count, remarks, phone, "", currentDate]);
}

// 在 Code.gs 中添加查詢報名資料和修改報名資料的函數
// ...

function getRegistrationData(phoneToModify) {
  //var ss = SpreadsheetApp.getActiveSpreadsheet();
  //var sheet = ss.getSheetByName("#B_活動報名清單");
  var sheet = SpreadsheetApp.openById("1len_8cGvtTID7jlWo-6zCUnlf7mGnka1jgjMtXtGkGiHhh0hWOdjscS6").getSheetByName("#B_活動報名清單");

  var data = sheet.getDataRange().getValues();
  for (var i = 2; i < data.length; i++) {
    var registrationPhone = data[i][6]; // 假設電話在第七欄
    if (registrationPhone == phoneToModify) {
      return {
        name: data[i][3], // 假設姓名在第四欄
        count: data[i][4], // 假設人數在第五欄
        remarks: data[i][5] // 假設備註在第六欄
      };
    }
  }
  return null;
}

function modifyData(modifyData) {
  //var ss = SpreadsheetApp.getActiveSpreadsheet();
  //var sheet = ss.getSheetByName("#B_活動報名清單");
  var sheet = SpreadsheetApp.openById("1len_8cGvtTID7jlWo-6zCUnlf7mGnka1jgjMtXtGkGiHhh0hWOdjscS6").getSheetByName("#B_活動報名清單");

  var data = sheet.getDataRange().getValues();
  for (var i = 2; i < data.length; i++) {
    var registrationPhone = data[i][6]; // 假設電話在第七欄
    if (registrationPhone == modifyData.phone) {
      // 修改姓名、人數、備註
      data[i][3] = modifyData.name;
      data[i][4] = modifyData.count;
      data[i][5] = modifyData.remarks;
      sheet.getRange(i + 1, 1, 1, data[i].length).setValues([data[i]]);
      return;
    }
  }
}

