// Code.gs

// 連結HTML檔案
function doGet() {
  var html = HtmlService.createTemplateFromFile("index_test2");
  var check = html.evaluate();
  var show = check.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  return show;
}


function getItemDetails() {
  var sheetA = SpreadsheetApp.getActiveSpreadsheet.getSheetByName("#A_參數設定");
  var sheetB = SpreadsheetApp.getActiveSpreadsheets.getSheetByName("#B_活動報名清單");

  var itemDetailsA = ['A1', 'B1', 'C1', 'D1', 'E1', 'F1', 'G1', 'H1', 'I1', 'J1', 'K1', 'L1', 'M1', 'N1', 'O1', 'P1', 'A2', 'B2', 'C2', 'D2', 'E2', 'F2', 'G2', 'H2', 'I2', 'J2', 'K2', 'L2', 'M2', 'N2', 'O2', 'P2'].reduce((details, cell) => {
    details[cell] = sheetA.getRange(cell).getValue();
    return details;
  }, {});

  var itemDetailsB = ['A1', 'B1', 'C1', 'D1', 'E1', 'F1', 'G1', 'H1', 'I1', 'J1', 'K1', 'L1', 'M1', 'N1', 'O1', 'P1', 'A2', 'B2', 'C2', 'D2', 'E2', 'F2', 'G2', 'H2', 'I2', 'J2', 'K2', 'L2', 'M2', 'N2', 'O2', 'P2'].reduce((details, cell) => {
    details[cell] = sheetB.getRange(cell).getValue();
    return details;
  }, {});

  return { itemDetailsA, ...itemDetailsB };
}
