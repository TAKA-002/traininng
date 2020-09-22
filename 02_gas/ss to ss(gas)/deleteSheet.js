//ボタンを押した月を取得して、その1ヶ月前を取得する
function checkDate() {
  var date = new Date();
  var targetYear = date.getFullYear();
  //セルの月を２桁に統一する
  var Month = ("0" + (date.getMonth() + 1)).slice(-2);

  switch (Month) {
    case "01":
      targetYear = targetYear - 1;
      var targetMonth = "12";
      var target = targetYear + targetMonth;
      return target;
      break;
    case "02":
      var targetMonth = "01";
      var target = targetYear + targetMonth;
      return target;
      break;
    case "03":
      var targetMonth = "02";
      var target = targetYear + targetMonth;
      return target;
      break;
    case "04":
      var targetMonth = "03";
      var target = targetYear + targetMonth;
      return target;
      break;
    case "05":
      var targetMonth = "04";
      var target = targetYear + targetMonth;
      return target;
      break;
    case "06":
      var targetMonth = "05";
      var target = targetYear + targetMonth;
      return target;
      break;
    case "07":
      var targetMonth = "06";
      var target = targetYear + targetMonth;
      return target;
      break;
    case "08":
      var targetMonth = "07";
      var target = targetYear + targetMonth;
      return target;
      break;
    case "09":
      var targetMonth = "08";
      var target = targetYear + targetMonth;
      return target;
      break;
    case "10":
      var targetMonth = "09";
      var target = targetYear + targetMonth;
      return target;
      break;
    case "11":
      var targetMonth = "10";
      var target = targetYear + targetMonth;
      return target;
      break;
    case "12":
      var targetMonth = "11";
      var target = targetYear + targetMonth;
      return target;
      break;
  }
}


//削除されるシートの条件：非表示、シート名が先月のもの
function deleteHiddenSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();//開いているシートを取得
  const sheetCount = ss.getNumSheets();//シート数をカウント
  var target = checkDate();

  for (var i = 0; i <= sheetCount; i++) {//シートの数だけ繰り返す
    var sheet = ss.getSheets()[i];
    if (sheet != null) {
      var sheetName = sheet.getSheetName();//シート名を取得する
      if (sheetName.match(target) && sheet.isSheetHidden() === true) {//シート名が先月のyyyymmがふくまれていて、かつ「非表示」である場合
        ss.deleteSheet(sheet);//削除する
        i--;//削除したらiをマイナス１してひとつ飛ばさないようにする
      } else {
        continue;
      }
    }
    if (sheet == null) {
      break;
    }
  }
}


//削除されるシートの条件：表示、シート名が半角数字8桁
function deleteShowSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();//開いているシートを取得
  const sheetCount = ss.getNumSheets();//シート数をカウント

  for (var i = 0; i <= sheetCount; i++) {//シートの数だけ繰り返す
    var sheet = ss.getSheets()[i];
    if (sheet != null) {
      var sheetName = sheet.getSheetName();//シート名を取得する
      if (sheetName.match(/[\d]{8}/) && sheet.isSheetHidden() === false) {//もしもシート名が半角数字8桁で、かつ「表示」である場合
        ss.deleteSheet(sheet);//削除する
        i--;
      } else {
        continue;
      }
    }
    if (sheet == null) {
      break;
    }
  }
}


function showSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();//開いているシートを取得
  const sheetCount = ss.getNumSheets();//シート数をカウント

  for (var i = 0; i <= sheetCount; i++) {
    var sheet = ss.getSheets()[i];
    if (sheet != null && sheet.isSheetHidden() === true) {
      sheet.showSheet();
    }
  }
}


function openDialogBoxforBackup() {
  var result = Browser.msgBox("注意：バックアップ用のファイルで実行してください！！実行してよろしいですか？", Browser.Buttons.OK_CANCEL);
  if (result == "ok") {
    deleteShowSheets();
    showSheet();
  } //OKの処置
  if (result == "cancel") {
    return;
  } //Cancelの処置
}


function openDialogBoxforActive() {
  var result = Browser.msgBox("注意：非表示の先月のシートが削除されます！！実行してよろしいですか？", Browser.Buttons.OK_CANCEL);
  if (result == "ok") {
    deleteHiddenSheets();
  } //OKの処置
  if (result == "cancel") {
    return;
  } //Cancelの処置
}
