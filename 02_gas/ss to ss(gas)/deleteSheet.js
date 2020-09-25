//------▼▼コードの目的▼▼------
//イベントボタンを押下し、不要なシートを削除する

//------▼▼コード調整箇所▼▼------
//特になし


//ボタンを押した年月を取得して、その1ヶ月前を「yyyymm」の6桁で取得する
function checkDate() {
  var date = new Date();
  var targetYear = date.getFullYear();
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


//-----運用中「業務管理シート」で不要な非表示シートを削除する-----
//削除条件：「非表示」「シート名が前月」

//-----処理-----
//１：開いているスプレッドシートを取得
//２：シート数をカウントする（シートの数だけ以下フローを繰り返す）
//３：ボタンを押した月の前月「yyyymm」をtargetに格納
//４：開いているスプレッドシートのシートインデックス[i]を取得
//取得したシートが存在していたら、
//５：そのシートのシート名を取得する
//６：「もしも取得したシートのシート名にyyyymmが含まれていた」「そのシートが非表示のシートである」→ そのシートを削除する
//７：削除した場合、シート数が減るためループする際にインデックスがずれてしまう。なので、マイナス１する
//８：削除しない場合はその回のループは飛ばす
//９：もしもシートインデックス[i]が存在しない場合、ループを終了する

function deleteHiddenSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetCount = ss.getNumSheets();
  var target = checkDate();

  for (var i = 0; i <= sheetCount; i++) {
    var sheet = ss.getSheets()[i];
    if (sheet != null) {
      var sheetName = sheet.getSheetName();
      if (sheetName.match(target) && sheet.isSheetHidden() === true) {
        ss.deleteSheet(sheet);
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


//-----バックアップ用スプレッドシートの体裁を調整する-----
//削除条件：「表示」「シート名が半角数字8桁」

//-----処理-----
//１：現在開いているスプレッドシートを取得
//２：シート数をカウントする（シートの数だけ以下ループを繰り返す）
//３：開いているシートのシートインデックス番号[i]をsheetに格納
//４：もしもシートが存在していたら、そのシートのシート名を取得してsheetNameに格納
//５：「シート名が半角数字８桁」「シートが非表示ではない」場合は、そのシートを削除する
//６：削除してインデックス番号がずれるので、次のインデックス番号を同じにするためiを1マイナスする
//７：もしも５で条件が合わなかったら、そのループは飛ばす
//８:もしも４でシートが存在していない場合、ループを終了する

function deleteShowSheets() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetCount = ss.getNumSheets();

  for (var i = 0; i <= sheetCount; i++) {
    var sheet = ss.getSheets()[i];
    if (sheet != null) {
      var sheetName = sheet.getSheetName();
      if (sheetName.match(/[\d]{8}/) && sheet.isSheetHidden() === false) {
        ss.deleteSheet(sheet);
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


//-----バックアップ用スプレッドシートの体裁を調整する-----
//deleteShowSheetsで、不要な表示シートを削除したあと、必要な非表示のシートを表示にする

function showSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetCount = ss.getNumSheets();

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