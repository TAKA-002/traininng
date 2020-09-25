//------▼▼コードの目的▼▼------
//シフト表のシートの年月からgoogleカレンダーに基づいてシートを作成する

//------▼▼コード調整箇所▼▼------
//①シフト表のスプレッドシートのIDを記述
//②シフト表のシート名を記述
//③シフト表の年月日を入力しているセルを記述
//④シフト表のスプレッドシートのIDを記述（①と同じもの）
//⑤シフト表のシート名を記述（②と同じもの）
//⑥シフト表の年月日を入力しているセルを記述（③と同じ）
//⑦生成したいシートのシート名を記述（シート名の左に生成）
//⑧土日用のテンプレート名を記述
//⑨祝日用のテンプレート名を記述
//⑩平日用のテンプレート名を記述

//-----▼▼業務管理シートの変更箇所▼▼-----
//１：テンプレートの土日・祝日・平日のシートの年月日の表示を「yyyy年mm月dd日」へ変更する必要がある
//２：シフト表のスプレッドシートに常に作成したい月のシフト表を入れておく必要がある（シート名は汎用のものにすると変更しなくていい②⑤）


//-----シートを作成する月の日数を調べる-----
function findLastDays() {
  var year = getMakingYear();
  var month = getMakingMonth();

  var lastDay = new Date(parseInt(year, 10), parseInt(month, 10), 0).getDate();
  return lastDay;
}


//-----シフト表のスプレッドシートから「月」を取得-----
function getMakingMonth() {
  const ss = SpreadsheetApp.openById('1DNxOMCa8u3qcZLVws9kvxh2TdLMr7rLdu4ffc7aLy0M'); //①
  const sheet = ss.getSheetByName('10月シフト'); //②
  const range = sheet.getRange('B2').getValue(); //③
  var shiftDate = new Date(range);

  var month = ("0" + (shiftDate.getMonth() + 1)).slice(-2);
  return month;
}


//-----シフト表のスプレッドシートから「年」を取得-----
function getMakingYear() {
  const ss = SpreadsheetApp.openById('1DNxOMCa8u3qcZLVws9kvxh2TdLMr7rLdu4ffc7aLy0M'); //④
  const sheet = ss.getSheetByName('10月シフト'); //⑤
  const range = sheet.getRange('B2').getValue(); //⑥
  var shiftDate = new Date(range);
  var year = shiftDate.getFullYear();
  return year;
}


//-----業務管理シートのスプレッドシートで新シートを作成する場所を指定-----
// 現在アクティブなスプレッドシートを取得し、そのスプレッドシートにある"平日テンプレート改"という名前のシートを取得（このシートの左に作成するため）
function getTargetSheet() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("平日テンプレート改"); //⑦

  if (sheet != null) {
    // そのシートのインデックス番号を取得
    var targetSheetIndex = sheet.getIndex();
    return targetSheetIndex;
  }
}


//-----曜日祝日に合わせてテンプレートを選択し、シートを複製（複製時にリネーム）-----
function createSheet() {
  var year = getMakingYear();
  var month = getMakingMonth();
  var lastDay = findLastDays();
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var calJa = CalendarApp.getCalendarById('ja.japanese#holiday@group.v.calendar.google.com');

  for (var i = 1; i <= lastDay; i++) {
    var day = ("00" + i).slice(-2);
    var date = new Date(`${year}\/${month}\/${day}`);

    if (date.getDay() == 0 || date.getDay() == 6) {
      //土日用シート複製
      var tempWESheet = ss.getSheetByName('土日テンプレート改'); //⑧
      var copiedSheet = tempWESheet.copyTo(ss);
      copiedSheet.setName(year + month + day);

      //シートの見出しを該当する年月日へ変更
      var textFinder = copiedSheet.createTextFinder('yyyy年mm月dd日');
      textFinder.replaceAllWith(`${year}年${month}月${day}日`);

      //作成したシートの位置を取得したインデックス番号に移動（そのインデックス番号になるようにする）
      ss.setActiveSheet(copiedSheet);
      var targetSheet = getTargetSheet();
      ss.moveActiveSheet(targetSheet);

    } else if (calJa.getEventsForDay(date).length > 0) {
      //祝日用シート複製
      var tempHDSheet = ss.getSheetByName('祝日テンプレート改'); //⑨
      var copiedSheet = tempHDSheet.copyTo(ss);

      //シートの見出しを該当する年月日へ変更
      copiedSheet.setName(year + month + day);
      var textFinder = copiedSheet.createTextFinder('yyyy年mm月dd日');
      textFinder.replaceAllWith(`${year}年${month}月${day}日`);

      //作成したシートの位置を取得したインデックス番号に移動（そのインデックス番号になるようにする）
      ss.setActiveSheet(copiedSheet);
      var targetSheet = getTargetSheet();
      ss.moveActiveSheet(targetSheet);

    } else {
      //平日用シート複製
      var tempWDSheet = ss.getSheetByName('平日テンプレート改'); //⑩
      var copiedSheet = tempWDSheet.copyTo(ss);

      //シートの見出しを該当する年月日へ変更
      copiedSheet.setName(year + month + day);
      var textFinder = copiedSheet.createTextFinder('yyyy年mm月dd日');
      textFinder.replaceAllWith(`${year}年${month}月${day}日`);

      //作成したシートの位置を取得したインデックス番号に移動（そのインデックス番号になるようにする）
      ss.setActiveSheet(copiedSheet);
      var targetSheet = getTargetSheet();
      ss.moveActiveSheet(targetSheet);
    }
  }
}