//①〜⑩を状況に応じて変更

//＜ 業務管理シートの変更箇所 ＞
//１：テンプレートの土日・祝日・平日のシートの年月日の表示を「yyyy年mm月dd日」へ変更する必要がある
//２：シフト表のスプレッドシートに常に作成したい月のシフト表を入れておく必要がある（シート名は汎用のものにすると変更しなくていい②⑤）


//対象の月の日数を調べる
function findLastDays() {
  var year = getMakingYear();
  var month = getMakingMonth();

  //シフト表の年月の最終日を取得
  var lastDay = new Date(parseInt(year, 10), parseInt(month, 10), 0).getDate();
  return lastDay;
}


//仮想シフト表から「月を取得」してmm形式でリターン
function getMakingMonth() {
  //仮想シフト表のIDを取得する
  const ss = SpreadsheetApp.openById('1DNxOMCa8u3qcZLVws9kvxh2TdLMr7rLdu4ffc7aLy0M');//★★★①シフト表のスプレッドシートのIDを入力★★★

  //取得したスプレッドシートのシフト表のシートを取得する
  const sheet = ss.getSheetByName('10月シフト');//★★★②シフト表のシート名を入力★★★

  //シフト表のシートの作成月の記載されているセルを取得する
  const range = sheet.getRange('B2').getValue();//★★★③シフト表の年月日を入力しているセルを入力★★★

  //Dateオブジェクトで指定したセルの値をインスタンス化する
  var shiftDate = new Date(range);

  //セルの月を２桁に統一する
  var month = ("0" + (shiftDate.getMonth() + 1)).slice(-2);
  return month;
}


//仮想シフト表から「年」を取得を取得してyyyy形式でリターン
function getMakingYear() {
  const ss = SpreadsheetApp.openById('1DNxOMCa8u3qcZLVws9kvxh2TdLMr7rLdu4ffc7aLy0M');//★★★④シフト表のIDを入力（①と同じ）★★★
  const sheet = ss.getSheetByName('10月シフト');//★★★⑤シフト表のシート名を入力（②と同じ）★★★
  const range = sheet.getRange('B2').getValue();//★★★⑥シフト表の年月日を入力しているセルを入力（③と同じ）★★★
  var shiftDate = new Date(range);
  var year = shiftDate.getFullYear();
  return year;
}


//「平日テンプレート改」のシートのインデックス番号を取得
function getTargetSheet() {
  // 現在アクティブなスプレッドシートを取得し、
  // そのスプレッドシートにある"平日テンプレート改"という名前のシートを取得（このシートの左に作成するため）
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("平日テンプレート改");//★★★⑦どのシートの左に生成したいか。シート名を記述★★★
  if (sheet != null) {
    // そのシートのインデックス番号を取得
    var targetSheetIndex = sheet.getIndex();
    return targetSheetIndex;
  }
}


//曜日祝日にあわせたモノを選択してコピー。シート見出しも併せて変更。
function createSheet() {
  var year = getMakingYear();
  var month = getMakingMonth();
  var lastDay = findLastDays();
  var ss = SpreadsheetApp.getActiveSpreadsheet();//開いているシートを取得
  var calJa = CalendarApp.getCalendarById('ja.japanese#holiday@group.v.calendar.google.com');

  for (var i = 1; i <= lastDay; i++) {
    var day = ("00" + i).slice(-2);
    var date = new Date(`${year}\/${month}\/${day}`);

    if (date.getDay() == 0 || date.getDay() == 6) {
      //土日用シート複製
      var tempWESheet = ss.getSheetByName('土日テンプレート改');//★★★⑧土日用のテンプレート名を記述★★★
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
      var tempHDSheet = ss.getSheetByName('祝日テンプレート改');//★★★⑨祝日用のテンプレート名を記述★★★
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
      var tempWDSheet = ss.getSheetByName('平日テンプレート改');//★★★⑩平日用のテンプレート名を記述★★★
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
