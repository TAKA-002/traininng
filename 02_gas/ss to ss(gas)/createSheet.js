//①〜⑤を状況に応じて変更

//＜ 業務管理シートの変更箇所 ＞
//１：テンプレートの土日・祝日・平日のシートの年月日の表示を「yyyy年mm月dd日」へ変更する必要がある

//対象の月の日数を調べる
function findLastDays() {
    var year = getMakingYear();
    var month = getMakingMonth();

    //シフト表の年月の最終日を取得
    var lastDay = new Date(parseInt(year, 10), parseInt(month, 10), 0).getDate();
    return lastDay;
}


//複製したファイルを取得してcreateSheet関数を実行
function getDuplicatedFile(){
    var year = getMakingYear();
    var month = getMakingMonth();

    //ドライブの中の複製先のフォルダを指定し、その中のファイルを取得
    var files = DriveApp.getFolderById('19uX3BktgUEpZyNXd4COca31Mrc8iyEKt').getFiles();//★★★①複製先フォルダIDを入力★★★
    var fileName = '業務管理シート_' + year + month;//★★★②新ファイル（複製したファイル）名★★★

    while(files.hasNext()){
      var file = files.next();//戻り値はFile型
      var duplicatedFileName = file.getName();
      var fileId = file.getId();
      if(duplicatedFileName === fileName){
      createSheet(fileId, year, month);
      break;
      }
    }
}


//テンプレートシートから曜日祝日にあわせたモノを選択してコピー。シート見出しも併せて変更。
function createSheet(fileId, year, month) {
    var lastDay = findLastDays();
    var ss = SpreadsheetApp.openById(fileId);
    var calJa = CalendarApp.getCalendarById('ja.japanese#holiday@group.v.calendar.google.com');

    for (var i = lastDay; i >= 1; i--) {
        var day = ("00" + i).slice(-2);
        var date = new Date(`${year}\/${month}\/${day}`);

        if (date.getDay() == 0 || date.getDay() == 6) {
          //土日用シート複製
          var tempWESheet = ss.getSheetByName('土日テンプレート改');//★★★③土日テンプレートシート名★★★
          var copiedSheet = tempWESheet.copyTo(ss);
          copiedSheet.setName(year + month + day);
          
          //シートの見出しを該当する年月日へ変更
          var textFinder = copiedSheet.createTextFinder('yyyy年mm月dd日');
          textFinder.replaceAllWith(`${year}年${month}月${day}日`);
          
          //作成したシートの位置を左から2番目に移動
          ss.setActiveSheet(copiedSheet);
          ss.moveActiveSheet(2);
        } else if (calJa.getEventsForDay(date).length > 0) {
          //祝日用シート複製
          var tempHDSheet =  ss.getSheetByName('祝日テンプレート改');//★★★④祝日テンプレートシート名★★★
          var copiedSheet = tempHDSheet.copyTo(ss);
          
          //シートの見出しを該当する年月日へ変更
          copiedSheet.setName(year + month + day);
          var textFinder = copiedSheet.createTextFinder('yyyy年mm月dd日');
          textFinder.replaceAllWith(`${year}年${month}月${day}日`);

          //作成したシートの位置を左から2番目に移動
          ss.setActiveSheet(copiedSheet);
          ss.moveActiveSheet(2);
        }else{
          //平日用シート複製
          var tempWDSheet = ss.getSheetByName('平日テンプレート改');//★★★⑤平日テンプレートシート名★★★
          var copiedSheet = tempWDSheet.copyTo(ss);
          
　　　　　　//シートの見出しを該当する年月日へ変更
          copiedSheet.setName(year + month + day);
          var textFinder = copiedSheet.createTextFinder('yyyy年mm月dd日');
          textFinder.replaceAllWith(`${year}年${month}月${day}日`);

          //作成したシートの位置を左から2番目に移動
          ss.setActiveSheet(copiedSheet);
          ss.moveActiveSheet(2);
       }
    }
}//①〜⑤を状況に応じて変更

//＜ 業務管理シートの変更箇所 ＞
//１：テンプレートの土日・祝日・平日のシートの年月日の表示を「yyyy年mm月dd日」へ変更する必要がある

//対象の月の日数を調べる
function findLastDays() {
    var year = getMakingYear();
    var month = getMakingMonth();

    //シフト表の年月の最終日を取得
    var lastDay = new Date(parseInt(year, 10), parseInt(month, 10), 0).getDate();
    return lastDay;
}


//複製したファイルを取得してcreateSheet関数を実行
function getDuplicatedFile(){
    var year = getMakingYear();
    var month = getMakingMonth();

    //ドライブの中の複製先のフォルダを指定し、その中のファイルを取得
    var files = DriveApp.getFolderById('19uX3BktgUEpZyNXd4COca31Mrc8iyEKt').getFiles();//★★★①複製先フォルダIDを入力★★★
    var fileName = '業務管理シート_' + year + month;//★★★②新ファイル（複製したファイル）名★★★

    while(files.hasNext()){
      var file = files.next();//戻り値はFile型
      var duplicatedFileName = file.getName();
      var fileId = file.getId();
      if(duplicatedFileName === fileName){
      createSheet(fileId, year, month);
      break;
      }
    }
}


//テンプレートシートから曜日祝日にあわせたモノを選択してコピー。シート見出しも併せて変更。
function createSheet(fileId, year, month) {
    var lastDay = findLastDays();
    var ss = SpreadsheetApp.openById(fileId);
    var calJa = CalendarApp.getCalendarById('ja.japanese#holiday@group.v.calendar.google.com');

    for (var i = lastDay; i >= 1; i--) {
        var day = ("00" + i).slice(-2);
        var date = new Date(`${year}\/${month}\/${day}`);

        if (date.getDay() == 0 || date.getDay() == 6) {
          //土日用シート複製
          var tempWESheet = ss.getSheetByName('土日テンプレート改');//★★★③土日テンプレートシート名★★★
          var copiedSheet = tempWESheet.copyTo(ss);
          copiedSheet.setName(year + month + day);
          
          //シートの見出しを該当する年月日へ変更
          var textFinder = copiedSheet.createTextFinder('yyyy年mm月dd日');
          textFinder.replaceAllWith(`${year}年${month}月${day}日`);
          
          //作成したシートの位置を左から2番目に移動
          ss.setActiveSheet(copiedSheet);
          ss.moveActiveSheet(2);
        } else if (calJa.getEventsForDay(date).length > 0) {
          //祝日用シート複製
          var tempHDSheet =  ss.getSheetByName('祝日テンプレート改');//★★★④祝日テンプレートシート名★★★
          var copiedSheet = tempHDSheet.copyTo(ss);
          
          //シートの見出しを該当する年月日へ変更
          copiedSheet.setName(year + month + day);
          var textFinder = copiedSheet.createTextFinder('yyyy年mm月dd日');
          textFinder.replaceAllWith(`${year}年${month}月${day}日`);

          //作成したシートの位置を左から2番目に移動
          ss.setActiveSheet(copiedSheet);
          ss.moveActiveSheet(2);
        }else{
          //平日用シート複製
          var tempWDSheet = ss.getSheetByName('平日テンプレート改');//★★★⑤平日テンプレートシート名★★★
          var copiedSheet = tempWDSheet.copyTo(ss);
          
　　　　　　//シートの見出しを該当する年月日へ変更
          copiedSheet.setName(year + month + day);
          var textFinder = copiedSheet.createTextFinder('yyyy年mm月dd日');
          textFinder.replaceAllWith(`${year}年${month}月${day}日`);

          //作成したシートの位置を左から2番目に移動
          ss.setActiveSheet(copiedSheet);
          ss.moveActiveSheet(2);
       }
    }
}
