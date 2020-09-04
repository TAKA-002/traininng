//①〜⑤を状況に応じて変更

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
    var files = DriveApp.getFolderById('1ekDBmDsLyo8gqxxRenS3neNhz6Sg0W91').getFiles();//★★★①複製先フォルダIDを入力★★★
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


//テンプレートシートから曜日祝日にあわせたモノを選択してコピーする
function createSheet(fileId, year, month) {
    var lastDay = findLastDays();
    var calJa = CalendarApp.getCalendarById('ja.japanese#holiday@group.v.calendar.google.com');
    var ss = SpreadsheetApp.openById(fileId);

    for (var i = 1; i <= lastDay; i++) {
        var day = ("00" + i).slice(-2);
        Logger.log(day);
        var date = new Date(`${year}\/${month}\/${day}`);

        if (date.getDay() == 0 || date.getDay() == 6) {
          //曜日が土日ならこちら
          var tempWESheet = ss.getSheetByName('土日テンプレート改');//★★★③土日テンプレートシート名★★★
          var copiedSheet = tempWESheet.copyTo(ss);
          copiedSheet.setName(year + month + day);
        }
//      if (calJa.getEventsForDay(date).length > 0) {
          //祝日ならこちら
//        var tempHDSheet =  ss.getSheetByName('祝日テンプレート改');//★★★④祝日テンプレートシート名★★★
//        var copiedSheet = tempHDSheet.copyTo(ss);
//        copiedSheet.setName(year + month + day);
//      }
        else{
          //平日ならこちら
          var tempWDSheet = ss.getSheetByName('平日テンプレート改');//★★★⑤平日テンプレートシート名★★★
          var copiedSheet = tempWDSheet.copyTo(ss);
          copiedSheet.setName(year + month + day);
       }
    }
}
