//対象の月の日数を調べる
function findLastDays() {
    //シフト表の年を取得
    var year = getMakingYear();

    //シフト表の月を取得
    var month = getMakingMonth();

    //シフト表の年月の最終日を取得
    var lastDay = new Date(parseInt(year, 10), parseInt(month, 10), 0).getDate();
    return lastDay;
}


function getDuplicatedFile(){
    //シフト表の年を取得
    var year = getMakingYear();

    //シフト表の月を取得
    var month = getMakingMonth();

    //ドライブの中の複製先のフォルダを指定し、その中のファイルを取得
    var files = DriveApp.getFolderById('19uX3BktgUEpZyNXd4COca31Mrc8iyEKt').getFiles();
    var fileName = '業務管理シート_' + year + month;

    while(files.hasNext()){
      var file = files.next();//戻り値はFile型
      var duplicatedFileName = file.getName();
      var fileId = file.getId();
      Logger.log(fileId);

      if(duplicatedFileName === fileName){
      createSheet(fileId, year, month);
      break;
      }
    }
}


function createSheet(fileId, year, month) {
    var lastDay = findLastDays();
    var calJa = CalendarApp.getCalendarById('ja.japanese#holiday@group.v.calendar.google.com');
    var ss = SpreadsheetApp.openById(fileId);

    for (var i = 1; i <= lastDay; i++) {
        var day = ("00" + i).slice(-2);
        Logger.log(day);
        var date = new Date(`${year}\/${month}\/${day}`);

        if (date.getDay() == 0 || date.getDay() == 6) {
            Logger.log(date.getDay());
            //スプレッドシートをアクティブにする
            var tempWESheet = ss.getSheetByName('土日テンプレート改');
            //土日のシートを取得して複製する
            var copiedSheet = tempWESheet.copyTo(ss);
            copiedSheet.setName(year + month + day);
        }
//        if (calJa.getEventsForDay(date).length > 0) {
//            var tempHDSheet =  ss.getSheetByName('祝日テンプレート改');
//            var copiedSheet = tempHDSheet.copyTo(ss);
//            copiedSheet.setName(year + month + day);
//        }
        else{
            var tempWDSheet = ss.getSheetByName('平日テンプレート改');
            var copiedSheet = tempWDSheet.copyTo(ss);
            copiedSheet.setName(year + month + day);
        }
    }
}
