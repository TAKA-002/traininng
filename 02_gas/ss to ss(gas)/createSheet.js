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

function createSheet() {
    //シフト表の年を取得
    var year = getMakingYear();

    //シフト表の月を取得
    var month = getMakingMonth();

    //ドライブの中の複製先のフォルダを指定し、その中のファイルを取得
    var duplicatedSS = DriveApp.getFolderById('1ekDBmDsLyo8gqxxRenS3neNhz6Sg0W91').getSheet('業務管理シート_'+ year + month);
    console.log(duplicatedSS);

    var lastDay = findLastDays();
    var calJa = CalendarApp.getCalendarById('ja.japanese#holiday@group.v.calendar.google.com');

    for (var i = 1; i <= lastDay; i++) {
        var day = ("00" + i).slice(-2);
        var date = new Date(`${year}\/${month}\/${day}`);
        console.log(day);
        console.log(date);
        if (date.getDay() == 0 || date.getDay() == 6) {
            var tempWESheet = duplicatedSS.getSheetByName('土日テンプレート（新コロ）');
            var copiedSheet = tempWESheet.copyTo(duplicatedSS);
            //土日のシートを取得して複製する
            copiedSheet.setName(year + month + i);
        } else if (calJa.getEventsForDay(date).length > 0) {
            return;
            //祝日のシートを取得して複製する
        } else {
            return;
            //平日のシートを取得して複製する
        }
    }
}