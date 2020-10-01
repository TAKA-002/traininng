//--------------------------------------------------
//実施想定タイミング：夜勤者が日付が変更されてから実行する
//--------------------------------------------------

//--------------------------------------------------
//--調整箇所 ① & ②--
//--------------------------------------------------
//もしも勤務者全員が取得できていなかった場合は、ここを調整。
//ここは、当日のシートの取得範囲
//（仮に2020年10月1日夜勤者の場合は、2020年10月2日を取得）
//①を調整した場合、②もセットで調整する。
//①と同じ範囲以上の範囲を②も設定すること。

//--------------------------------------------------
//--調整箇所 ③ & ④--
//--------------------------------------------------
//内容は① & ②と同じ



//当日の年月日を取得
function getTodaySheetName() {
    var date = new Date();
    var year = date.getFullYear();
    var month = date.getMonth() + 1;
    month = ("0" + month).slice(-2);
    var day = date.getDate();
    day = ("0" + day).slice(-2);
    var todaySheetName = year + month + day;
    return todaySheetName;
}


//当日の業務管理シートを取得
function getTodaysShift() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var todaySheetName = getTodaySheetName();
    var todaySheet = ss.getSheetByName(todaySheetName);
    var todaysShiftArea = todaySheet.getRange("A7:B38");//①「"A7:B38"」の箇所を調整
    var todaysShiftValues = todaysShiftArea.getValues();
    Logger.log(todaysShiftValues);
    return todaysShiftValues;
}


//昨日の年月日を取得
function getLastDaySheetName() {
    var date = new Date();
    var year = date.getFullYear();
    var month = date.getMonth() + 1;
    var day = date.getDate();

    //もしも元日だった場合、前年大晦日にする
    if (day === 1 && month === 1) {
        year = year - 1;
        month = 12;
        day = 31;
    }

    //もしも月初一日だった場合、前月にして最終日を取得する
    if (day === 1) {
        month = month - 1;
        var dt = new Date(year, month, 0);
        day = dt.getDate();
        Logger.log(day);
    }

    month = ("0" + month).slice(-2);
    day = ("0" + day).slice(-2);
    var lastDaySheetName = year + month + day;
    return lastDaySheetName;
}


//昨日の業務管理シートを取得
function getLastDaysShift() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var lastDaySheetName = getLastDaySheetName();
    var lastDaySheet = ss.getSheetByName(lastDaySheetName);
    var lastdaysShiftArea = lastDaySheet.getRange("A7:B38");//③
    var lastdaysShiftValues = lastdaysShiftArea.getValues();
    Logger.log(lastdaysShiftValues);
    return lastdaysShiftValues;
}


//昨日の業務管理シートのシフトをイベントシートに貼り付け（実行関数）
function setValues() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("イベントシート");

    var todaysShiftValues = getTodaysShift();
    var lastdaysShiftValues = getLastDaysShift();
    sheet.getRange("C35:D66").setValues(lastdaysShiftValues);//④
    sheet.getRange("F35:G66").setValues(todaysShiftValues);//②「"F35:G66"」の箇所を調整
}