//--------------------------------------------------
//実施想定タイミング：夜勤者が日付が変更されてから実行する
//--------------------------------------------------

//--------------------------------------------------
//　--調整箇所 ① & ②--
//--------------------------------------------------
//もしも勤務者全員が取得できていなかった場合は、ここを調整。
//ここは、当日のシートの取得範囲
//（仮に2020年10月1日夜勤者の場合は、2020年10月2日を取得）
//①を調整した場合、②もセットで調整する。
//①と同じ範囲以上の範囲を②も設定すること。

//--------------------------------------------------
//　--調整箇所 ③ & ④--
//--------------------------------------------------
//内容は① & ②と同じ


//昨日の年月日を取得
function makeLastDaySheetName() {
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
    }

    //上記２パターン以外だった場合は、前日を取得するためにマイナス１する
    if(day !== 1){
        day = day - 1;
    }

    month = ("0" + month).slice(-2);
    day = ("0" + day).slice(-2);
    var lastDaySheetName = year + month + day;
    Logger.log(lastDaySheetName);
    return lastDaySheetName;
}


//昨日の業務管理シートを取得
function getLastDaysShift() {
    var ss = SpreadsheetApp.openById("1DPQVKf7NFl2qq7xMoHlKzS1_zpVltJrhuFj4InU2xvk");
    var lastDaySheetName = makeLastDaySheetName();
    var lastDaySheet = ss.getSheetByName(lastDaySheetName);
    var lastdaysShiftArea = lastDaySheet.getRange("A7:B36");//③
    var lastdaysShiftValues = lastdaysShiftArea.getValues();
    return lastdaysShiftValues;
}


//当日の年月日を取得（例：10月1日夜勤なら「20201002」）
function makeTodaySheetName() {
    var date = new Date();
    var year = date.getFullYear();
    var month = date.getMonth() + 1;
    var day = date.getDate();

    month = ("0" + month).slice(-2);
    day = ("0" + day).slice(-2);
    var todaySheetName = year + month + day;
    return todaySheetName;
}


//当日の業務管理シートを取得
function getTodaysShift() {
    var ss = SpreadsheetApp.openById("1DPQVKf7NFl2qq7xMoHlKzS1_zpVltJrhuFj4InU2xvk");
    var todaySheetName = makeTodaySheetName();
    var todaySheet = ss.getSheetByName(todaySheetName);
    var todaysShiftArea = todaySheet.getRange("A7:B36");//①「"セル範囲"」を調整
    var todaysShiftValues = todaysShiftArea.getValues();
    return todaysShiftValues;
}


//昨日の業務管理シートのシフトをイベントシートに貼り付け（実行関数）
function setValues() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("データ取得用");
    
    var todaysShiftValues = getTodaysShift();
    var lastdaysShiftValues = getLastDaysShift();
    sheet.getRange("A1:B30").setValues(lastdaysShiftValues);//④：③の前日シフト出力場所「"セル範囲"」を調整
    sheet.getRange("C1:D30").setValues(todaysShiftValues);//②：①の当日シフト出力場所「"セル範囲"」を調整
}

//----------------制作中

//平日か祝日か土日か判定
function checkDate(){
    var date = new Date();
    var day = date.getDay();
    
    //土日だったらfalse
    if(day === 0 || day === 6){
        return false;
    }

    //祝日だったら
    var calJa = CalendarApp.getCalendarById('ja.japanese#holiday@group.v.calendar.google.com');
    if(calJa.getEventsForDay(day).length > 0){
        return false;
    }

    //平日


}


//マスタの取得（実行関数）
function getMasterValues(){
    var ss = SpreadsheetApp.openById("1DPQVKf7NFl2qq7xMoHlKzS1_zpVltJrhuFj4InU2xvk");
    
}
