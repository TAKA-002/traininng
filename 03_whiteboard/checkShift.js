//--------------------------------------------------

//■実施想定

//＜夜勤者がボタンを押したタイミングが日が変わる前の場合＞
//当日のシフトデータと、翌日のシフトデータを取得

//＜夜勤者がボタンを押したタイミングが日が変わった後の場合＞
//前日のシフトデータと、当日のシフトデータを取得

//--------------------------------------------------

//--------------------------------------------------
//　--調整箇所 ① & ②--
//--------------------------------------------------
//もしも勤務者全員が取得できていなかった場合は、ここを調整。
//ここは、当日のシートの取得範囲
//（仮に2020年10月1日夜勤者の場合は、2020年10月2日を取得）
//①を調整した場合、②もセットで調整する。
//①と同じ範囲以上の範囲を②も設定すること。


//===============================
//　--調整箇所--
//===============================

const MemberCount = 31;//デジステメンバー数が変わったら変更すること
const SheetName = 'シフトデータ確認表';//作成シート
const ManagementSheetID = '1DPQVKf7NFl2qq7xMoHlKzS1_zpVltJrhuFj4InU2xvk';//業務管理シートのID

//===============================
//もしもボタンをおしたタイミングが0時から7時の場合は、「前日のシフト」と「当日のシフト」を取得して貼り付け
//もしもボタンをおしたタイミングが21時から24時の場合は、「当日のシフト」と「翌日のシフト」を取得して貼り付け
//===============================

//現在時刻を取得
function getTime(){
    var now = new Date();
    var Hour = now.getHours();
    return Hour;
}


//昨日の業務管理シートのシフトをイベントシートに貼り付け
function setValues(beforeValues, afterValues) {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SheetName);
    
    sheet.getRange(1, 1, MemberCount, 2).setValues(beforeValues);
    sheet.getRange(1, 4, MemberCount, 2).setValues(afterValues);
}


//実行関数：実行時間で処理を分岐
function setTwoValues(){
    var Hour = getTime();
    if(Hour >= 21 && Hour <= 23){
        var todaysShiftValues = getTodaysShift();
        var tomorrowShiftValues = getTomorrowShift();
        setValues(todaysShiftValues,tomorrowShiftValues);
    }
    if(Hour >= 0 && Hour <= 6){
        var lastdaysShiftValues = getLastDaysShift();
        var todaysShiftValues = getTodaysShift();
        setValues(lastdaysShiftValues,todaysShiftValues);
    }
}


//===============================
//　前日
//===============================

//前日の年月日を取得
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
    return lastDaySheetName;
}


//前日の業務管理シートを取得
function getLastDaysShift() {
    var ss = SpreadsheetApp.openById(ManagementSheetID);
    var lastDaySheetName = makeLastDaySheetName();
    var lastDaySheet = ss.getSheetByName(lastDaySheetName);
    var lastdaysShiftArea = lastDaySheet.getRange(7, 1, MemberCount, 2);
    var lastdaysShiftValues = lastdaysShiftArea.getValues();
    return lastdaysShiftValues;
}


//===============================
//　当日
//===============================

//当日の年月日を取得
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
    var ss = SpreadsheetApp.openById(ManagementSheetID);
    var todaySheetName = makeTodaySheetName();
    var todaySheet = ss.getSheetByName(todaySheetName);
    var todaysShiftArea = todaySheet.getRange(7, 1, MemberCount, 2);
    var todaysShiftValues = todaysShiftArea.getValues();
    return todaysShiftValues;
}


//===============================
//　翌日
//===============================

//翌日の年月日を取得
function makeTomorrowSheetName() {
    var date = new Date();
    var year = date.getFullYear();
    var month = date.getMonth() + 1;
    var day = date.getDate();

    var dt = new Date(year, month, 0);
    var lastDay = dt.getDate();

    //もしも大晦日にボタンを押した場合、年を１プラスする
    if(day === 31 || month === 12){
        year = year + 1;
        month = 1;
        day = 1;
    }

    //もしも月末にボタンを押した場合、翌月にして日を月初にする
    if(lastDay === day){
        month = month + 1;
        day = 1;
    }

    //月末以外だった場合は、取得した日数に１プラスする
    if(day !== lastDay){
        day = day + 1;
    }
    
    month = ("0" + month).slice(-2);
    day = ("0" + day).slice(-2);
    var tomorrowSheetName = year + month + day;
    return tomorrowSheetName;
}

//翌日の業務管理シートを取得
function getTomorrowShift() {
    var ss = SpreadsheetApp.openById(ManagementSheetID);
    var tomorrowSheetName = makeTomorrowSheetName();
    var tomorrowSheet = ss.getSheetByName(tomorrowSheetName);
    var tomorrowShiftArea = tomorrowSheet.getRange(7, 1, MemberCount, 2);
    var tomorrowShiftValues = tomorrowShiftArea.getValues();
    return tomorrowShiftValues;
}