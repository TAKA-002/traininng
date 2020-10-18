/*============================
コードの目的
============================*/
/*
シフト表のシートの年月から、googleカレンダーの曜日祝日データ基づいてシートを作成
作成した新シート１ヶ月分に、シフト表に基づいて「シフト」を入力
所要時間：約3分
*/

/*============================
業務管理シートの変更箇所
============================*/
/*
１：テンプレートの土日・祝日・平日のシートの年月日の表示を「yyyy年mm月dd日」へ変更する必要がある
２：実行関数を実行する前に、シフト表のスプレッドシートに作成したい月のシフト表を入れておく必要がある（シート名は汎用のものにすると変更しなくていい）
*/

/*============================
コード調整箇所
============================*/
/*
①シフト表のスプレッドシートのIDを記述
②シフト表のシート名を記述
③シフト表の年月日を入力しているセルを記述
④新しく生成するシートの生成位置となるシートを記述（このシートの左に生成）
⑤土日用のテンプレート名を記述
⑥祝日用のテンプレート名を記述
⑦平日用のテンプレート名を記述
⑧デジステメンバー数
⑨政治就活メンバー数
⑩削除したい行数
*/
const SHIFT_SS_ID = '1DNxOMCa8u3qcZLVws9kvxh2TdLMr7rLdu4ffc7aLy0M';//①
const SHIFT_SHEET_NAME = 'シフト表';//②
const SHIFT_YMD_CELL = 'B2';//③
const MAKE_NEW_SHEET_POSITION = '平日テンプレート改';//④
const WEEKEND_TEMP_SHEET = '土日テンプレート改';//⑤
const HOLIDAY_TEMP_SHEET = '祝日テンプレート改';//⑥
const WEEKDAY_TEMP_SHEET = '平日テンプレート改';//⑦
const DESISTA_MEMBER_COUNT = 32;//⑧
const SEIJI_MEMBER_COUNT = 5;//⑨
const DELETE_LINE_COUNT = 6;//⑩


//実行関数
function createAndSetValueToSheet() {
    createSheet();//シートを作る
    copyShiftSheet();//元のシフト表を保存する
    repairShiftSheet();//シフト表を改修
    chengeAndsetValues();//シフト表の値に合わせて各シートに値を反映
}


//シートを作成する月の日数を調べる
function findLastDays() {
    var year = getMakingYear();
    var month = getMakingMonth();

    var lastDay = new Date(parseInt(year, 10), parseInt(month, 10), 0).getDate();
    return lastDay;
}


//シフト表のスプレッドシートから作成するシートの「月」を取得
function getMakingMonth() {
    const ss = SpreadsheetApp.openById(SHIFT_SS_ID); //①
    const sheet = ss.getSheetByName(SHIFT_SHEET_NAME); //②
    const range = sheet.getRange(SHIFT_YMD_CELL).getValue(); //③
    var shiftDate = new Date(range);

    var month = ("0" + (shiftDate.getMonth() + 1)).slice(-2);
    return month;
}


//シフト表のスプレッドシートから作成するシートの「年」を取得
function getMakingYear() {
    const ss = SpreadsheetApp.openById(SHIFT_SS_ID); //①
    const sheet = ss.getSheetByName(SHIFT_SHEET_NAME); //②
    const range = sheet.getRange(SHIFT_YMD_CELL).getValue(); //③
    var shiftDate = new Date(range);
    var year = shiftDate.getFullYear();
    return year;
}


//業務管理シートのスプレッドシートで新シートを作成する場所を取得（平日テンプレシートの左に作成）
function getTargetSheet() {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(MAKE_NEW_SHEET_POSITION); //④

    if (sheet != null) {
        // そのシートのインデックス番号を取得
        var targetSheetIndex = sheet.getIndex();
        return targetSheetIndex;
    }
}


//実行関数１：曜日祝日に合わせてテンプレートを選択し、シートを複製（複製時にリネーム）
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
            var tempWESheet = ss.getSheetByName(WEEKEND_TEMP_SHEET); //⑤
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
            var tempHDSheet = ss.getSheetByName(HOLIDAY_TEMP_SHEET); //⑥
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
            var tempWDSheet = ss.getSheetByName(WEEKDAY_TEMP_SHEET); //⑦
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


//シフト表の値から、マスターの値に配列を変換して新シートにセットする
function chengeAndsetValues() {
    var year = getMakingYear();
    var month = getMakingMonth();
    var lastDay = findLastDays();
    var startCol = 3;
    var shiftSheet = getShiftSheet();

    for (var i = 1; i <= lastDay; i++) {
        var day = ("00" + i).slice(-2);

        var shiftRanges = shiftSheet.getRange(6, startCol + i, DESISTA_MEMBER_COUNT + DELETE_LINE_COUNT, 1);//③
        var shiftValues = shiftRanges.getValues();
        var AdjustmentValues = Array.prototype.concat.apply([], shiftValues);
        var valuesCount = AdjustmentValues.length;

        // Logger.log(valuesCount);//38
        // Logger.log(shiftValues);//二次元配列の値
        // Logger.log(shiftValues.length);//1

        for (var j = 0; j < shiftValues.length; j++) {
            for (var k = 0; k < valuesCount; k++) {
                if (shiftValues[j][k] === "M①") {
                    shiftValues[j].splice(k, 1, 'マネージャー①');
                }
                if (shiftValues[j][k] === "M②") {
                    shiftValues[j].splice(k, 1, 'マネージャー②');
                }
                if (shiftValues[j][k] === "M③") {
                    shiftValues[j].splice(k, 1, 'マネージャー③');
                }
                if (shiftValues[j][k] === "M④") {
                    shiftValues[j].splice(k, 1, 'マネージャー④');
                }
                if (shiftValues[j][k] === "M⑤") {
                    shiftValues[j].splice(k, 1, 'マネージャー⑤');
                }
                if (shiftValues[j][k] === "制①") {
                    shiftValues[j].splice(k, 1, '制作①');
                }
                if (shiftValues[j][k] === "制②") {
                    shiftValues[j].splice(k, 1, '制作②');
                }
                if (shiftValues[j][k] === "制③") {
                    shiftValues[j].splice(k, 1, '制作③');
                }
                if (shiftValues[j][k] === "制④") {
                    shiftValues[j].splice(k, 1, '制作④');
                }
                if (shiftValues[j][k] === "制⑤") {
                    shiftValues[j].splice(k, 1, '制作⑤');
                }
                if (shiftValues[j][k] === "制⑥") {
                    shiftValues[j].splice(k, 1, '制作⑥');
                }
                if (shiftValues[j][k] === "制⑦") {
                    shiftValues[j].splice(k, 1, '制作⑦');
                }
                if (shiftValues[j][k] === "制⑧") {
                    shiftValues[j].splice(k, 1, '制作⑧');
                }
                if (shiftValues[j][k] === "制") {
                    shiftValues[j].splice(k, 1, '制作');
                }
                if (shiftValues[j][k] === "選") {
                    shiftValues[j].splice(k, 1, '選挙');
                }
                if (shiftValues[j][k] === "特1") {
                    shiftValues[j].splice(k, 1, '特集①');
                }
                if (shiftValues[j][k] === "特2") {
                    shiftValues[j].splice(k, 1, '特集②');
                }
                if (shiftValues[j][k] === "ES") {
                    shiftValues[j].splice(k, 1, 'EASY');
                }
                if (shiftValues[j][k] === "特3") {
                    shiftValues[j].splice(k, 1, '特集③');
                }
                if (shiftValues[j][k] === "地1") {
                    shiftValues[j].splice(k, 1, '地域①');
                }
                if (shiftValues[j][k] === "地2") {
                    shiftValues[j].splice(k, 1, '地域②');
                }
                if (shiftValues[j][k] === "朝L") {
                    shiftValues[j].splice(k, 1, '朝リーダー');
                }
                if (shiftValues[j][k] === "朝1") {
                    shiftValues[j].splice(k, 1, '朝①');
                }
                if (shiftValues[j][k] === "昼1") {
                    shiftValues[j].splice(k, 1, '昼①');
                }
                if (shiftValues[j][k] === "昼2") {
                    shiftValues[j].splice(k, 1, '昼②');
                }
                if (shiftValues[j][k] === "昼3") {
                    shiftValues[j].splice(k, 1, '昼③');
                }
                if (shiftValues[j][k] === "昼L") {
                    shiftValues[j].splice(k, 1, '昼リーダー');
                }
                if (shiftValues[j][k] === "夜L") {
                    shiftValues[j].splice(k, 1, '夜リーダー');
                }
                if (shiftValues[j][k] === "夜1") {
                    shiftValues[j].splice(k, 1, '夜①');
                }
                if (shiftValues[j][k] === "夜制") {
                    shiftValues[j].splice(k, 1, '夜勤');
                }
                if (shiftValues[j][k] === "休") {
                    shiftValues[j].splice(k, 1, '休');
                }
                if (shiftValues[j][k] === "") {
                    shiftValues[j].splice(k, 1, '');
                }
            }
        }
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var targetSheet = ss.getSheetByName(year + month + day);
        targetSheet.getRange(7, 2, DESISTA_MEMBER_COUNT + DELETE_LINE_COUNT, 1).setValues(shiftValues);//③

        //  Logger.log(ss);
        //  Logger.log(targetSheet);
        //  Logger.log(shiftValues);
    }
}


//シフト表を取得
function getShiftSheet() {
    var shiftSS = SpreadsheetApp.openById(SHIFT_SS_ID); //①
    var shiftSheet = shiftSS.getSheetByName(SHIFT_SHEET_NAME); //②
    return shiftSheet;
}

//実行関数２：シフト表をコピーして原本を残す
function copyShiftSheet() {
    var shiftSS = SpreadsheetApp.openById(SHIFT_SS_ID); //①
    var shiftSheet = shiftSS.getSheetByName(SHIFT_SHEET_NAME); //②
    var copiedSheet = shiftSheet.copyTo(shiftSS);
    copiedSheet.setName('原本：' + SHIFT_SHEET_NAME);
}


//実行関数３：シフト表を改修
function repairShiftSheet() {
    var shiftSheet = getShiftSheet();
    shiftSheet.deleteColumn(19);
    shiftSheet.deleteRows(DESISTA_MEMBER_COUNT + DELETE_LINE_COUNT, 5);
    shiftSheet.deleteRows(DESISTA_MEMBER_COUNT + SEIJI_MEMBER_COUNT + DELETE_LINE_COUNT, 2);
    shiftSheet.insertRows(DESISTA_MEMBER_COUNT + DELETE_LINE_COUNT, 1)
}
