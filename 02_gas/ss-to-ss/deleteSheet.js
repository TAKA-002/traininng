/*============================
コードの目的
============================*/
/*
複製したファイル（バックアップ用ファイル）にて以下実施
不要シートを表示・非表示に限らず削除
非表示になっている先月のシートを表示
*/

/*============================
実行関数１：openDialogBoxforBackup関数の運用
============================*/
/*
業務管理シートは、10月に運用中であり、9月のバックアップとるためにボタンをおす
9月（前月）のシートは全て非表示になっている：：表示にする
10月（当月）のシートは一部非表示になっている可能性がある：：削除
表示シートは10月（当月）のシートとなっている：：削除
*/

/*============================
実行関数２：openDialogBoxforActive関数の運用例
============================*/
/*
業務管理シートは、10月に運用中であり、不要な9月の非表示シートを削除するためにボタンをおす
非表示になっているのは10月（当月）の運用には必要ない9月（前月）のシート：：削除
表示シートの中に10月（当月）シートはない
非表示のシートの中に8月（前々月）シートはない
*/

/*============================
調整箇所について
============================*/
//なし


//ボタンを押した月から「前月」を戻り値にする
function getLastMonth() {
    var date = new Date();
    var targetYear = date.getFullYear();
    //セルの月を２桁に統一する
    var Month = ("0" + (date.getMonth() + 1)).slice(-2);

    switch (Month) {
        case "01":
            targetYear = targetYear - 1;
            var targetMonth = "12";
            var target = targetYear + targetMonth;
            return target;
            break;
        case "02":
            var targetMonth = "01";
            var target = targetYear + targetMonth;
            return target;
            break;
        case "03":
            var targetMonth = "02";
            var target = targetYear + targetMonth;
            return target;
            break;
        case "04":
            var targetMonth = "03";
            var target = targetYear + targetMonth;
            return target;
            break;
        case "05":
            var targetMonth = "04";
            var target = targetYear + targetMonth;
            return target;
            break;
        case "06":
            var targetMonth = "05";
            var target = targetYear + targetMonth;
            return target;
            break;
        case "07":
            var targetMonth = "06";
            var target = targetYear + targetMonth;
            return target;
            break;
        case "08":
            var targetMonth = "07";
            var target = targetYear + targetMonth;
            return target;
            break;
        case "09":
            var targetMonth = "08";
            var target = targetYear + targetMonth;
            return target;
            break;
        case "10":
            var targetMonth = "09";
            var target = targetYear + targetMonth;
            return target;
            break;
        case "11":
            var targetMonth = "10";
            var target = targetYear + targetMonth;
            return target;
            break;
        case "12":
            var targetMonth = "11";
            var target = targetYear + targetMonth;
            return target;
            break;
    }
}


//ボタンを押した「年月（yyyymm）」を取得する
function getTargetYM() {
    var date = new Date();
    var targetYear = date.getFullYear();
    //セルの月を２桁に統一する
    var targetMonth = ("0" + (date.getMonth() + 1)).slice(-2);
    var target = targetYear + targetMonth;
    return target;
}


//不要なシートを削除する１
//【条件】：非表示、シート名が当月のもの
function deleteHiddenNowSheets() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetCount = ss.getNumSheets();
    var target = getTargetYM();

    for (var i = 0; i <= sheetCount; i++) {
        var sheet = ss.getSheets()[i];
        if (sheet != null) {
            var sheetName = sheet.getSheetName();
            if (sheetName.match(target) && sheet.isSheetHidden() === true) {//シート名が今月のyyyymmがふくまれていて、かつ「非表示」である場合
                ss.deleteSheet(sheet);
                i--;//削除したらiをマイナス１してひとつ飛ばさないようにする
            } else {
                continue;
            }
        }
        if (sheet == null) {
            break;
        }
    }
}


//不要なシートを削除する２
//【条件】：非表示、シート名が前月のもの
function deleteHiddenLastSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetCount = ss.getNumSheets();
    var target = getLastMonth();

    for (var i = 0; i <= sheetCount; i++) {
        var sheet = ss.getSheets()[i];
        if (sheet != null) {
            var sheetName = sheet.getSheetName();
            if (sheetName.match(target) && sheet.isSheetHidden() === true) {//シート名が前月のyyyymmがふくまれていて、かつ「非表示」である場合
                ss.deleteSheet(sheet);
                i--;//削除したらiをマイナス１してひとつ飛ばさないようにする
            } else {
                continue;
            }
        }
        if (sheet == null) {
            break;
        }
    }
}


//不要なシートを削除する３
//【条件】：表示、シート名が半角数字8桁
function deleteShowSheets() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetCount = ss.getNumSheets();

    for (var i = 0; i <= sheetCount; i++) {
        var sheet = ss.getSheets()[i];
        if (sheet != null) {
            var sheetName = sheet.getSheetName();
            if (sheetName.match(/[\d]{8}/) && sheet.isSheetHidden() === false) {//もしもシート名が半角数字8桁で、かつ「表示」である場合
                ss.deleteSheet(sheet);
                i--;//削除したらiをマイナス１してひとつ飛ばさないようにする
            } else {
                continue;
            }
        }
        if (sheet == null) {
            break;
        }
    }
}


//非表示のシートを表示にする
function showHiddenSheet() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetCount = ss.getNumSheets();

    for (var i = 0; i <= sheetCount; i++) {
        var sheet = ss.getSheets()[i];
        if (sheet != null && sheet.isSheetHidden() === true) {
            sheet.showSheet();
        }
    }
}


//実行関数１
function openDialogBoxforBackup() {
    var result = Browser.msgBox("注意：バックアップ用のファイルで実行してください！！実行してよろしいですか？", Browser.Buttons.OK_CANCEL);
    if (result == "ok") {
        var ss = SpreadsheetApp.getActiveSpreadsheet();
        var name = ss.getName();
        if (name === "NHK-業務管理シート") {
            Browser.msgBox("「業務管理シート」のため、スクリプトは実施されませんでした。");
            return;
        }
        if (name != "NHK-業務管理シート") {
            deleteShowSheets();
            deleteHiddenNowSheets();
            showHiddenSheet();
        }
    }
    if (result == "cancel") {
        return;
    }
}


//実行関数２
function openDialogBoxforActive() {
    var result = Browser.msgBox("注意：非表示の先月のシートが削除されます！！実行してよろしいですか？", Browser.Buttons.OK_CANCEL);
    if (result == "ok") {
        deleteHiddenLastSheet();
    }
    if (result == "cancel") {
        return;
    }
}
