//シフト表のセル範囲を取得する
function get_target_range_of_shift(day) {
    //仮想シフト表のIDを取得する
    const ss = SpreadsheetApp.openById('1DNxOMCa8u3qcZLVws9kvxh2TdLMr7rLdu4ffc7aLy0M');//★★★①シフト表のIDを入力★★★

    //取得したスプレッドシートのシフト表のシートを取得する
    const sheet = ss.getSheetByName('9月シフト');//★★★②シフト表のシート名を入力★★★
    var lastDay = findLastDays();

    //初期列を取得する
    var startCol = 3;

    for (var i = 1; i < lastDay; i++) {
        var shiftRanges = sheet.getRange(6, startCol + day, 32, 1).getValues();
        Logger.log(shiftRanges);
        return shiftRanges;
    }
}

function set_shift_value_to_taeget() {
    var ss = getSS();
    var year = getMakingYear();
    var month = getMakingMonth();
    var lastDay = findLastDays();

    for (var i = 1; i <= lastDay; i++) {
        var day = ("00" + i).slice(-2);
        var targetSheet = ss.getSheetByName(year + month + day);//ターゲットのスプレッドシートをシート名から取得する（20200901）
        var targetRanges = targetSheet.getRange(7, 2, 32, 1);//セットする場所

        //ここで対象のシフト表のセルの値を取得
        var shiftValues = get_target_range_of_shift(i);

        //ここで配列の値を変換する

        //セットする
        targetRanges.setValues(shiftValues);
    }
}

//複製したファイルを取得する
function getSS() {
    var year = getMakingYear();
    var month = getMakingMonth();

    //ドライブの中の複製先のフォルダを指定し、その中のファイルを取得
    var files = DriveApp.getFolderById('19uX3BktgUEpZyNXd4COca31Mrc8iyEKt').getFiles();//★★★①複製先フォルダIDを入力★★★
    var fileName = '業務管理シート_' + year + month;//★★★②新ファイル（複製したファイル）名★★★

    while (files.hasNext()) {
        var file = files.next();
        var duplicatedFileName = file.getName();
        var fileId = file.getId();
        if (duplicatedFileName === fileName) {
            //スプレッドシートを取得
            var ss = SpreadsheetApp.openById(fileId);
            break;
        }
    }
    Logger.log(ss);//「Spreadsheet」がログにでる
    return ss;
}
