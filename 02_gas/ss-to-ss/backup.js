//============================
//コードの目的
//============================
//開いているファイル自体を複製する
//自動で複製したファイルのファイル名をリネームする

//============================
//コード調整箇所
//============================
//①「業務管理シート」のファイルIDを記述する
//②複製したファイルを格納するフォルダを記述する

//①
const MANAGEMENT_SHEET_ID = '1L5fqXV1iRkLWeSqe2z8dwTd7GKu2RGTR9G9RBcphPnU';
//②
const OUTPUT_FOLDER_ID = '19uX3BktgUEpZyNXd4COca31Mrc8iyEKt';


//実行関数
function copy_this_file() {
    var year = targetYear();
    var month = targetMonth();
    var willBeCopiedFill = DriveApp.getFileById(MANAGEMENT_SHEET_ID); //①
    var outputFolder = DriveApp.getFolderById(OUTPUT_FOLDER_ID); //②

    //業務管理シートのファイルをコピーする
    willBeCopiedFill.makeCopy(willBeCopiedFill.getName() + '（' + year + month + '）', outputFolder);
}


function targetYear() {
    var date = new Date();
    var targetYear = date.getFullYear();
    //年をまたぐ時
    var month = date.getMonth() + 1;
    if (month === 1) {
        var targetYear = targetYear - 1;
    }
    return targetYear;
}


function targetMonth() {
    var date = new Date();
    //セルの月を２桁に統一する
    var targetMonth = ("0" + date.getMonth()).slice(-2);
    //1月に12月分を作成する時
    if (targetMonth === "00") {
        var targetMonth = 12;
    }
    return targetMonth;
}