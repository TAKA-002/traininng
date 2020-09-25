//------▼▼コードの目的▼▼------
//コンテナバインドで開いているファイル自体を複製する
//複製したファイル名をリネームする

//------▼▼コード調整箇所▼▼------
//①「業務管理シート」のファイルIDを記述する
//②複製したファイルを格納するフォルダを記述する


function copy_this_file() {
    var year = targetYear();
    var month = targetMonth();

    var willBeCopiedFill = DriveApp.getFileById('1L5fqXV1iRkLWeSqe2z8dwTd7GKu2RGTR9G9RBcphPnU'); //①

    var outputFolder = DriveApp.getFolderById('19uX3BktgUEpZyNXd4COca31Mrc8iyEKt'); //②

    //業務管理シートのファイルをコピーする
    willBeCopiedFill.makeCopy(willBeCopiedFill.getName() + '（' + year + month + '）', outputFolder);
}

function targetYear() {
    var date = new Date();
    var year = date.getFullYear();

    //年をまたぐ時
    var month = date.getMonth() + 1;
    if (month === 1) {
        var year = year - 1;
    }

    return year;
}

function targetMonth() {
    var date = new Date();
    //セルの月を２桁に統一する
    var month = ("0" + date.getMonth()).slice(-2);

    //1月に12月分を作成する時
    if (month === "00") {
        var month = 12;
    }

    return month;
}