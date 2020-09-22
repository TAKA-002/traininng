//このファイル自体を複製
//複製したファイル名をリネームする

function copy_this_file() {
    var year = targetYear();
    var month = targetMonth();

    //★★★①「業務管理シート」のファイルIDを記述★★★
    var willBeCopiedFill = DriveApp.getFileById('1L5fqXV1iRkLWeSqe2z8dwTd7GKu2RGTR9G9RBcphPnU');

    //★★★②複製したファイルの移動先を指定★★★
    var outputFolder = DriveApp.getFolderById('19uX3BktgUEpZyNXd4COca31Mrc8iyEKt');

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
