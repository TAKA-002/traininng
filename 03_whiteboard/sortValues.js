//A列をkeyにしてB列をvalueの配列を作る
//D列をkeyにしてE列をvalueの配列を作る

//valueに「休」が入っている配列を削除する

//残った配列をシフト表の7時から21時のシフトの順番にならべかえる

//D列とE列をそれぞれ取得して連想配列を作成
function makeMembers() {


    //開いているシートを取得する
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    //シフトデータ確認表のシートを取得
    var getValuesSheet = ss.getSheetByName(SheetName);

    //シフトデータ確認票の１行目D列の最終行までの値を取得 => Membersへ
    var Members = getValuesSheet.getRange(1, 4, getValuesSheet.getLastRow()).getValues().flat();

    Logger.log(Members);

    return Members;
}

//D列とE列をそれぞれ取得して連想配列を作成
function makePositions() {

    //開いているシートを取得する
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    //シフトデータ確認表のシートを取得
    var getValuesSheet = ss.getSheetByName(SheetName);

    //同じくE列の最終行までの値を取得＝Positionsへ
    var Positions = getValuesSheet.getRange(1, 5, getValuesSheet.getLastRow()).getValues().flat();

    Logger.log(Positions);

    return Positions;
}

//配列を作成
function test() {
    let member = ['水野', '熊谷', '三浦'];
    let position = ['マネージャー①', '朝リーダー', '夜リーダー'];
    let hairetsu = member.reduce(function (hairetsu, field, index) {
        hairetsu[position[index]] = field;
        return hairetsu;
    }, {});

    Logger.log(hairetsu);
}