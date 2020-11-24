//valueに「休」が入っている配列を削除する

//残った配列をシフト表の7時から21時のシフトの順番にならべかえる


//===============================
//F列G列：2行目-最終行を取得して、休みを除外した配列を作成する
//===============================

function getNextShiftData() {
    var SS = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = SS.getActiveSheet();
    var nextShiftData = sheet.getRange(2,6,32,2).getValues();

//    Logger.log(nextShiftData);
//    Logger.log(nextShiftData[0]);//[木下,休]
//    Logger.log(nextShiftData[0][0]);//木下
//    Logger.log(nextShiftData[0][1]);//休
//    Logger.log(nextShiftData.length);//32

    var Shift = [];
    for(var i = 0; i < nextShiftData.length;i++){
        if(nextShiftData[i][1] === "休"){
            continue;
        }
        //「休」以外のシフトを配列Shiftに追加
        Shift.push(nextShiftData[i]); 
    }
    return Shift;
//    Logger.log(Shift);
}


//===============================
//残りの想定される問題
//===============================

//ホワイトボードにはるシフトの日が「平日」「祝日」「土日」のどれかをgoogleカレンダーから判断
//３つのなかから判断したものの業務管理シートの「マスタ」のA2:B最終行のvalueを取得
//B列の時間が早い順番にソート
//ソートした配列をスプレッドシートに記載
//記載されたシフトのvalueだけを配列で取得
//取得した配列の上から順番に、その値を「休」を除外した配列Shiftから探す
//ヒットしたらそのvalue(だれか)を取得して連想配列を作成
//なかったら、空文字をいれる
//新しい配列が完成したら、シートに出力

function sortShiftData(){







}