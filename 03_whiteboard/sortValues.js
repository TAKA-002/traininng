//===============================
//実行関数
//===============================
function action() {
    //マスタシートからタスクと出勤時間を取得
    const masterData = getMasterData();

    //J2からP40の値を削除して初期化
    deleteValues();

    //マスタシートのデータを出勤時間の早い順にソート
    sortMasterData(masterData);

    //マスタのデータの最終行を取得
    const row_last_num = getMasterColumn();

    //当日のシフトデータを取得（休以外）
    //const Shift = getNextShiftData();

    //シフトを取得してな書き込み
    getShiftToDrew(row_last_num);
}


//===============================
//業務管理シートを取得
//===============================
function getManagementSheet() {
    var managementSS = SpreadsheetApp.openById(ManagementSheetID);
    return managementSS;
}

//===============================
//F列G列：2行目-最終行を取得して、休みを除外した配列を作成する（未使用）
//===============================

// function getNextShiftData() {
//     var SS = SpreadsheetApp.getActiveSpreadsheet();
//     var sheet = SS.getActiveSheet();
//     var nextShiftData = sheet.getRange(2, 6, 32, 2).getValues();

//     //    Logger.log(nextShiftData);
//     //    Logger.log(nextShiftData[0]);//[木下,休]
//     //    Logger.log(nextShiftData[0][0]);//木下
//     //    Logger.log(nextShiftData[0][1]);//休
//     //    Logger.log(nextShiftData.length);//32

//     var Shift = [];
//     for (var i = 0; i < nextShiftData.length; i++) {
//         if (nextShiftData[i][1] === "休") {
//             continue;
//         }
//         //「休」以外のシフトを配列Shiftに追加
//         Shift.push(nextShiftData[i]);
//     }
//     Logger.log(Shift)
//     return Shift;
// }


//===============================
//J2からP40の値を削除
//===============================

function deleteValues() {
    var SS = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = SS.getActiveSheet();
    var range = sheet.getRange("J2:P40");
    range.clearContent();
}

//===============================
//業務管理シートのマスタデータを取得
//===============================

function getMasterData() {
    var SS = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = SS.getActiveSheet();
    var nextShiftDate = sheet.getRange("G1").getValues();
    // var nextShiftDate = "祝日";
    var ManagementSS = getManagementSheet();

    if (nextShiftDate == '平日') {
        var lastRow = ManagementSS.getSheetByName("平日マスタ").getLastRow();
        var masterDataRange = ManagementSS.getSheetByName("平日マスタ").getRange(2, 1, lastRow - 3, 2);//0:00の行を除外するため-3
        var masterData = masterDataRange.getDisplayValues();
        return masterData;
    }
    if (nextShiftDate == '土日') {
        var lastRow = ManagementSS.getSheetByName("土日マスタ").getLastRow();
        var masterDataRange = ManagementSS.getSheetByName("土日マスタ").getRange(2, 1, lastRow - 3, 2);
        var masterData = masterDataRange.getDisplayValues();
        Logger.log(masterData);
        return masterData;
    }
    if (nextShiftDate == '祝日') {
        var lastRow = ManagementSS.getSheetByName("祝日マスタ").getLastRow();
        var masterDataRange = ManagementSS.getSheetByName("祝日マスタ").getRange(2, 1, lastRow - 3, 2);
        var masterData = masterDataRange.getDisplayValues();
        return masterData;
    }
}

//===============================
//業務管理シートのマスタデータを貼り付け　→　時間の早い順にソート
//===============================

//貼り付ける前に削除する必要がある

function sortMasterData(masterData) {
    const masterDataCount = masterData.length;
    var SS = SpreadsheetApp.getActiveSpreadsheet();
    var range = SS.getActiveSheet().getRange(2, 10, masterDataCount, 2);
    range.setValues(masterData);
    range.sort([{ column: 11, ascending: true }])
}

//===============================
//貼り付けたマスターデータのタスクが記載されている列を取得する →　最終行を取得
//===============================

function getMasterColumn() {
    var SS = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = SS.getActiveSheet();
    var ss_range = sheet.getRange("J:J").getValues();
    var row_last_num = ss_range.filter(String).length;
    return row_last_num;
}


//===============================
//J列の値を上からG列内で検索し、ヒットしたらその行のF列の名前を取得
//取得した名前を検索したJ列の行のL列に記述
//休みなら飛ばす
//G列の値の数だけ実施する
//===============================

function getShiftToDrew(row_last_num) {
    var SS = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = SS.getActiveSheet();


    //メンバーの数だけ実施
    var memberCount = 32;
    var row = 1;
    while (row <= memberCount) {//1<=32true

        //G列2行目をとってきてsfにいれる  1+1
        var sf = sheet.getRange((1 + row), 7, 1, 1);

        //sfの値をShiftにいれる
        var Shift = sf.getValue();

        if (Shift === '') {
            break;
        }

        //値が「休」だったら、とばす（次の行にいくために+1）
        if (Shift === "休") {
            row = row + 1;
            continue;
        }

        // 初期値を2で行場号の箱rowNumに入れる(初期化)
        var rowNum = 1;

        // Shiftのデータが空じゃなかったら処理をする
        //さらに「休」じゃなかったら、シートのJ列を取得する
        while (rowNum <= row_last_num) {
            var sh = sheet.getRange((1 + rowNum), 10, 1, 1);//J列

            //J列の値をMasterにいれる
            var Master = sh.getValue();

            //もしもShiftとMasterが違ったら次のJ列にいくため１足、終了
            if (Shift !== Master) {
                rowNum = rowNum + 1;
                continue;
            }

            //もしもShiftの値とMasterの値が同じだったら
            if (Shift === Master) {

                // Shiftのセルをアクティブにして
                sf.activate();

                // 左となりの値をnameに取得する
                var name = sf.offset(0, -1).getValue();

                //Masterのセルをアクティブにしてセルチェックカウントを０に。
                sh.activate();
                var cellcheckcount = 0;
                while (cellcheckcount < 5) {//5人までできる。それ以上はこの数字を変える。５なら6人目以降はスルーされる。
                    var checkedCell = sh.offset(0, (2 + cellcheckcount));
                    var checkedValue = checkedCell.getValue();
                    if (checkedValue !== '') {
                        cellcheckcount++;
                        continue;
                    }

                    if (checkedValue === '') {
                        checkedCell.setValue(name);
                        break;
                    }
                }

                // sh.offset(0, 2).setValue(name);
                //Masterのセルの右に２つずれたところにセットする


                //セットが終わったら、次の行をチェックするために１足す
                rowNum = rowNum + 1;
                continue;
            }
        }
        row = row + 1;
    }
}
