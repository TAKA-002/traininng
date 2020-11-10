const OUTPUTSHEET = "ホワイトボード";

const M1 = "マネージャー①";
const M2 = "マネージャー②";
const M3 = "マネージャー③";
const M4 = "マネージャー④";
const M5 = "マネージャー⑤";
const M6 = "マネージャー⑥";
const MORNING_LEADER = "朝リーダー";
const MORNING_OPERATOR = "朝①";
const SPECIAL_FEATURE1 = "特集①";
const SPECIAL_FEATURE2 = "特集②";
const SPECIAL_FEATURE3 = "特集③";
const PRODUCTION = "制作";
const PRODUCTION1 = "制作①";
const PRODUCTION2 = "制作②";
const PRODUCTION3 = "制作③";
const PRODUCTION4 = "制作④";
const PRODUCTION5 = "制作⑤";
const PRODUCTION6 = "制作⑥";
const PRODUCTION7 = "制作⑦";
const PRODUCTION8 = "制作⑧";
const ELECTION1 = "選挙①"
const ELECTION2 = "選挙②"
const EASY = "EASY";
const DAYTIME_LEADER = "昼リーダー";
const DAYTIME_OPARATOR1 = "昼①";
const DAYTIME_OPARATOR2 = "昼②";
const AREA = "地域";
const NIGHT_LEADER = "昼リーダー";
const NIGHT_OPARATOR1 = "昼①";
const NIGHT_PRODUCTION = "夜勤";

var indexNum;

function event() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var getValuesSheet = ss.getSheetByName(SheetName);//シフトデータ確認票のシートを取得
    var outputSheet = ss.getSheetByName(OUTPUTSHEET);//ホワイトボードのシートを取得

    //シフトデータ確認票の１行目D列の最終行までの値を取得 => Membersへ
    var Members = getValuesSheet.getRange(1, 4, getValuesSheet.getLastRow()).getValues().flat();

    //同じくE列の最終行までの値を取得＝Positionsへ
    var Positions = getValuesSheet.getRange(1, 5, getValuesSheet.getLastRow()).getValues().flat();

    // Logger.log(Members);//1次元配列で取得
    // Logger.log(Positions);//1次元配列で取得
    // Logger.log(Members.length);//31
    // Logger.log(Positions.length);//31
    // Logger.log(Positions[0]);//マネージャー①
    // Logger.log(Positions[1]);//マネージャー①
    // Logger.log(Positions[2]);//休
    // Logger.log(Positions[3]);//マネージャー⑤
    from7To11Position(Positions, Members, outputSheet);
}

function from7To11Position(Positions, Members, outputSheet) {
    var startPositionCell = "A1";//A1セルを取得
    var cell = outputSheet.getRange(startPositionCell);
    cell.activate();//A1セルをアクティブにする
    
    //もしもPositionsの中に朝リーダーの値があったら、アクティブなセルにセット
    //朝リーダー
    if (Positions.indexOf(MORNING_LEADER) !== -1) {
        indexNum = Positions.indexOf(MORNING_LEADER);
        cell.offset(0, 0).setValue(Positions[indexNum]);
        cell.offset(0, 1).setValue(Members[indexNum]);

    }




    //変数にインデックス番号を追加したら、A1セルに朝リーダーを記述
    //Membersの中から、同じインデックス番号の値をB１に記述
    //Positionsのそのインデックス番号の配列を削除
}