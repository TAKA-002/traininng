const OUTPUTSHEET = "ホワイトボード";

const M1 = "マネージャー①";
const MORNING_LEADER = "朝リーダー";
const MORNING_OPERATOR = "朝①";


function event() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var getValuesSheet = ss.getSheetByName(SheetName);
    var outputSheet = ss.getSheetByName(OUTPUTSHEET);

    var Members = getValuesSheet.getRange(1, 4, getValuesSheet.getLastRow()).getValues().flat();
    var Positions = getValuesSheet.getRange(1, 5, getValuesSheet.getLastRow()).getValues().flat();
    Logger.log(Members);//1次元配列で取得
    Logger.log(Positions);//1次元配列で取得
    Logger.log(Members.length);//31
    Logger.log(Positions.length);//31
    Logger.log(Positions[0]);//マネージャー①
    Logger.log(Positions[1]);//マネージャー①
    Logger.log(Positions[2]);//休
    Logger.log(Positions[3]);//マネージャー⑤

}
