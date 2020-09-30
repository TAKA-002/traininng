//------▼▼コード▼▼------
//対象のシフト表を取得
//シフト表を複製
//複製したシフト表の改修
//改修したシフト表を取得
//取得した改修したシフト表の1日の値を取得
//取得した値を置き換え
//業務管理シートを取得
//業務管理シートの

//------▼▼コード調整箇所▼▼------
//①シフト表のスプレッドシートのIDを記述
//②シフト表のシート名を記述（同じシート名にしてもらう）
//③人の増減で「38」の部分の調整が必要


function main() {
  copyShiftSheet();
  repairShiftSheet();
}

function chengeAndsetValues() {
  var year = getMakingYear();
  var month = getMakingMonth();
  var lastDay = findLastDays();
  var startCol = 3;
  var shiftSheet = getShiftSheet();

  for (var i = 1; i <= lastDay; i++) {
    var day = ("00" + i).slice(-2);

    var shiftRanges = shiftSheet.getRange(6, startCol + i, 38, 1);//③
    var shiftValues = shiftRanges.getValues();
    var AdjustmentValues = Array.prototype.concat.apply([], shiftValues);
    var valuesCount = AdjustmentValues.length;

    // Logger.log(valuesCount);//38
    // Logger.log(shiftValues);//二次元配列の値
    // Logger.log(shiftValues.length);//1

    for (var j = 0; j < shiftValues.length; j++) {
      for (var k = 0; k < valuesCount; k++) {
        if (shiftValues[j][k] === "M①") {
          shiftValues[j].splice(k, 1, 'マネージャー①');
        }
        if (shiftValues[j][k] === "M②") {
          shiftValues[j].splice(k, 1, 'マネージャー②');
        }
        if (shiftValues[j][k] === "M③") {
          shiftValues[j].splice(k, 1, 'マネージャー③');
        }
        if (shiftValues[j][k] === "M④") {
          shiftValues[j].splice(k, 1, 'マネージャー④');
        }
        if (shiftValues[j][k] === "M⑤") {
          shiftValues[j].splice(k, 1, 'マネージャー⑤');
        }
        if (shiftValues[j][k] === "制①") {
          shiftValues[j].splice(k, 1, '制作①');
        }
        if (shiftValues[j][k] === "制②") {
          shiftValues[j].splice(k, 1, '制作②');
        }
        if (shiftValues[j][k] === "制③") {
          shiftValues[j].splice(k, 1, '制作③');
        }
        if (shiftValues[j][k] === "制④") {
          shiftValues[j].splice(k, 1, '制作④');
        }
        if (shiftValues[j][k] === "制⑤") {
          shiftValues[j].splice(k, 1, '制作⑤');
        }
        if (shiftValues[j][k] === "制⑥") {
          shiftValues[j].splice(k, 1, '制作⑥');
        }
        if (shiftValues[j][k] === "制⑦") {
          shiftValues[j].splice(k, 1, '制作⑦');
        }
        if (shiftValues[j][k] === "制⑧") {
          shiftValues[j].splice(k, 1, '制作⑧');
        }
        if (shiftValues[j][k] === "制") {
          shiftValues[j].splice(k, 1, '制作');
        }
        if (shiftValues[j][k] === "選") {
          shiftValues[j].splice(k, 1, '選挙');
        }
        if (shiftValues[j][k] === "特1") {
          shiftValues[j].splice(k, 1, '特集①');
        }
        if (shiftValues[j][k] === "特2") {
          shiftValues[j].splice(k, 1, '特集②');
        }
        if (shiftValues[j][k] === "ES") {
          shiftValues[j].splice(k, 1, 'EASY');
        }
        if (shiftValues[j][k] === "特3") {
          shiftValues[j].splice(k, 1, '特集③');
        }
        if (shiftValues[j][k] === "地1") {
          shiftValues[j].splice(k, 1, '地域①');
        }
        if (shiftValues[j][k] === "地2") {
          shiftValues[j].splice(k, 1, '地域②');
        }
        if (shiftValues[j][k] === "朝L") {
          shiftValues[j].splice(k, 1, '朝リーダー');
        }
        if (shiftValues[j][k] === "朝1") {
          shiftValues[j].splice(k, 1, '朝①');
        }
        if (shiftValues[j][k] === "昼1") {
          shiftValues[j].splice(k, 1, '昼①');
        }
        if (shiftValues[j][k] === "昼2") {
          shiftValues[j].splice(k, 1, '昼②');
        }
        if (shiftValues[j][k] === "昼3") {
          shiftValues[j].splice(k, 1, '昼③');
        }
        if (shiftValues[j][k] === "昼L") {
          shiftValues[j].splice(k, 1, '昼リーダー');
        }
        if (shiftValues[j][k] === "夜L") {
          shiftValues[j].splice(k, 1, '夜リーダー');
        }
        if (shiftValues[j][k] === "夜1") {
          shiftValues[j].splice(k, 1, '夜①');
        }
        if (shiftValues[j][k] === "休") {
          shiftValues[j].splice(k, 1, '休');
        }
        if (shiftValues[j][k] === "") {
          shiftValues[j].splice(k, 1, '');
        }
        else {
          shiftValues[j].splice(k, 1, '要確認');
        }
      }
    }
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var targetSheet = ss.getSheetByName(year + month + day);
    targetSheet.getRange(7, 2, 38, 1).setValues(shiftValues);//③

  //  Logger.log(ss);
  //  Logger.log(targetSheet);
  //  Logger.log(shiftValues);
  }
}


//シフト表を取得
function getShiftSheet() {
  var shiftSS = SpreadsheetApp.openById('1DNxOMCa8u3qcZLVws9kvxh2TdLMr7rLdu4ffc7aLy0M'); //①
  var shiftSheet = shiftSS.getSheetByName('シフト表'); //②
  return shiftSheet;
}

//シートをコピーして原本を残す(1番目起動)
function copyShiftSheet() {
  var shiftSS = SpreadsheetApp.openById('1DNxOMCa8u3qcZLVws9kvxh2TdLMr7rLdu4ffc7aLy0M'); //①
  var shiftSheet = shiftSS.getSheetByName('シフト表'); //②
  var copiedSheet = shiftSheet.copyTo(shiftSS);
  copiedSheet.setName('原本：シフト表');
}


//原本は残して、改修(2番目起動)
function repairShiftSheet() {
  var shiftSheet = getShiftSheet();
  shiftSheet.deleteColumn(19);
  shiftSheet.deleteRows(38, 5);
  shiftSheet.deleteRows(43, 2);
  shiftSheet.insertRows(38, 1)
}

