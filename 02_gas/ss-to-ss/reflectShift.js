//シフト表のセル範囲を取得する
function set_range_of_shift() {
  //仮想シフト表のIDを取得する
  const shift = SpreadsheetApp.openById('1DNxOMCa8u3qcZLVws9kvxh2TdLMr7rLdu4ffc7aLy0M');//★★★①シフト表のIDを入力★★★

  //取得したシフト表の対象シートを取得する
  const shiftSheet = shift.getSheetByName('10月シフト');//★★★②シフト表のシート名を入力★★★

  //運用中のスプレッドシートを取得
  var targetSheet = SpreadsheetApp.getActiveSpreadsheet();

  var lastDay = findLastDays();
  var year = getMakingYear();
  var month = getMakingMonth();

  //シフト表の初期列を取得
  //var startCol = 4;//D列
  //シフト表の最終列を取得
  //var lastCol = startCol + lastDay - 1;//初期列＋月の日数ー１

  for (var i = 0; i < lastDay; i++) {
    var day = ("00" + (i+1)).slice(-2);
    var wantToChengeSheet = targetSheet.getSheetByName(year + month + day);//運用中のスプレッドシートをシート名から取得する（20200901）
    var targetRanges = wantToChengeSheet.getRange(7, 2, 37);//セットしたいシートのセットする場所を取得
    var shiftValues = shiftSheet.getRange(6, (i + 4), 37).getValues();//シフトのセルの値を配列で取得(※最初はD6からD42まで)

    //ここで配列の値を変換する
    for (var j = 0; j < shiftValues[i].length; j++) {
      switch (shiftValues[i][j]) {
        case 'M①':
          shiftValues[i].splice(j, 1, 'マネージャー①');
          break;
        case 'M②':
          shiftValues[i].splice(j, 1, 'マネージャー②');
          break;
        case 'M③':
          shiftValues[i].splice(j, 1, 'マネージャー③');
          break;
        case 'M④':
          shiftValues[i].splice(j, 1, 'マネージャー④');
          break;
        case 'M⑤':
          shiftValues[i].splice(j, 1, 'マネージャー⑤');
          break;
        case '制①':
          shiftValues[i].splice(j, 1, '制作①');
          break;
        case '制②':
          shiftValues[i].splice(j, 1, '制作②');
          break;
        case '制③':
          shiftValues[i].splice(j, 1, '制作③');
          break;
        case '制④':
          shiftValues[i].splice(j, 1, '制作④');
          break;
        case '制⑤':
          shiftValues[i].splice(j, 1, '制作⑤');
          break;
        case '制⑥':
          shiftValues[i].splice(j, 1, '制作⑥');
          break;
        case '制⑦':
          shiftValues[i].splice(j, 1, '制作⑦');
          break;
        case '制⑧':
          shiftValues[i].splice(j, 1, '制作⑧');
          break;
        case '選':
          shiftValues[i].splice(j, 1, '選挙');
          break;
        case '特1':
          shiftValues[i].splice(j, 1, '特集①');
          break;
        case 'ES':
          shiftValues[i].splice(j, 1, 'EASY');
          break;
        case '特2':
          shiftValues[i].splice(j, 1, '特集②');
          break;
        case '特3':
          shiftValues[i].splice(j, 1, '特集③');
          break;
        case '地2':
          shiftValues[i].splice(j, 1, '地域②');
          break;
        case '地1':
          shiftValues[i].splice(j, 1, '地域①');
          break;
        case '朝L':
          shiftValues[i].splice(j, 1, '朝リーダー');
          break;
        case '朝1':
          shiftValues[i].splice(j, 1, '朝①');
          break;
        case '昼1':
          shiftValues[i].splice(j, 1, '昼①');
          break;
        case '昼2':
          shiftValues[i].splice(j, 1, '昼②');
          break;
        case '昼3':
          shiftValues[i].splice(j, 1, '昼③');
          break;
        case '昼L':
          shiftValues[i].splice(j, 1, '昼リーダー');
          break;
        case '制':
          shiftValues[i].splice(j, 1, '制作');
          break;
        case 'L':
          shiftValues[i].splice(j, 1, '夜リーダー');
          break;
        case '夜1':
          shiftValues[i].splice(j, 1, '夜①');
          break;
        case '休':
          shiftValues[i].splice(j, 1, '休');
          break;

        default:
          shiftValues[i].splice(j, 1, '');
          break;
      }
    }
    Logger.log(shiftValues);
    targetRanges.setValues(shiftValues);
  }
}
