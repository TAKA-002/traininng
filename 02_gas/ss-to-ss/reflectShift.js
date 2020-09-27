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

//③シフト表の年月日を入力しているセルを記述
//④シフト表のスプレッドシートのIDを記述（①と同じもの）
//⑤シフト表のシート名を記述（②と同じもの）
//⑥シフト表の年月日を入力しているセルを記述（③と同じ）
//⑦生成したいシートのシート名を記述（シート名の左に生成）
//⑧土日用のテンプレート名を記述
//⑨祝日用のテンプレート名を記述
//⑩平日用のテンプレート名を記述


function main() {
  copyShiftSheet();
  repairShiftSheet();
}

function chengeValues() {
  var lastDay = findLastDays();
  var startCol = 3;
  var shiftSheet = getShiftSheet();

  for (var i = 1; i <= lastDay; i++) {
    var day = ("00" + i).slice(-2);

    var shiftRanges = shiftSheet.getRange(5, startCol + i, 37, 1);
    var shiftValues = shiftRanges.getValues();
    //Logger.log(shiftValues);
    //Logger.log(shiftValues[0][0]);
    for (var i = 0; i < shiftValues[i].length; i++) {
      for (var j = 0; j < shiftValues, length; j++) {
        if (shiftValues[i][j] === "M1") {
          shiftValues[i].splice(j, 1, 'マネージャー①');
        }
        if (shiftValues[i][j] === "M2") {
          shiftValues[i].splice(j, 1, 'マネージャー②');
        }
        if (shiftValues[i][j] === "M③") {
          shiftValues[i].splice(j, 1, 'マネージャー③');
        }
        if (shiftValues[i][j] === "M④") {
          shiftValues[i].splice(j, 1, 'マネージャー④');
        }
        if (shiftValues[i][j] === "M⑤") {
          shiftValues[i].splice(j, 1, 'マネージャー⑤');
        }
        if (shiftValues[i][j] === "制①") {
          shiftValues[i].splice(j, 1, '制作①');
        } else {
          shiftValues[i].splice(j, 1, '手動箇所');
        }
      }
    }
    //    shiftSheet.shiftRanges.setValues(shiftValues);
    //    Logger.log(shiftValues);
  }
}





//業務管理シートを取得
function getManagementSheet() {
  var managementSS = SpreadsheetApp.openById('1OIefKiiBnA3CXMxwNeMX9MxkThkNqPRGvkENO2C1LtQ');
  return managementSS;
}


//シフト表を取得
function getShiftSheet() {
  var shiftSS = SpreadsheetApp.openById('1YWp9IUB1LGed4LZGYd2DNKvQx4RqSTozmWYs-hSiAWI'); //①
  var shiftSheet = shiftSS.getSheetByName('10月シフト表'); //②
  return shiftSheet;
}

//シートをコピーして原本を残す(1番目起動)
function copyShiftSheet() {
  var shiftSS = SpreadsheetApp.openById('1YWp9IUB1LGed4LZGYd2DNKvQx4RqSTozmWYs-hSiAWI'); //①
  var shiftSheet = shiftSS.getSheetByName('10月シフト表'); //②
  var copiedSheet = shiftSheet.copyTo(shiftSS);
  copiedSheet.setName('原本：10月シフト表');
}


//原本は残して、改修(2番目起動)
function repairShiftSheet() {
  var shiftSheet = getShiftSheet();
  shiftSheet.deleteColumn(19);
  shiftSheet.deleteRows(36, 5);
  shiftSheet.deleteRows(43, 2);
  shiftSheet.insertRows(36);
}











//シフト表のセル範囲を取得する
function get_target_range_of_shift(day) {
  //仮想シフト表のIDを取得する
  const ss = SpreadsheetApp.openById('1DNxOMCa8u3qcZLVws9kvxh2TdLMr7rLdu4ffc7aLy0M'); //★★★①シフト表のIDを入力★★★

  //取得したスプレッドシートのシフト表のシートを取得する
  const sheet = ss.getSheetByName('9月シフト'); //★★★②シフト表のシート名を入力★★★

  //初期列を取得する
  var startCol = 3;

  for (var i = 1; i < lastDay; i++) {
    var shiftRanges = sheet.getRange(6, startCol + day, 32, 1).getValues();

    //ここで配列の値を変換する
    for (var j = 0; j < shiftRanges.length; j++) {
      for (var k = 0; k < shiftRanges[j].length; k++) {
        switch (shiftRanges[j][k]) {
          case 'M①':
            shiftRanges[j].splice(k, 1, 'マネージャー①');
            break;
          case 'M②':
            shiftRanges[j].splice(k, 1, 'マネージャー②');
            break;
          case 'M③':
            shiftRanges[j].splice(k, 1, 'マネージャー③');
            break;
          case 'M④':
            shiftRanges[j].splice(k, 1, 'マネージャー④');
            break;
          case 'M⑤':
            shiftRanges[j].splice(k, 1, 'マネージャー⑤');
            break;
          case '制①':
            shiftRanges[j].splice(k, 1, '制作①');
            break;
          case '制②':
            shiftRanges[j].splice(k, 1, '制作②');
            break;
          case '制③':
            shiftRanges[j].splice(k, 1, '制作③');
            break;
          case '制④':
            shiftRanges[j].splice(k, 1, '制作④');
            break;
          case '制⑤':
            shiftRanges[j].splice(k, 1, '制作⑤');
            break;
          case '制⑥':
            shiftRanges[j].splice(k, 1, '制作⑥');
            break;
          case '制⑦':
            shiftRanges[j].splice(k, 1, '制作⑦');
            break;
          case '制⑧':
            shiftRanges[j].splice(k, 1, '制作⑧');
            break;
          case '選':
            shiftRanges[j].splice(k, 1, '選挙');
            break;
          case '特1':
            shiftRanges[j].splice(k, 1, '特集①');
            break;
          case 'ES':
            shiftRanges[j].splice(k, 1, 'EASY');
            break;
          case '特2':
            shiftRanges[j].splice(k, 1, '特集②');
            break;
          case '特3':
            shiftRanges[j].splice(k, 1, '特集③');
            break;
          case '地2':
            shiftRanges[j].splice(k, 1, '地域②');
            break;
          case '地1':
            shiftRanges[j].splice(k, 1, '地域①');
            break;
          case '朝L':
            shiftRanges[j].splice(k, 1, '朝リーダー');
            break;
          case '朝1':
            shiftRanges[j].splice(k, 1, '朝①');
            break;
          case '昼1':
            shiftRanges[j].splice(k, 1, '昼①');
            break;
          case '昼2':
            shiftRanges[j].splice(k, 1, '昼②');
            break;
          case '昼3':
            shiftRanges[j].splice(k, 1, '昼③');
            break;
          case '昼L':
            shiftRanges[j].splice(k, 1, '昼リーダー');
            break;
          case '制':
            shiftRanges[j].splice(k, 1, '制作');
            break;
          case 'L':
            shiftRanges[j].splice(k, 1, '夜リーダー');
            break;
          case '夜1':
            shiftRanges[j].splice(k, 1, '夜①');
            break;
          case '休':
            shiftRanges[j].splice(k, 1, '休');
            break;

          default:
            shiftRanges[j].splice(k, 1, '');
            break;
        }
      }
    }
  }
  Logger.log(shiftRanges);
  return shiftRanges;
}

function set_shift_value_to_taeget() {
  var ss = getSS();
  var year = getMakingYear();
  var month = getMakingMonth();
  var lastDay = findLastDays();

  for (var i = 1; i <= lastDay; i++) {
    var day = ("00" + i).slice(-2);
    var targetSheet = ss.getSheetByName(year + month + day); //ターゲットのスプレッドシートをシート名から取得する（20200901）
    var targetRanges = targetSheet.getRange(7, 2, 32, 1); //セットする場所

    //ここで対象のシフト表のセルの値を取得
    var shiftValues = get_target_range_of_shift(i);
    //セットする
    targetRanges.setValues(shiftValues);
  }
}

//複製したファイルを取得する
function getSS() {
  var year = getMakingYear();
  var month = getMakingMonth();

  //ドライブの中の複製先のフォルダを指定し、その中のファイルを取得
  var files = DriveApp.getFolderById('19uX3BktgUEpZyNXd4COca31Mrc8iyEKt').getFiles(); //★★★①複製先フォルダIDを入力★★★
  var fileName = '業務管理シート_' + year + month; //★★★②新ファイル（複製したファイル）名★★★

  while (files.hasNext()) {
    var file = files.next();
    var duplicatedFileName = file.getName();
    var fileId = file.getId();
    if (duplicatedFileName === fileName) {
      //スプレッドシートを取得
      var ss = SpreadsheetApp.openById(fileId);
      break;
    }
  }
  //    Logger.log(ss);//「Spreadsheet」がログにでる
  return ss;
}