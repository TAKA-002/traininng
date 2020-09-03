//仮想シフト表から「月を取得」してmm形式でリターン
function getMakingMonth() {
  //仮想シフト表のIDを取得する
  const ss = SpreadsheetApp.openById('1DNxOMCa8u3qcZLVws9kvxh2TdLMr7rLdu4ffc7aLy0M');
  
  //取得したスプレッドシートのシフト表のシートを取得する
  const sheet = ss.getSheetByName('シート1');//シフト表のシート名にする
  
  //シフト表のシートの作成月の記載されているセルを取得する
  const range = sheet.getRange('B2').getValue();//シフト表の年月日を入力しているセルにする
  
  //Dateオブジェクトで指定したセルの値をインスタンスする
  var shiftDate = new Date(range);
  
  //セルの月を２桁に統一する
  var month = ("0" + (shiftDate.getMonth() + 1)).slice(-2);
  return month;
}

//仮想シフト表から「年」を取得を取得してyyyy形式でリターン
function getMakingYear() {
  const ss = SpreadsheetApp.openById('1DNxOMCa8u3qcZLVws9kvxh2TdLMr7rLdu4ffc7aLy0M');
  const sheet = ss.getSheetByName('シート1');
  const range = sheet.getRange('B2').getValue();
  var shiftDate = new Date(range);
  var year = shiftDate.getFullYear();
  return year;
}

//業務管理シートのスプレッドシートのテンプレートを複製する
function copyTemplete(){
  //作成する年と月を取得
  var makingMonth = getMakingMonth();
  var makingYear = getMakingYear();
  
  //テンプレートとなる業務管理シートを取得
  var templeteFile = DriveApp.getFileById('1KbccBa1YAqPQpY_tAYFeVB2W8MfX4nIn1E6pgI97fY0');
  
  //出力先のフォルダーを取得
  var OutputFolder = DriveApp.getFolderById('19uX3BktgUEpZyNXd4COca31Mrc8iyEKt');
  
  //出力する時のファイル名を指示する
  //  var OutputFileName = templeteFile.getName().replace('（三浦）NHK-業務管理シート のコピー', '')+Utilities.formatDate(new Date(), 'JST', 'yyyyMM');
  var OutputFileName = templeteFile.getName().replace('（三浦）NHK-業務管理シート のコピー', '業務管理シート_' + makingYear + makingMonth);
  
  //実行
  templeteFile.makeCopy(OutputFileName, OutputFolder);
}
