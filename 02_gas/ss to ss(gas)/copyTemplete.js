//①〜⑨を状況に応じて変更

//仮想シフト表から「月を取得」してmm形式でリターン
function getMakingMonth() {
  //仮想シフト表のIDを取得する
  const ss = SpreadsheetApp.openById('1DNxOMCa8u3qcZLVws9kvxh2TdLMr7rLdu4ffc7aLy0M');//★★★①シフト表のIDを入力★★★
  
  //取得したスプレッドシートのシフト表のシートを取得する
  const sheet = ss.getSheetByName('シート1');//★★★②シフト表のシート名を入力★★★
  
  //シフト表のシートの作成月の記載されているセルを取得する
  const range = sheet.getRange('B2').getValue();//★★★③シフト表の年月日を入力しているセルを入力★★★
  
  //Dateオブジェクトで指定したセルの値をインスタンス化する
  var shiftDate = new Date(range);
  
  //セルの月を２桁に統一する
  var month = ("0" + (shiftDate.getMonth() + 1)).slice(-2);
  return month;
}


//仮想シフト表から「年」を取得を取得してyyyy形式でリターン
function getMakingYear() {
  const ss = SpreadsheetApp.openById('1DNxOMCa8u3qcZLVws9kvxh2TdLMr7rLdu4ffc7aLy0M');//★★★④シフト表のIDを入力★★★
  const sheet = ss.getSheetByName('シート1');//★★★⑤シフト表のシート名を入力★★★
  const range = sheet.getRange('B2').getValue();//★★★⑥シフト表の年月日を入力しているセルを入力★★★
  var shiftDate = new Date(range);
  var year = shiftDate.getFullYear();
  return year;
}


//業務管理シートのスプレッドシートのテンプレートを複製する
function copyTemplete(){
  var makingMonth = getMakingMonth();
  var makingYear = getMakingYear();
  
  //テンプレートとなる業務管理シートを取得
  var templeteFile = DriveApp.getFileById('1KbccBa1YAqPQpY_tAYFeVB2W8MfX4nIn1E6pgI97fY0');//★★★⑦テンプレートの業務管理シートファイルIDを入力★★★
  
  //出力先のフォルダーを取得
  var OutputFolder = DriveApp.getFolderById('19uX3BktgUEpZyNXd4COca31Mrc8iyEKt');//★★★⑧出力先フォルダーIDを入力★★★
  
  //出力する時のファイル名を指示する
  var OutputFileName = templeteFile.getName().replace('（三浦）NHK-業務管理シート のコピー', '業務管理シート_' + makingYear + makingMonth);//★★★⑨replace('テンプレートのファイル名', '新ファイル名'）を入力★★★
  
  //実行
  templeteFile.makeCopy(OutputFileName, OutputFolder);
}
