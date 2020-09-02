//仮想シフト表から「月を取得」してmm形式でリターン
function getMakingMonth() {
  //仮想シフト表のIDを取得する
  const ss = SpreadsheetApp.openById('1YWp9IUB1LGed4LZGYd2DNKvQx4RqSTozmWYs-hSiAWI');
  
  //取得したスプレッドシートのシフト表のシートを取得する
  const sheet = ss.getSheetByName('シート1');//シフト表のシート名にする
  
  //シフト表のシートの作成月の記載されているセルを取得する
  const range = sheet.getRange('B3').getValue();//シフト表の年月日を入力しているセルにする
  
  //Dateオブジェクトで指定したセルの値をインスタンスする
  var shiftDate = new Date(range);
  
  //セルの月を２桁に統一する
  var month = ("0" + (shiftDate.getMonth() + 1)).slice(-2);
  return month;
}

//仮想シフト表から「年」を取得を取得してyyyy形式でリターン
function getMakingYear() {
  const ss = SpreadsheetApp.openById('1YWp9IUB1LGed4LZGYd2DNKvQx4RqSTozmWYs-hSiAWI');
  const sheet = ss.getSheetByName('シート1');
  const range = sheet.getRange('B3').getValue();
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
  var templeteFile = DriveApp.getFileById('1lsxixwYUSTfWOStrcM2mmocPjgyFvNyTbj2B395T6WU');
  
  //出力先のフォルダーを取得
  var OutputFolder = DriveApp.getFolderById('1ekDBmDsLyo8gqxxRenS3neNhz6Sg0W91');
  
  //出力する時のファイル名を指示する
  //  var OutputFileName = templeteFile.getName().replace('仮想業務管理シート', '')+Utilities.formatDate(new Date(), 'JST', 'yyyyMM');
  var OutputFileName = templeteFile.getName().replace('仮想業務管理シート', '業務管理シート_' + makingYear + makingMonth);
  
  //実行
  templeteFile.makeCopy(OutputFileName, OutputFolder);
}