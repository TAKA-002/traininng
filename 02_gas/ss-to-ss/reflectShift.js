/*============================
コードの目的
============================*/
/*
シフト表から各シートに反映させたシフトにもとづいて、タイムバーを表示させる
*/

/*============================
業務管理シートの調整箇所
============================*/
/*
１：平日テンプレート、土日テンプレート、祝日テンプレートそれぞれにボタンを配置する
２：設置ボタンにchenge関数をセット
*/

/*============================
調整箇所
============================*/
const SHIFT_START_ROW = 7;
const SHIFT_START_COLUMN = 2;


//実行関数：各シートのボタンを押した時に起動
function chenge() {
  //ボタンをおしたら、アクティブなシートを取得
  var sheet = SpreadsheetApp.getActiveSheet();

  //アクティブなsheetの名前を取得して、「平日・休日・祝日マスタ」「平日・休日・祝日テンプレート改」ならこのスクリプト何もせずは終了（確認OK）
  if (sheet.getName() == WEEKDAY_MASTER_SHEET_LABRL || sheet.getName() == HOLIDAY_MASTER_SHEET_LABRL || sheet.getName() == PUBLIC_HOLIDAY_MASTER_SHEET_LABRL || sheet.getName() == WEEKDAY_TEMP_SHEET || sheet.getName() == WEEKEND_TEMP_SHEET || sheet.getName() == HOLIDAY_TEMP_SHEET) {
    return
  }

  //マスタとテンプレ以外なら、セルの取得からsetTimeまでをループさせる
  var AreaRanges = sheet.getRange(SHIFT_START_ROW, SHIFT_START_COLUMN, DESISTA_MEMBER_COUNT + SEIJI_MEMBER_COUNT + 1);
  var AreaValues = AreaRanges.getValues();
  var AdjustmentShiftValues = Array.prototype.concat.apply([], AreaValues);
  var AreaCount = AdjustmentShiftValues.length;

  for (var i = 1; i <= AreaCount; i++) {
    //そのスプレッドシート のB7セルをアクティブ（選択）にする。
    var cell = sheet.getRange((6 + i), 2).activate();

    //取得したセルが２カラム目（B列）以外なら終了。51行目より下の行なら終了
    if ((cell.getColumn() != 2) || (cell.getRow() > COLUMN_INVALID_LINE)) {
      return;
    }

    // アクティブな部分が空だったらリセットして終了
    if (cell.getValue() == "") {
      resetColor(cell.getRow(), sheet);
      continue;
    }

    //アクティブなセルの値が「休」ならループを飛ばす
    if (cell.getValue() == "休") {
      continue;
    }

    //セルがアクティブになったら、LABEL_POSITIONの値「平日」「休日」「祝日」のどれかを取得
    var labelWeek = sheet.getRange(LABEL_POSITION).getValue();
    //開いているシートを再度取得
    var document = SpreadsheetApp.getActive();

    //取得したラベル（AW4セルの値）が「平日」であったら
    if (labelWeek == LABEL_WEEKDAY) {
      //B7の行（7）を取得、その値を取得、開いているシートそのもの、開いているスプレッドシートの「平日マスタ」シートの名前を取得
      setTime(cell.getRow(), cell.getValue(), sheet, document.getSheetByName(WEEKDAY_MASTER_SHEET_LABRL));
    }
    //「休日」の場合
    else if (labelWeek == LABEL_HOLIDAY) {
      setTime(cell.getRow(), cell.getValue(), sheet, document.getSheetByName(HOLIDAY_MASTER_SHEET_LABRL));
    }
    //「祝日」の場合
    else {
      setTime(cell.getRow(), cell.getValue(), sheet, document.getSheetByName(PUBLIC_HOLIDAY_MASTER_SHEET_LABRL));
    }

    // 一番下の行の時下線を引く
  }
}

function setTime(row, value, sheet, masterSheet) {
  //「平日マスタ」「休日マスタ」「祝日マスタ」のシートの最終行を取得
  var lastRow = masterSheet.getLastRow();
  var list = masterSheet.getRange(COLUMN_NAME + 1 + ":" + COLUMN_TOKUSHU2_END + lastRow).getValues();

  //A1:AM38の値を二次元配列で取得する。それをlistに格納。
  var masterRow = -1;
  for (var i0 = 0; i0 < list.length; i0++) {
    if (list[i0][INDEX_NAME] == value) {//listの一次元インデックス０のINDEX_NAME（0）が「cell.getValue()」なら、0を代入。
      masterRow = i0;
      break;
    }
  }

  if (masterRow == -1) {
    return false;
  }

  resetColor(row, sheet);

  // dateオブジェクトを文字列に変換する
  var startTime = Utilities.formatDate(list[masterRow][INDEX_START], "JST", "H");
  var startMinute = Utilities.formatDate(list[masterRow][INDEX_START], "JST", "m");
  var endTime = Utilities.formatDate(list[masterRow][INDEX_END], "JST", "H");
  var endMinute = Utilities.formatDate(list[masterRow][INDEX_END], "JST", "m");

  var halfTime = 0;
  if (startMinute == 30) {
    halfTime = 1;
  }

  var start = parseInt(startTime) * 2 + INDEX_TIME_START + halfTime;
  var range;
  if ((endTime - startTime) > 0) {
    range = (endTime - startTime) * 2;
    sheet.getRange(row, start, 1, range).setBackground(list[masterRow][INDEX_BACKGROUND_COLOR]);
    sheet.getRange(row, start, 1, range).setFontColor(list[masterRow][INDEX_FONT_COLOR]);
  } else {
    range = (24 - parseInt(startTime)) * 2;
    sheet.getRange(row, start, 1, range).setBackground(list[masterRow][INDEX_BACKGROUND_COLOR]);
    sheet.getRange(row, start, 1, range).setFontColor(list[masterRow][INDEX_FONT_COLOR]);
    start = INDEX_TIME_START;
    if (endMinute == 30) {
      halfTime = 1;
    }
    range = parseInt(endTime) * 2 + halfTime;
    sheet.getRange(row, start, 1, range).setBackground(list[masterRow][INDEX_BACKGROUND_COLOR]);
    sheet.getRange(row, start, 1, range).setFontColor(list[masterRow][INDEX_FONT_COLOR]);
  }

  var restStartCell = list[masterRow][INDEX_REST_START];
  var restEndCell = list[masterRow][INDEX_REST_END];
  if (restStartCell != '' || restEndCell != '') {
    var restStartTime = Utilities.formatDate(restStartCell, "JST", "H");
    var restStartMinute = Utilities.formatDate(restStartCell, "JST", "m");
    var restEndTime = Utilities.formatDate(restEndCell, "JST", "H");
    var restEndMinute = Utilities.formatDate(restEndCell, "JST", "m");
    halfTime = 0;
    if (restStartMinute == 30) {
      halfTime = 1;
    }
    start = parseInt(restStartTime) * 2 + INDEX_TIME_START + halfTime;
    range = (restEndTime - restStartTime) * 2;
    var restRange = sheet.getRange(row, start, 1, range);
    restRange.merge();
    restRange.setValue("休憩");
    restRange.setBorder(false, true, null, true, true, true, null, SpreadsheetApp.BorderStyle.SOLID);
  }

  var Pickup1StartCell = list[masterRow][INDEX_PICKUP1_START];
  var Pickup1EndCell = list[masterRow][INDEX_PICKUP1_END];
  if (Pickup1StartCell != '' || Pickup1EndCell != '') {
    var Pickup1StartTime = Utilities.formatDate(Pickup1StartCell, "JST", "H");
    var Pickup1StartMinute = Utilities.formatDate(Pickup1StartCell, "JST", "m");
    var Pickup1EndTime = Utilities.formatDate(Pickup1EndCell, "JST", "H");
    var Pickup1EndMinute = Utilities.formatDate(Pickup1EndCell, "JST", "m");
    halfTime = 0;
    if (Pickup1StartMinute == 30) {
      halfTime = 1;
    }
    var halfTime1 = 0;
    if ((Pickup1StartMinute == 30) && (Pickup1EndMinute == 00)) {
      halfTime1 = -1;
    } else if ((Pickup1StartMinute == 00) && (Pickup1EndMinute == 30)) {
      halfTime1 = 1;
    } else if ((Pickup1StartMinute == 30) && (Pickup1EndMinute == 30) || (Pickup1StartMinute == 00) && (Pickup1EndMinute == 00)) {
      halfTime1 = 0;
    }
    start = parseInt(Pickup1StartTime) * 2 + INDEX_TIME_START + halfTime;
    range = (Pickup1EndTime - Pickup1StartTime) * 2 + halfTime1;
    var Pickup1Range = sheet.getRange(row, start, 1, range);
    Pickup1Range.merge();
    if ((Pickup1EndTime - Pickup1StartTime) <= 1) {
      if (((Pickup1EndMinute - Pickup1StartMinute) == 30) || ((Pickup1EndMinute - Pickup1StartMinute) == -30)) {
        Pickup1Range.setValue("P外");
      } else {
        Pickup1Range.setValue("Pickup");
      }
    } else {
      Pickup1Range.setValue("Pickup");
    }
    Pickup1Range.setBorder(false, true, null, true, true, true, null, SpreadsheetApp.BorderStyle.SOLID);
  }

  var Pickup2StartCell = list[masterRow][INDEX_PICKUP2_START];
  var Pickup2EndCell = list[masterRow][INDEX_PICKUP2_END];
  if (Pickup2StartCell != '' || Pickup2EndCell != '') {
    var Pickup2StartTime = Utilities.formatDate(Pickup2StartCell, "JST", "H");
    var Pickup2StartMinute = Utilities.formatDate(Pickup2StartCell, "JST", "m");
    var Pickup2EndTime = Utilities.formatDate(Pickup2EndCell, "JST", "H");
    var Pickup2EndMinute = Utilities.formatDate(Pickup2EndCell, "JST", "m");
    halfTime = 0;
    if (Pickup2StartMinute == 30) {
      halfTime = 1;
    }
    var halfTime1 = 0;
    if ((Pickup2StartMinute == 30) && (Pickup2EndMinute == 00)) {
      halfTime1 = -1;
    } else if ((Pickup2StartMinute == 00) && (Pickup2EndMinute == 30)) {
      halfTime1 = 1;
    } else if ((Pickup2StartMinute == 30) && (Pickup2EndMinute == 30) || (Pickup2StartMinute == 00) && (Pickup2EndMinute == 00)) {
      halfTime1 = 0;
    }
    start = parseInt(Pickup2StartTime) * 2 + INDEX_TIME_START + halfTime;
    range = (Pickup2EndTime - Pickup2StartTime) * 2 + halfTime1;
    var Pickup2Range = sheet.getRange(row, start, 1, range);
    Pickup2Range.merge();
    if ((Pickup2EndTime - Pickup2StartTime) <= 1) {
      if (((Pickup2EndMinute - Pickup2StartMinute) == 30) || ((Pickup2EndMinute - Pickup2StartMinute) == -30)) {
        Pickup2Range.setValue("P外");
      } else {
        Pickup2Range.setValue("Pickup");
      }
    } else {
      Pickup2Range.setValue("Pickup");
    }
    Pickup2Range.setBorder(false, true, null, true, true, true, null, SpreadsheetApp.BorderStyle.SOLID);
  }

  var Pickup3StartCell = list[masterRow][INDEX_PICKUP3_START];
  var Pickup3EndCell = list[masterRow][INDEX_PICKUP3_END];
  if (Pickup3StartCell != '' || Pickup3EndCell != '') {
    var Pickup3StartTime = Utilities.formatDate(Pickup3StartCell, "JST", "H");
    var Pickup3StartMinute = Utilities.formatDate(Pickup3StartCell, "JST", "m");
    var Pickup3EndTime = Utilities.formatDate(Pickup3EndCell, "JST", "H");
    var Pickup3EndMinute = Utilities.formatDate(Pickup3EndCell, "JST", "m");
    halfTime = 0;
    if (Pickup3StartMinute == 30) {
      halfTime = 1;
    }
    var halfTime1 = 0;
    if ((Pickup3StartMinute == 30) && (Pickup3EndMinute == 00)) {
      halfTime1 = -1;
    } else if ((Pickup3StartMinute == 00) && (Pickup3EndMinute == 30)) {
      halfTime1 = 1;
    } else if ((Pickup3StartMinute == 30) && (Pickup3EndMinute == 30) || (Pickup3StartMinute == 00) && (Pickup3EndMinute == 00)) {
      halfTime1 = 0;
    }
    start = parseInt(Pickup3StartTime) * 2 + INDEX_TIME_START + halfTime;
    range = (Pickup3EndTime - Pickup3StartTime) * 2 + halfTime1;
    var Pickup3Range = sheet.getRange(row, start, 1, range);
    Pickup3Range.merge();
    if ((Pickup3EndTime - Pickup3StartTime) <= 1) {
      if (((Pickup3EndMinute - Pickup3StartMinute) == 30) || ((Pickup3EndMinute - Pickup3StartMinute) == -30)) {
        Pickup3Range.setValue("P外");
      } else {
        Pickup3Range.setValue("Pickup");
      }
    } else {
      Pickup3Range.setValue("Pickup");
    }
    Pickup3Range.setBorder(false, true, null, true, true, true, null, SpreadsheetApp.BorderStyle.SOLID);
  }

  var WeatherStartCell = list[masterRow][INDEX_WEATHER_START];
  var WeatherEndCell = list[masterRow][INDEX_WEATHER_END];
  if (WeatherStartCell != '' || WeatherEndCell != '') {
    var WeatherStartTime = Utilities.formatDate(WeatherStartCell, "JST", "H");
    var WeatherStartMinute = Utilities.formatDate(WeatherStartCell, "JST", "m");
    var WeatherEndTime = Utilities.formatDate(WeatherEndCell, "JST", "H");
    var WeatherEndMinute = Utilities.formatDate(WeatherEndCell, "JST", "m");
    halfTime = 0;
    if (WeatherStartMinute == 30) {
      halfTime = 1;
    }
    var halfTime1 = 0;
    if ((WeatherStartMinute == 30) && (WeatherStartMinute == 00)) {
      halfTime1 = -1;
    } else if ((WeatherStartMinute == 00) && (WeatherStartMinute == 30)) {
      halfTime1 = 1;
    } else if ((WeatherStartMinute == 30) && (WeatherStartMinute == 30) || (WeatherStartMinute == 00) && (WeatherStartMinute == 00)) {
      halfTime1 = 0;
    }
    start = parseInt(WeatherStartTime) * 2 + INDEX_TIME_START + halfTime;
    range = (WeatherEndTime - WeatherStartTime) + halfTime1;
    var WeatherRange = sheet.getRange(row, start, 1, range);
    WeatherRange.setValue("気");
    WeatherRange.setBorder(false, true, null, true, true, true, null, SpreadsheetApp.BorderStyle.SOLID);
  }

  var PickWeatherStartCell = list[masterRow][INDEX_PICKWEATHER_START];
  var PickWeatherEndCell = list[masterRow][INDEX_PICKWEATHER_END];
  if (PickWeatherStartCell != '' || PickWeatherEndCell != '') {
    var PickWeatherStartTime = Utilities.formatDate(PickWeatherStartCell, "JST", "H");
    var PickWeatherStartMinute = Utilities.formatDate(PickWeatherStartCell, "JST", "m");
    var PickWeatherEndTime = Utilities.formatDate(PickWeatherEndCell, "JST", "H");
    var PickWeatherEndMinute = Utilities.formatDate(PickWeatherEndCell, "JST", "m");
    halfTime = 0;
    if (PickWeatherStartMinute == 30) {
      halfTime = 1;
    }
    var halfTime1 = 0;
    if ((PickWeatherStartMinute == 30) && (PickWeatherStartMinute == 00)) {
      halfTime1 = -1;
    } else if ((PickWeatherStartMinute == 00) && (PickWeatherStartMinute == 30)) {
      halfTime1 = 1;
    } else if ((PickWeatherStartMinute == 30) && (PickWeatherStartMinute == 30) || (PickWeatherStartMinute == 00) && (PickWeatherStartMinute == 00)) {
      halfTime1 = 0;
    }
    start = parseInt(PickWeatherStartTime) * 2 + INDEX_TIME_START + halfTime;
    range = (PickWeatherEndTime - PickWeatherStartTime) * 2 + halfTime1;
    var PickWeatherRange = sheet.getRange(row, start, 1, range);
    PickWeatherRange.merge();
    PickWeatherRange.setValue("P外・気");
    PickWeatherRange.setBorder(false, true, null, true, true, true, null, SpreadsheetApp.BorderStyle.SOLID);
  }

  // var Nod1StartCell = list[masterRow][INDEX_NOD1_START];
  // var Nod1EndCell = list[masterRow][INDEX_NOD1_END];
  // if (Nod1StartCell != '' || Nod1EndCell != '') {
  //   var Nod1StartTime = Utilities.formatDate(Nod1StartCell, "JST", "H");
  //   var Nod1StartMinute = Utilities.formatDate(Nod1StartCell, "JST", "m");
  //   var Nod1EndTime = Utilities.formatDate(Nod1EndCell, "JST", "H");
  //   var Nod1EndMinute = Utilities.formatDate(Nod1EndCell, "JST", "m");
  //   halfTime = 0;
  //   if (Nod1StartMinute == 30) {
  //     halfTime = 1;
  //   }
  //   var halfTime1 = 0;
  //   if ((Nod1StartMinute == 30) && (Nod1EndMinute == 00)) {
  //     halfTime1 = -1;
  //   } else if ((Nod1StartMinute == 00) && (Nod1EndMinute == 30)) {
  //     halfTime1 = 1;
  //   } else if ((Nod1StartMinute == 30) && (Nod1EndMinute == 30) || (Nod1StartMinute == 00) && (Nod1EndMinute == 00)) {
  //     halfTime1 = 0;
  //   }
  //   start = parseInt(Nod1StartTime) * 2 + INDEX_TIME_START + halfTime;
  //   range = (Nod1EndTime - Nod1StartTime) * 2 + halfTime1;
  //   var Nod1Range = sheet.getRange(row, start, 1, range);
  //   Nod1Range.merge();
  //   Nod1Range.setValue("Ｎプラ");
  //   Nod1Range.setBorder(false, true, null, true, true, true, null, SpreadsheetApp.BorderStyle.SOLID);
  // }

  var Nod2StartCell = list[masterRow][INDEX_NOD2_START];
  var Nod2EndCell = list[masterRow][INDEX_NOD2_END];
  if (Nod2StartCell != '' || Nod2EndCell != '') {
    var Nod2StartTime = Utilities.formatDate(Nod2StartCell, "JST", "H");
    var Nod2StartMinute = Utilities.formatDate(Nod2StartCell, "JST", "m");
    var Nod2EndTime = Utilities.formatDate(Nod2EndCell, "JST", "H");
    var Nod2EndMinute = Utilities.formatDate(Nod2EndCell, "JST", "m");
    halfTime = 0;
    if (Nod2StartMinute == 30) {
      halfTime = 1;
    }
    var halfTime1 = 0;
    if ((Nod2StartMinute == 30) && (Nod2EndMinute == 00)) {
      halfTime1 = -1;
    } else if ((Nod2StartMinute == 00) && (Nod2EndMinute == 30)) {
      halfTime1 = 1;
    } else if ((Nod2StartMinute == 30) && (Nod2EndMinute == 30) || (Nod2StartMinute == 00) && (Nod2EndMinute == 00)) {
      halfTime1 = 0;
    }
    start = parseInt(Nod2StartTime) * 2 + INDEX_TIME_START + halfTime;
    range = (Nod2EndTime - Nod2StartTime) * 2 + halfTime1;
    var Nod2Range = sheet.getRange(row, start, 1, range);
    Nod2Range.merge();
    Nod2Range.setValue("Ｎプラ");
    Nod2Range.setBorder(false, true, null, true, true, true, null, SpreadsheetApp.BorderStyle.SOLID);
  }

  var Region1StartCell = list[masterRow][INDEX_REGION1_START];
  var Region1EndCell = list[masterRow][INDEX_REGION1_END];
  if (Region1StartCell != '' || Region1EndCell != '') {
    var Region1StartTime = Utilities.formatDate(Region1StartCell, "JST", "H");
    var Region1StartMinute = Utilities.formatDate(Region1StartCell, "JST", "m");
    var Region1EndTime = Utilities.formatDate(Region1EndCell, "JST", "H");
    var Region1EndMinute = Utilities.formatDate(Region1EndCell, "JST", "m");
    halfTime = 0;
    if (Region1StartMinute == 30) {
      halfTime = 1;
    }
    var halfTime1 = 0;
    if ((Region1StartMinute == 30) && (Region1EndMinute == 00)) {
      halfTime1 = -1;
    } else if ((Region1StartMinute == 00) && (Region1EndMinute == 30)) {
      halfTime1 = 1;
    } else if ((Region1StartMinute == 30) && (Region1EndMinute == 30) || (Region1StartMinute == 00) && (Region1EndMinute == 00)) {
      halfTime1 = 0;
    }
    start = parseInt(Region1StartTime) * 2 + INDEX_TIME_START + halfTime;
    range = (Region1EndTime - Region1StartTime) * 2 + halfTime1;
    var Region1Range = sheet.getRange(row, start, 1, range);
    Region1Range.merge();
    Region1Range.setValue("地域");
    Region1Range.setBorder(false, true, null, true, true, true, null, SpreadsheetApp.BorderStyle.SOLID);
  }

  var Region2StartCell = list[masterRow][INDEX_REGION2_START];
  var Region2EndCell = list[masterRow][INDEX_REGION2_END];
  if (Region2StartCell != '' || Region2EndCell != '') {
    var Region2StartTime = Utilities.formatDate(Region2StartCell, "JST", "H");
    var Region2StartMinute = Utilities.formatDate(Region2StartCell, "JST", "m");
    var Region2EndTime = Utilities.formatDate(Region2EndCell, "JST", "H");
    var Region2EndMinute = Utilities.formatDate(Region2EndCell, "JST", "m");
    halfTime = 0;
    if (Region2StartMinute == 30) {
      halfTime = 1;
    }
    var halfTime1 = 0;
    if ((Region2StartMinute == 30) && (Region2EndMinute == 00)) {
      halfTime1 = -1;
    } else if ((Region2StartMinute == 00) && (Region2EndMinute == 30)) {
      halfTime1 = 1;
    } else if ((Region2StartMinute == 30) && (Region2EndMinute == 30) || (Region2StartMinute == 00) && (Region2EndMinute == 00)) {
      halfTime1 = 0;
    }
    start = parseInt(Region2StartTime) * 2 + INDEX_TIME_START + halfTime;
    range = (Region2EndTime - Region2StartTime) * 2 + halfTime1;
    var Region2Range = sheet.getRange(row, start, 1, range);
    Region2Range.merge();
    Region2Range.setValue("地域");
    Region2Range.setBorder(false, true, null, true, true, true, null, SpreadsheetApp.BorderStyle.SOLID);
  }

  var Leader1StartCell = list[masterRow][INDEX_LEADER1_START];
  var Leader1EndCell = list[masterRow][INDEX_LEADER1_END];
  if (Leader1StartCell != '' || Leader1EndCell != '') {
    var Leader1StartTime = Utilities.formatDate(Leader1StartCell, "JST", "H");
    var Leader1StartMinute = Utilities.formatDate(Leader1StartCell, "JST", "m");
    var Leader1EndTime = Utilities.formatDate(Leader1EndCell, "JST", "H");
    var Leader1EndMinute = Utilities.formatDate(Leader1EndCell, "JST", "m");
    halfTime = 0;
    if (Leader1StartMinute == 30) {
      halfTime = 1;
    }
    var halfTime1 = 0;
    if ((Leader1StartMinute == 30) && (Leader1EndMinute == 00)) {
      halfTime1 = -1;
    } else if ((Leader1StartMinute == 00) && (Leader1EndMinute == 30)) {
      halfTime1 = 1;
    } else if ((Leader1StartMinute == 30) && (Leader1EndMinute == 30) || (Leader1StartMinute == 30) && (Leader1EndMinute == 30)) {
      halfTime1 = 0;
    }
    start = parseInt(Leader1StartTime) * 2 + INDEX_TIME_START + halfTime;
    range = (Leader1EndTime - Leader1StartTime) * 2 + halfTime1;
    var Leader1Range = sheet.getRange(row, start, 1, range);
    Leader1Range.merge();
    Leader1Range.setValue("リーダー");
    Leader1Range.setBorder(false, true, null, true, true, true, null, SpreadsheetApp.BorderStyle.SOLID);
  }

  var Leader2StartCell = list[masterRow][INDEX_LEADER2_START];
  var Leader2EndCell = list[masterRow][INDEX_LEADER2_END];
  if (Leader2StartCell != '' || Leader2EndCell != '') {
    var Leader2StartTime = Utilities.formatDate(Leader2StartCell, "JST", "H");
    var Leader2StartMinute = Utilities.formatDate(Leader2StartCell, "JST", "m");
    var Leader2EndTime = Utilities.formatDate(Leader2EndCell, "JST", "H");
    var Leader2EndMinute = Utilities.formatDate(Leader2EndCell, "JST", "m");
    halfTime = 0;
    if (Leader2StartMinute == 30) {
      halfTime = 1;
    }
    var halfTime1 = 0;
    if ((Leader2StartMinute == 30) && (Leader2EndMinute == 00)) {
      halfTime1 = -1;
    } else if ((Leader2StartMinute == 00) && (Leader2EndMinute == 30)) {
      halfTime1 = 1;
    } else if ((Leader2StartMinute == 30) && (Leader2EndMinute == 30) || (Leader2StartMinute == 00) && (Leader2EndMinute == 00)) {
      halfTime1 = 0;
    }
    start = parseInt(Leader2StartTime) * 2 + INDEX_TIME_START + halfTime;
    range = (Leader2EndTime - Leader2StartTime) * 2 + halfTime1;
    var Leader2Range = sheet.getRange(row, start, 1, range);
    Leader2Range.merge();
    Leader2Range.setValue("リーダー");
    Leader2Range.setBorder(false, true, null, true, true, true, null, SpreadsheetApp.BorderStyle.SOLID);
  }

  var Line1StartCell = list[masterRow][INDEX_LINE1_START];
  var Line1EndCell = list[masterRow][INDEX_LINE1_END];
  if (Line1StartCell != '' || Line1EndCell != '') {
    var Line1StartTime = Utilities.formatDate(Line1StartCell, "JST", "H");
    var Line1StartMinute = Utilities.formatDate(Line1StartCell, "JST", "m");
    var Line1EndTime = Utilities.formatDate(Line1EndCell, "JST", "H");
    var Line1EndMinute = Utilities.formatDate(Line1EndCell, "JST", "m");
    halfTime = 0;
    if (Line1StartMinute == 30) {
      halfTime = 1;
    }
    var halfTime1 = 0;
    if ((Line1StartMinute == 30) && (Line1EndMinute == 00)) {
      halfTime1 = -1;
    } else if ((Line1StartMinute == 00) && (Line1EndMinute == 30)) {
      halfTime1 = 1;
    } else if ((Line1StartMinute == 30) && (Line1EndMinute == 30) || (Line1StartMinute == 00) && (Line1EndMinute == 00)) {
      halfTime1 = 0;
    }
    start = parseInt(Line1StartTime) * 2 + INDEX_TIME_START + halfTime;
    range = (Line1EndTime - Line1StartTime) * 2 + halfTime1;
    var Line1Range = sheet.getRange(row, start, 1, range);
    Line1Range.merge();
    Line1Range.setValue("LINE");
    Line1Range.setBorder(false, true, null, true, true, true, null, SpreadsheetApp.BorderStyle.SOLID);
  }

  var Line2StartCell = list[masterRow][INDEX_LINE2_START];
  var Line2EndCell = list[masterRow][INDEX_LINE2_END];
  if (Line2StartCell != '' || Line2EndCell != '') {
    var Line2StartTime = Utilities.formatDate(Line2StartCell, "JST", "H");
    var Line2StartMinute = Utilities.formatDate(Line2StartCell, "JST", "m");
    var Line2EndTime = Utilities.formatDate(Line2EndCell, "JST", "H");
    var Line2EndMinute = Utilities.formatDate(Line2EndCell, "JST", "m");
    halfTime = 0;
    if (Line2StartMinute == 30) {
      halfTime = 1;
    }
    var halfTime1 = 0;
    if ((Line2StartMinute == 30) && (Line2EndMinute == 00)) {
      halfTime1 = -1;
    } else if ((Line2StartMinute == 00) && (Line2EndMinute == 30)) {
      halfTime1 = 1;
    } else if ((Line2StartMinute == 30) && (Line2EndMinute == 30) || (Line2StartMinute == 00) && (Line2EndMinute == 00)) {
      halfTime1 = 0;
    }
    start = parseInt(Line2StartTime) * 2 + INDEX_TIME_START + halfTime;
    range = (Line2EndTime - Line2StartTime) * 2 + halfTime1;
    var Line2Range = sheet.getRange(row, start, 1, range);
    Line2Range.merge();
    Line2Range.setValue("LI");
    Line2Range.setBorder(false, true, null, true, true, true, null, SpreadsheetApp.BorderStyle.SOLID);
  }

  var Line3StartCell = list[masterRow][INDEX_LINE3_START];
  var Line3EndCell = list[masterRow][INDEX_LINE3_END];
  if (Line3StartCell != '' || Line3EndCell != '') {
    var Line3StartTime = Utilities.formatDate(Line3StartCell, "JST", "H");
    var Line3StartMinute = Utilities.formatDate(Line3StartCell, "JST", "m");
    var Line3EndTime = Utilities.formatDate(Line3EndCell, "JST", "H");
    var Line3EndMinute = Utilities.formatDate(Line3EndCell, "JST", "m");
    halfTime = 0;
    if (Line3StartMinute == 30) {
      halfTime = 1;
    }
    var halfTime1 = 0;
    if ((Line3StartMinute == 30) && (Line3EndMinute == 00)) {
      halfTime1 = -1;
    } else if ((Line3StartMinute == 00) && (Line3EndMinute == 30)) {
      halfTime1 = 1;
    } else if ((Line3StartMinute == 30) && (Line3EndMinute == 30) || (Line3StartMinute == 00) && (Line3EndMinute == 00)) {
      halfTime1 = 0;
    }
    start = parseInt(Line3StartTime) * 2 + INDEX_TIME_START + halfTime;
    range = (Line3EndTime - Line3StartTime) * 2 + halfTime1;
    var Line3Range = sheet.getRange(row, start, 1, range);
    Line3Range.merge();
    Line3Range.setValue("LINE");
    Line3Range.setBorder(false, true, null, true, true, true, null, SpreadsheetApp.BorderStyle.SOLID);
  }

  var Tokushu1StartCell = list[masterRow][INDEX_TOKUSHU1_START];
  var Tokushu1EndCell = list[masterRow][INDEX_TOKUSHU1_END];
  if (Tokushu1StartCell != '' || Tokushu1EndCell != '') {
    var Tokushu1StartTime = Utilities.formatDate(Tokushu1StartCell, "JST", "H");
    var Tokushu1StartMinute = Utilities.formatDate(Tokushu1StartCell, "JST", "m");
    var Tokushu1EndTime = Utilities.formatDate(Tokushu1EndCell, "JST", "H");
    var Tokushu1EndMinute = Utilities.formatDate(Tokushu1EndCell, "JST", "m");
    halfTime = 0;
    if (Tokushu1StartMinute == 30) {
      halfTime = 1;
    }
    var halfTime1 = 0;
    if ((Tokushu1StartMinute == 30) && (Tokushu1EndMinute == 00)) {
      halfTime1 = -1;
    } else if ((Tokushu1StartMinute == 00) && (Tokushu1EndMinute == 30)) {
      halfTime1 = 1;
    } else if ((Tokushu1StartMinute == 30) && (Tokushu1EndMinute == 30) || (Tokushu1StartMinute == 00) && (Tokushu1EndMinute == 00)) {
      halfTime1 = 0;
    }
    start = parseInt(Tokushu1StartTime) * 2 + INDEX_TIME_START + halfTime;
    range = (Tokushu1EndTime - Tokushu1StartTime) * 2 + halfTime1;
    var Tokushu1Range = sheet.getRange(row, start, 1, range);
    Tokushu1Range.merge();
    Tokushu1Range.setValue("特集");
    Tokushu1Range.setBorder(false, true, null, true, true, true, null, SpreadsheetApp.BorderStyle.SOLID);
  }

  var Tokushu2StartCell = list[masterRow][INDEX_TOKUSHU2_START];
  var Tokushu2EndCell = list[masterRow][INDEX_TOKUSHU2_END];
  if (Tokushu2StartCell != '' || Tokushu2EndCell != '') {
    var Tokushu2StartTime = Utilities.formatDate(Tokushu2StartCell, "JST", "H");
    var Tokushu2StartMinute = Utilities.formatDate(Tokushu2StartCell, "JST", "m");
    var Tokushu2EndTime = Utilities.formatDate(Tokushu2EndCell, "JST", "H");
    var Tokushu2EndMinute = Utilities.formatDate(Tokushu2EndCell, "JST", "m");
    halfTime = 0;
    if (Tokushu2StartMinute == 30) {
      halfTime = 1;
    }
    var halfTime1 = 0;
    if ((Tokushu2StartMinute == 30) && (Tokushu2EndMinute == 00)) {
      halfTime1 = -1;
    } else if ((Tokushu2StartMinute == 00) && (Tokushu2EndMinute == 30)) {
      halfTime1 = 1;
    } else if ((Tokushu2StartMinute == 30) && (Tokushu2EndMinute == 30) || (Tokushu2StartMinute == 00) && (Tokushu2EndMinute == 00)) {
      halfTime1 = 0;
    }
    start = parseInt(Tokushu2StartTime) * 2 + INDEX_TIME_START + halfTime;
    range = (Tokushu2EndTime - Tokushu2StartTime) * 2 + halfTime1;
    var Tokushu2Range = sheet.getRange(row, start, 1, range);
    Tokushu2Range.merge();
    Tokushu2Range.setValue("特集");
    Tokushu2Range.setBorder(false, true, null, true, true, true, null, SpreadsheetApp.BorderStyle.SOLID);
  }

}

// アクティブになっているセルの色をかえる
function resetColor(row, sheet) {
  sheet.getRange(COLUMN_TIME_START + row + ":" + COLUMN_TIME_END + row).setBackground(COLOR_BLANK_BG);//C列の行からAX列の行までを＃FFFにする
  sheet.getRange(COLUMN_TIME_START + row + ":" + COLUMN_TIME_END + row).setFontColor(COLOR_BLANK_FONT);//C列の行からAX列の行までのフォントカラーを#000にする
  sheet.getRange(COLUMN_TIME_START + row + ":" + COLUMN_TIME_END + row).breakApart();//C列の行からAX列の行までのセルの結合を解除する
  sheet.getRange(COLUMN_TIME_START + row + ":" + COLUMN_TIME_END + row).setBorder(false, false, null, false, false, false, null, null);//C列の行からAX列の行までのラインを低位置を変更する
  sheet.getRange(COLUMN_TIME_START + row).setBorder(false, true, null, false, false, false, null, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  sheet.getRange(COLUMN_TIME_END + row).setBorder(false, false, null, true, false, false, null, SpreadsheetApp.BorderStyle.SOLID_THICK);
  sheet.getRange(COLUMN_TIME_START + row + ":" + COLUMN_TIME_END + row).setValue("");//valueも空にする
}


function won() {
  Browser.msgBox('Garururu!!');
}