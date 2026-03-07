/********************
 * 21_curve_history_slope_signal.js
 * 生成关键期限历史、曲线斜率与 ETF 信号。
 ********************/

/**
 * 从 yc_curve 提取国债关键期限，重建 curve_history。
 */
function buildCurveHistory_() {
  var ss = SpreadsheetApp.getActive();
  var src = ss.getSheetByName(SHEET_CURVE);
  if (!src) return;

  var dst = ss.getSheetByName(SHEET_HIST) || ss.insertSheet(SHEET_HIST);
  var values = src.getDataRange().getValues();
  if (values.length < 2) return;

  var header = values[0];
  var idxY1 = header.indexOf('Y_1');
  var idxY3 = header.indexOf('Y_3');
  var idxY5 = header.indexOf('Y_5');
  var idxY10 = header.indexOf('Y_10');

  var out = [['date', 'gov_1y', 'gov_3y', 'gov_5y', 'gov_10y']];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    if (row[1] !== '国债') continue;

    out.push([
      normYMD_(row[0]),
      idxY1 >= 0 ? row[idxY1] : '',
      idxY3 >= 0 ? row[idxY3] : '',
      idxY5 >= 0 ? row[idxY5] : '',
      idxY10 >= 0 ? row[idxY10] : ''
    ]);
  }

  dst.clearContents();
  dst.getRange(1, 1, out.length, out[0].length).setValues(out);
}

/**
 * 基于 curve_history 生成期限斜率表。
 */
function buildCurveSlope_() {
  var ss = SpreadsheetApp.getActive();
  var src = ss.getSheetByName(SHEET_HIST);
  if (!src) return;

  var dst = ss.getSheetByName(SHEET_SLOPE) || ss.insertSheet(SHEET_SLOPE);
  var values = src.getDataRange().getValues();
  if (values.length < 2) return;

  var out = [['date', '10Y-1Y', '10Y-3Y', '5Y-1Y']];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var y1 = row[1];
    var y3 = row[2];
    var y5 = row[3];
    var y10 = row[4];

    if (y1 === '' || y10 === '') continue;
    out.push([normYMD_(row[0]), y10 - y1, y10 - y3, y5 - y1]);
  }

  dst.clearContents();
  dst.getRange(1, 1, out.length, out[0].length).setValues(out);
}

/**
 * 根据 10Y-1Y 斜率生成简单 ETF 提示。
 */
function buildETFSignal_() {
  var ss = SpreadsheetApp.getActive();
  var src = ss.getSheetByName(SHEET_SLOPE);
  if (!src) return;

  var dst = ss.getSheetByName(SHEET_SIGNAL) || ss.insertSheet(SHEET_SIGNAL);
  var values = src.getDataRange().getValues();
  if (values.length < 2) return;

  var out = [['date', '10Y-1Y', 'signal']];
  for (var i = 1; i < values.length; i++) {
    var row = values[i];
    var steep = row[1];
    if (steep === '' || steep === null) continue;

    var signal = '中性';
    if (steep < SIGNAL_THRESHOLDS.steep_low) {
      signal = '长债机会';
    } else if (steep > SIGNAL_THRESHOLDS.steep_high) {
      signal = '短债优先';
    }

    out.push([normYMD_(row[0]), steep, signal]);
  }

  dst.clearContents();
  dst.getRange(1, 1, out.length, out[0].length).setValues(out);
}
