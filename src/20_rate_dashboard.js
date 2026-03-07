/********************
 * 20_rate_dashboard.js
 * 从 yc_curve 生成利差面板 rate_dashboard。
 ********************/

/**
 * 从 yc_curve 聚合生成 rate_dashboard。
 */
function updateDashboard_() {
  var ss = SpreadsheetApp.getActive();
  var src = ss.getSheetByName(SHEET_CURVE);
  if (!src) return;

  var dst = ss.getSheetByName(SHEET_DASH) || ss.insertSheet(SHEET_DASH);
  ensureDashboardHeader_(dst);

  var dashIndex = buildRateDashboardDateIndex_(dst, 0);
  var data = src.getDataRange().getValues();
  if (data.length < 2) return;

  var header = data[0];
  var idxY1 = header.indexOf('Y_1');
  var idxY3 = header.indexOf('Y_3');
  var idxY5 = header.indexOf('Y_5');
  var idxY10 = header.indexOf('Y_10');

  if (idxY1 < 0 || idxY5 < 0 || idxY10 < 0) {
    Logger.log('❌ yc_curve 缺少 Y_1/Y_5/Y_10 列');
    return;
  }

  var byDate = {};
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var dateValue = row[0];
    var curveName = row[1];
    if (!dateValue || !curveName) continue;

    var key = normYMD_(dateValue);
    if (!byDate[key]) byDate[key] = {};
    byDate[key][curveName] = row;
  }

  var inserted = 0;
  var skipped = 0;

  Object.keys(byDate).forEach(function (dateKey) {
    if (dashIndex.has(dateKey)) {
      skipped++;
      return;
    }

    var gov = byDate[dateKey].国债;
    var cdb = byDate[dateKey].国开债;
    var aaa = byDate[dateKey].AAA信用;
    if (!gov || !cdb || !aaa) return;

    var gov1 = gov[idxY1];
    var gov3 = idxY3 >= 0 ? gov[idxY3] : '';
    var gov5 = gov[idxY5];
    var gov10 = gov[idxY10];
    var cdb10 = cdb[idxY10];
    var aaa5 = aaa[idxY5];

    if (gov1 === '' || gov5 === '' || gov10 === '' || cdb10 === '' || aaa5 === '') return;

    dst.appendRow([
      dateKey,
      gov1,
      gov3,
      gov5,
      gov10,
      cdb10,
      aaa5,
      cdb10 - gov10,
      aaa5 - gov5,
      gov10 - gov1
    ]);
    inserted++;
  });

  Logger.log('rate_dashboard 新增=' + inserted + ' 跳过=' + skipped);
}

/**
 * 确保 rate_dashboard 表头存在。
 */
function ensureDashboardHeader_(sheet) {
  if (sheet.getLastRow() > 0) return;

  sheet.appendRow([
    'date',
    'gov_1y',
    'gov_3y',
    'gov_5y',
    'gov_10y',
    'cdb_10y',
    'aaa_5y',
    'policy_spread',
    'credit_spread',
    'term_spread'
  ]);
}

/**
 * 构建 rate_dashboard 目标表的日期索引。
 */
function buildRateDashboardDateIndex_(sheet, dateCol0Based) {
  var last = sheet.getLastRow();
  var set = new Set();
  if (last < 2) return set;

  var values = sheet.getRange(2, dateCol0Based + 1, last - 1, 1).getValues();
  for (var i = 0; i < values.length; i++) {
    var dateValue = values[i][0];
    if (!dateValue) continue;
    set.add(normYMD_(dateValue));
  }
  return set;
}
