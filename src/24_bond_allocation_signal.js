/********************
 * 24_bond_allocation_signal.gs
 * 基于 10Y、120日均线、250日分位、曲线斜率、DR007 生成债券配置建议。
 *
 * 输出 Sheet：bond_allocation_signal
 * 字段：
 *   - date           日期
 *   - 10Y            国债10Y收益率
 *   - MA120          10Y 最近120个有效样本均值
 *   - pct250         10Y 在最近250个有效样本中的分位（0~1）
 *   - slope10_1      10Y-1Y 曲线斜率
 *   - dr007          DR007 加权利率
 *   - credit_spread  AAA 5Y - 国债 5Y 信用利差
 *   - funding_view   资金面判断（偏松/中性/偏紧/未知）
 *   - credit_view    信用利差判断（偏宽/中性/偏窄/未知）
 *   - regime         配置状态
 *   - long_bond      长债建议比例
 *   - mid_bond       中债建议比例
 *   - short_bond     短债建议比例
 *   - cash           现金建议比例
 *   - comment        解释说明
 ********************/

/**
 * 手工运行入口。
 */
function runBondAllocationSignal_() {
  buildBondAllocationSignal_();
}

/**
 * 主函数：重建 bond_allocation_signal。
 */
function buildBondAllocationSignal_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var curveHistorySheet = mustGetSheet_(ss, 'curve_history');
  var curveSlopeSheet = mustGetSheet_(ss, 'curve_slope');
  var moneyMarketSheet = mustGetSheet_(ss, 'money_market');
  var rateDashboardSheet = mustGetSheet_(ss, 'rate_dashboard');
  var outSheet = mustGetSheet_(ss, 'bond_allocation_signal');

  // 1) 读取 curve_history
  var curveRows = readCurveHistoryRows_(curveHistorySheet);
  if (!curveRows.length) {
    throw new Error('curve_history 无有效数据。');
  }

  // 按日期升序，并对重复日期去重（保留最后一个）
  curveRows = sortAndDedupeByDate_(curveRows, function(r) {
    return r.dateKey;
  });

  // 2) 读取辅助表映射
  var slopeMap = readCurveSlopeMap_(curveSlopeSheet);                // dateKey -> slope10_1
  var dr007Map = readMoneyMarketDr007Map_(moneyMarketSheet);        // dateKey -> dr007
  var creditSpreadMap = readRateDashboardCreditSpreadMap_(rateDashboardSheet); // dateKey -> credit_spread

  // 3) 逐日按截至当日的历史窗口计算
  var resultsAsc = [];
  var tenYHistory = [];

  for (var i = 0; i < curveRows.length; i++) {
    var row = curveRows[i];
    var tenY = toNumberOrNull_(row.gov10y);

    if (!isFiniteNumber_(tenY)) {
      continue;
    }

    // 截至当日的历史序列（先把今天放进去）
    tenYHistory.push(tenY);

    var ma120 = rollingMean_(tenYHistory, 120);
    var pct250 = rollingPercentileRank_(tenYHistory, 250, tenY);

    var slope10_1 = toNumberOrNull_(slopeMap[row.dateKey]);
    var dr007 = toNumberOrNull_(dr007Map[row.dateKey]);
    var creditSpread = toNumberOrNull_(creditSpreadMap[row.dateKey]);

    var fundingView = classifyFundingView_(dr007);
    var creditView = classifyCreditView_(creditSpread);

    var regimeObj = classifyBondRegime_({
      tenY: tenY,
      ma120: ma120,
      pct250: pct250,
      slope10_1: slope10_1,
      dr007: dr007
    });

    // 输出时避免把 NaN/null 直接写进表
    var ma120Out = isFiniteNumber_(ma120) ? ma120 : '';
    var pct250Out = isFiniteNumber_(pct250) ? pct250 : '';
    var slopeOut = isFiniteNumber_(slope10_1) ? slope10_1 : '';
    var dr007Out = isFiniteNumber_(dr007) ? dr007 : '';
    var creditSpreadOut = isFiniteNumber_(creditSpread) ? creditSpread : '';

    resultsAsc.push([
      row.dateObj,
      tenY,
      ma120Out,
      pct250Out,
      slopeOut,
      dr007Out,
      creditSpreadOut,
      fundingView,
      creditView,
      regimeObj.regime,
      regimeObj.long_bond,
      regimeObj.mid_bond,
      regimeObj.short_bond,
      regimeObj.cash,
      mergeComment_(regimeObj.comment, fundingView, creditView)
    ]);
  }

  // 4) 倒序输出（最新在上）
  var resultsDesc = resultsAsc.slice().reverse();

  // 5) 写回表
  var header = [[
    'date',
    '10Y',
    'MA120',
    'pct250',
    'slope10_1',
    'dr007',
    'credit_spread',
    'funding_view',
    'credit_view',
    'regime',
    'long_bond',
    'mid_bond',
    'short_bond',
    'cash',
    'comment'
  ]];

  outSheet.clearContents();
  outSheet.getRange(1, 1, 1, header[0].length).setValues(header);

  if (resultsDesc.length > 0) {
    outSheet.getRange(2, 1, resultsDesc.length, resultsDesc[0].length).setValues(resultsDesc);
  }

  formatBondAllocationSignalSheet_(outSheet);
}

/**
 * 读取 curve_history。
 * 预期表头：
 * date | gov_1y | gov_3y | gov_5y | gov_10y
 */
function readCurveHistoryRows_(sheet) {
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];

  var header = values[0];
  var idx = buildHeaderIndex_(header);

  var dateCol = idx['date'];
  var gov1yCol = idx['gov_1y'];
  var gov3yCol = idx['gov_3y'];
  var gov5yCol = idx['gov_5y'];
  var gov10yCol = idx['gov_10y'];

  var rows = [];

  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    var dateObj = normalizeSheetDate_(r[dateCol]);
    var gov10y = toNumberOrNull_(r[gov10yCol]);

    if (!dateObj || !isFiniteNumber_(gov10y)) continue;

    rows.push({
      dateObj: dateObj,
      dateKey: formatDateKey_(dateObj),
      gov1y: gov1yCol == null ? null : toNumberOrNull_(r[gov1yCol]),
      gov3y: gov3yCol == null ? null : toNumberOrNull_(r[gov3yCol]),
      gov5y: gov5yCol == null ? null : toNumberOrNull_(r[gov5yCol]),
      gov10y: gov10y
    });
  }

  return rows;
}

/**
 * 读取 curve_slope 映射。
 * 预期表头：
 * date | 10Y-1Y | 10Y-3Y | 5Y-1Y
 */
function readCurveSlopeMap_(sheet) {
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return {};

  var header = values[0];
  var idx = buildHeaderIndex_(header);

  var dateCol = idx['date'];
  var slopeCol = idx['10y-1y'];

  var map = {};

  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    var dateObj = normalizeSheetDate_(r[dateCol]);
    var slope = toNumberOrNull_(r[slopeCol]);

    if (!dateObj || !isFiniteNumber_(slope)) continue;
    map[formatDateKey_(dateObj)] = slope;
  }

  return map;
}

/**
 * 读取 money_market 的 DR007_weightedRate。
 * 预期表头至少包含：
 * date | DR007_weightedRate
 */
function readMoneyMarketDr007Map_(sheet) {
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return {};

  var header = values[0];
  var idx = buildHeaderIndex_(header);

  var dateCol = idx['date'];
  var dr007Col = idx['dr007_weightedrate'];

  var map = {};

  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    var dateObj = normalizeLooseDate_(r[dateCol]);
    var dr007 = toNumberOrNull_(r[dr007Col]);

    if (!dateObj || !isFiniteNumber_(dr007)) continue;
    map[formatDateKey_(dateObj)] = dr007;
  }

  return map;
}

/**
 * 读取 rate_dashboard 的 credit_spread。
 * 预期表头至少包含：
 * date | credit_spread
 */
function readRateDashboardCreditSpreadMap_(sheet) {
  var values = sheet.getDataRange().getValues();
  if (values.length < 2) return {};

  var header = values[0];
  var idx = buildHeaderIndex_(header);

  var dateCol = idx['date'];
  var creditSpreadCol = idx['credit_spread'];

  var map = {};

  for (var i = 1; i < values.length; i++) {
    var r = values[i];
    var dateObj = normalizeLooseDate_(r[dateCol]);
    var creditSpread = toNumberOrNull_(r[creditSpreadCol]);

    if (!dateObj || !isFiniteNumber_(creditSpread)) continue;
    map[formatDateKey_(dateObj)] = creditSpread;
  }

  return map;
}

/**
 * regime 分类逻辑。
 */
function classifyBondRegime_(x) {
  var tenY = x.tenY;
  var ma120 = x.ma120;
  var pct250 = x.pct250;
  var slope10_1 = x.slope10_1;
  var dr007 = x.dr007;

  if (!isFiniteNumber_(tenY) || !isFiniteNumber_(ma120) || !isFiniteNumber_(pct250)) {
    return {
      regime: 'NEUTRAL',
      long_bond: 25,
      mid_bond: 35,
      short_bond: 30,
      cash: 10,
      comment: '历史数据不足（MA120/PCT250未就绪），中性配置'
    };
  }

  var fundingTight = isFiniteNumber_(dr007) && dr007 >= 1.80;
  var curveFlat = isFiniteNumber_(slope10_1) && slope10_1 <= 0.50;

  if (pct250 <= 0.10 && tenY < ma120) {
    return {
      regime: 'VERY_DEFENSIVE',
      long_bond: 0,
      mid_bond: 20,
      short_bond: 50,
      cash: 30,
      comment: '利率低位且弱于均线，久期偏贵'
    };
  }

  if (pct250 >= 0.80 && curveFlat && !fundingTight) {
    return {
      regime: 'STRONG_BUY_LONG_BOND',
      long_bond: 70,
      mid_bond: 20,
      short_bond: 10,
      cash: 0,
      comment: '利率高位且曲线偏平，长债性价比高'
    };
  }

  if (pct250 >= 0.65 && !fundingTight) {
    return {
      regime: 'BUY_LONG_BOND',
      long_bond: 50,
      mid_bond: 25,
      short_bond: 20,
      cash: 5,
      comment: '利率偏高，可适度拉长久期'
    };
  }

  if (pct250 <= 0.20) {
    return {
      regime: 'DEFENSIVE',
      long_bond: 10,
      mid_bond: 25,
      short_bond: 45,
      cash: 20,
      comment: '利率偏低，控制久期'
    };
  }

  return {
    regime: 'NEUTRAL',
    long_bond: 25,
    mid_bond: 35,
    short_bond: 30,
    cash: 10,
    comment: '中性配置'
  };
}

/**
 * 滚动均值。
 * 严格窗口：有效样本不足 windowSize 时返回 null。
 */
function rollingMean_(arr, windowSize) {
  if (!arr || !arr.length || windowSize <= 0) return null;

  var slice = arr.slice(Math.max(0, arr.length - windowSize));
  var sum = 0;
  var n = 0;

  for (var i = 0; i < slice.length; i++) {
    if (isFiniteNumber_(slice[i])) {
      sum += slice[i];
      n++;
    }
  }

  if (n < windowSize) return null;
  return sum / n;
}

/**
 * 滚动百分位排名。
 * 定义：窗口内 <= currentValue 的个数 / 窗口样本数
 * 严格窗口：有效样本不足 windowSize 时返回 null。
 */
function rollingPercentileRank_(arr, windowSize, currentValue) {
  if (!arr || !arr.length || windowSize <= 0 || !isFiniteNumber_(currentValue)) {
    return null;
  }

  var slice = arr.slice(Math.max(0, arr.length - windowSize));
  var n = 0;
  var leCount = 0;

  for (var i = 0; i < slice.length; i++) {
    var v = slice[i];
    if (!isFiniteNumber_(v)) continue;

    n++;
    if (v <= currentValue) leCount++;
  }

  if (n < windowSize) return null;
  return leCount / n;
}

/**
 * 按日期排序并对重复日期去重。
 * 保留同一 dateKey 的最后一个元素。
 */
function sortAndDedupeByDate_(rows, keyFn) {
  var map = {};
  for (var i = 0; i < rows.length; i++) {
    map[keyFn(rows[i])] = rows[i];
  }

  var deduped = Object.keys(map).map(function(k) {
    return map[k];
  });

  deduped.sort(function(a, b) {
    return a.dateObj.getTime() - b.dateObj.getTime();
  });

  return deduped;
}

/**
 * 输出格式。
 */
function formatBondAllocationSignalSheet_(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (lastRow < 1 || lastCol < 1) return;

  sheet.getRange(1, 1, 1, lastCol).setFontWeight('bold');

  if (lastRow >= 2) {
    sheet.getRange(2, 1, lastRow - 1, lastCol).setFontSize(10);
  }

  sheet.autoResizeColumns(1, lastCol);
}

/**
 * 调试：打印最近若干行。
 */
function testBondAllocationSignalLast10_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = mustGetSheet_(ss, 'bond_allocation_signal');
  buildBondAllocationSignal_();

  var values = sheet.getDataRange().getValues();
  Logger.log(values.slice(0, Math.min(values.length, 11)));
}

/**
 * 资金面分类。
 */
function classifyFundingView_(dr007) {
  if (!isFiniteNumber_(dr007)) return '资金未知';
  if (dr007 >= 1.90) return '资金偏紧';
  if (dr007 <= 1.60) return '资金偏松';
  return '资金中性';
}

/**
 * 信用利差分类。
 */
function classifyCreditView_(creditSpread) {
  if (!isFiniteNumber_(creditSpread)) return '信用中性';
  if (creditSpread >= 0.45) return '信用利差偏宽，信用债性价比改善';
  if (creditSpread <= 0.36) return '信用利差偏窄，信用保护垫较薄';
  return '信用中性';
}

/**
 * 合并说明文本。
 */
function mergeComment_(baseComment, fundingView, creditView) {
  var parts = [];

  if (baseComment) parts.push(baseComment);
  if (fundingView) parts.push(fundingView);
  if (creditView) parts.push(creditView);

  return parts.join('；');
}