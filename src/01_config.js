/********************
 * 01_config.js
 * 集中定义 Sheet 名称、期限列表、曲线 ID 与信号阈值。
 ********************/

var SHEET_CURVE = '原始_收益率曲线';
var SHEET_MM = '原始_资金面';
var SHEET_FUT = '原始_国债期货';

var SHEET_HIST = '衍生_曲线历史';
var SHEET_SLOPE = '衍生_曲线斜率';

var SHEET_SIGNAL = '信号_ETF';
var SHEET_ALLOC = '信号_债券配置';

var SHEET_DASH = '看板_利率总览';
var SHEET_MACRO = '看板_宏观总览';

/**
 * 固定期限列，单位为年。
 */
var TERMS = [
  0, 0.08, 0.17, 0.25, 0.5, 0.75,
  1, 2, 3, 4, 5, 6, 7, 8, 9, 10,
  15, 20, 30, 40, 50
];

/**
 * 中债曲线配置。
 */
var CURVES = [
  { name: '国债', id: '2c9081e50a2f9606010a3068cae70001' },
  { name: '国开债', id: '8a8b2ca037a7ca910137bfaa94fa5057' },
  { name: 'AAA信用', id: '2c9081e50a2f9606010a309f4af50111' },
  { name: 'AA+信用', id: '2c908188138b62cd01139a2ee6b51e25' }
];

/**
 * ETF 信号阈值。
 */
var SIGNAL_THRESHOLDS = {
  steep_low: 0.20,
  steep_high: 1.00
};
