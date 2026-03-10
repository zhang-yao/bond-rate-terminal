/********************
 * 01_config.js
 * 集中定义 Sheet 常量、期限列表、曲线 ID 与信号阈值。
 ********************/

/** 原始数据表 */
var SHEET_CURVE_RAW = '原始_收益率曲线';
var SHEET_MONEY_MARKET_RAW = '原始_货币';
var SHEET_FUTURES_RAW = '原始_国债期货';

var SHEET_POLICY_RATE_RAW = "原始_政策利率";


/** 汇总指标与信号表 */
var SHEET_METRICS = '指标';
var SHEET_SIGNAL = '信号';



/** 固定期限列，单位为年。 */
var TERMS = [
  0, 0.08, 0.17, 0.25, 0.5, 0.75,
  1, 2, 3, 4, 5, 6, 7, 8, 9, 10,
  15, 20, 30, 40, 50
];

/**
 * 中债曲线配置。
 * 约定：
 * - tier 仅用于固化分层，不影响现有表结构与指标/信号口径
 * - fetch_separately=true 的曲线需要单独请求，不能与其他曲线合并抓取
 */
var CURVES = [
  // 主曲线
  { name: '国债', id: '2c9081e50a2f9606010a3068cae70001', tier: 'main', aliases: ['国债收益率曲线', 'Government Bond'] },
  { name: '国开债', id: '8a8b2ca037a7ca910137bfaa94fa5057', tier: 'main', aliases: ['国开债收益率曲线', '政策性金融债', 'Policy Bank'] },
  { name: 'AAA信用', id: '2c9081e50a2f9606010a309f4af50111', tier: 'main', aliases: ['企业债收益率曲线(AAA)', '企业债AAA', 'Enterprise Bond AAA'] },
  { name: 'AA+信用', id: '2c908188138b62cd01139a2ee6b51e25', tier: 'main', aliases: ['企业债收益率曲线(AA+)', '企业债收益率曲线(AA＋)', '企业债AA+', 'Enterprise Bond AA+'] },
  { name: 'AAA+中票', id: '2c9081e9257ddf2a012590efdded1d35', tier: 'main', aliases: ['中短期票据收益率曲线(AAA+)', '中票AAA+', 'CP&Note AAA+'] },
  { name: 'AAA中票', id: '2c9081880fa9d507010fb8505b393fe7', tier: 'main', aliases: ['中短期票据收益率曲线(AAA)', '中票AAA', 'CP&Note AAA'] },
  { name: 'AAA存单', id: '8308218D1D030E0DE0540010E03EE6DA', tier: 'main', aliases: ['同业存单收益率曲线(AAA)', '存单AAA', 'NCD AAA', 'Negotiable CD AAA'] },

  // 扩展曲线
  { name: 'AAA城投', id: '2c9081e91b55cc84011be3c53b710598', tier: 'extended', aliases: ['城投债收益率曲线(AAA)', '城投AAA', 'LGFV AAA'] },
  { name: 'AAA银行债', id: '2c9081e9259b766a0125be8b5115149f', tier: 'extended', aliases: ['商业银行普通债收益率曲线(AAA)', '商业银行债收益率曲线(AAA)', '银行债AAA', 'Financial Bond of Commercial Bank AAA'] },

  // 历史较短曲线
  { name: '地方债', id: '998183ff8c00f640018c32d4721a0d16', tier: 'short_history', fetch_separately: true, aliases: ['地方政府债收益率曲线', '地方政府债', 'Local Government'] }
];

/**
 * 信号阈值。
 * 分位列取值范围为 0~1。
 */
var SIGNAL_THRESHOLDS = {
  duration_pct_high: 0.80,
  duration_pct_low: 0.20,

  policy_spread_high: 0.80,
  policy_spread_low: 0.20,

  credit_spread_high: 0.80,
  credit_spread_low: 0.20,

  sink_spread_high: 0.80,
  sink_spread_low: 0.20,

  ncd_pct_high: 0.80,
  ncd_pct_low: 0.20,

  ultra_long_slope_high: 0.55,
  ultra_long_slope_low: 0.25,

  curve_10_1_flat: 0.35,
  curve_10_1_steep: 0.90,

  funding_tight: 1.90,
  funding_loose: 1.60
};
