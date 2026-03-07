/********************
 * 00_main.js
 * 项目入口与任务编排。
 ********************/

/**
 * 本地手工测试入口。
 */
function test() {
  runEnhancedSystem();
}

/**
 * 主入口：抓取当日数据并重建派生表。
 */
function runEnhancedSystem() {
  var today = formatDate_(new Date());

  runDailyWide_(today);
  fetchPledgedRepoRates_();
  fetchBondFutures_();
  rebuildAll_();
}

/**
 * 从最近 30 天起点重新安全回补，每次最多处理 8 个非周末日期。
 */
function testBackfillSafe() {
  var end = new Date();
  var start = new Date(end);
  start.setDate(end.getDate() - 30);

  backfillBatch_(formatDate_(start), formatDate_(end), 8, true);
}

/**
 * 从回补游标继续补最近 120 天数据，每次最多处理 8 个非周末日期。
 */
function resumeBackfillSafe() {
  var end = new Date();
  var start = new Date(end);
  start.setDate(end.getDate() - 120);

  backfillBatch_(formatDate_(start), formatDate_(end), 8, false);
}

/**
 * 重建所有派生表与总览面板。
 */
function rebuildAll_() {
  updateDashboard_();
  buildCurveHistory_();
  buildCurveSlope_();
  buildETFSignal_();
  buildBondAllocationSignal_();
  buildMacroDashboard_();
}

/**
 * 输出当前回补游标。
 */
function showBackfillCursor() {
  Logger.log('BACKFILL_CURSOR=' + getBackfillCursor_());
}

/**
 * 清空当前回补游标。
 */
function resetBackfillCursor() {
  clearBackfillCursor_();
  Logger.log('BACKFILL_CURSOR cleared');
}

/**
 * 回补最近 120 天的 money_market 数据。
 */
function backfillMoneyMarketLast120Days() {
  var end = new Date();
  var start = new Date(end);
  start.setDate(end.getDate() - 120);
  backfillMoneyMarket(formatDate_(start), formatDate_(end));
}
