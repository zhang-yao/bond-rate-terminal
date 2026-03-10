/********************
 * 14_raw_policy_rate.gs
 * 政策利率原始事件抓取
 *
 * 表名：原始_政策利率
 * 常量：SHEET_POLICY_RATE_RAW
 *
 * 事件表字段：
 * date
 * type        // OMO / MLF / LPR
 * term        // 7D / 1Y / 5Y+
 * rate        // %
 * amount      // 亿元；LPR为空
 * source
 * fetched_at
 * note
 ********************/

var RAW_POLICY_RATE_HEADERS = [
  "date",
  "type",
  "term",
  "rate",
  "amount",
  "source",
  "fetched_at",
  "note"
];

var POLICY_RATE_SOURCE_CONFIG = {
  omo_list_url: "https://www.pbc.gov.cn/zhengcehuobisi/125207/125213/125431/125475/index.html",
  mlf_list_url: "https://www.pbc.gov.cn/zhengcehuobisi/125207/125213/125437/125446/125873/index.html",
  lpr_list_url: "https://www.pbc.gov.cn/zhengcehuobisi/125207/125213/125440/3876551/index.html"
};

/* =========================
 * 对外入口
 * ========================= */

function initRawPolicyRateSheet() {
  prEnsureRawPolicyRateSheet_();
  Logger.log("已初始化表：" + SHEET_POLICY_RATE_RAW);
}

function testFetchLatestPolicyRates() {
  var events = prFetchLatestPolicyRateEvents_();
  Logger.log(JSON.stringify(events, null, 2));
}

function syncRawPolicyRateLatest() {
  var sheet = prEnsureRawPolicyRateSheet_();
  var events = prFetchLatestPolicyRateEvents_();
  var inserted = 0;
  var updated = 0;

  for (var i = 0; i < events.length; i++) {
    var e = events[i];
    if (!e.date || !e.type || !e.term) continue;

    var existed = prFindPolicyRateEventRow_(sheet, e.date, e.type, e.term);
    prUpsertPolicyRateEvent_(sheet, e);
    if (existed) updated++;
    else inserted++;
  }

  prSortRawPolicyRateSheet_(sheet);
  Logger.log("政策利率最新事件同步完成 inserted=" + inserted + ", updated=" + updated);
}

function backfillPolicyRateRecent(omoLimit, mlfLimit, lprLimit) {
  omoLimit = Number(omoLimit || 30);
  mlfLimit = Number(mlfLimit || 200);
  lprLimit = Number(lprLimit || 240);

  var sheet = prEnsureRawPolicyRateSheet_();
  var events = []
    .concat(prFetchRecentOmoEvents_(omoLimit))
    .concat(prFetchRecentMlfEvents_(mlfLimit))
    .concat(prFetchRecentLprEvents_(lprLimit));

  var inserted = 0;
  var updated = 0;

  for (var i = 0; i < events.length; i++) {
    var e = events[i];
    if (!e.date || !e.type || !e.term) continue;

    var existed = prFindPolicyRateEventRow_(sheet, e.date, e.type, e.term);
    prUpsertPolicyRateEvent_(sheet, e);
    if (existed) updated++;
    else inserted++;
  }

  prSortRawPolicyRateSheet_(sheet);
  Logger.log("政策利率近期事件回补完成 inserted=" + inserted + ", updated=" + updated + ", total=" + events.length);
}

function readRawPolicyRateEvents_() {
  var sheet = prEnsureRawPolicyRateSheet_();
  return prReadPolicyRateRows_(sheet);
}

/* =========================
 * 最新抓取
 * ========================= */

function prFetchLatestPolicyRateEvents_() {
  var events = [];
  var fetchedAt = prFormatPolicyDateTime_(new Date());

  try {
    var omo = prFetchLatestOmo7d_();
    events.push({
      date: omo.date,
      type: "OMO",
      term: "7D",
      rate: omo.rate,
      amount: omo.amount,
      source: omo.url,
      fetched_at: fetchedAt,
      note: omo.title || ""
    });
  } catch (e) {
    Logger.log("OMO 抓取失败: " + e.message);
  }

  try {
    var mlf = prFetchLatestMlf1y_();
    events.push({
      date: mlf.date,
      type: "MLF",
      term: "1Y",
      rate: mlf.rate,
      amount: mlf.amount,
      source: mlf.url,
      fetched_at: fetchedAt,
      note: mlf.title || ""
    });
  } catch (e) {
    Logger.log("MLF 抓取失败: " + e.message);
  }

  try {
    var lpr = prFetchLatestLpr_();
    if (lpr.lpr_1y !== "") {
      events.push({
        date: lpr.date,
        type: "LPR",
        term: "1Y",
        rate: lpr.lpr_1y,
        amount: "",
        source: lpr.url,
        fetched_at: fetchedAt,
        note: lpr.title || ""
      });
    }
    if (lpr.lpr_5y !== "") {
      events.push({
        date: lpr.date,
        type: "LPR",
        term: "5Y+",
        rate: lpr.lpr_5y,
        amount: "",
        source: lpr.url,
        fetched_at: fetchedAt,
        note: lpr.title || ""
      });
    }
  } catch (e) {
    Logger.log("LPR 抓取失败: " + e.message);
  }

  return events;
}

function prFetchLatestOmo7d_() {
  var events = prFetchRecentOmoEvents_(1);
  if (!events.length) throw new Error("未找到最新 7 天 OMO 利率");

  return {
    date: events[0].date,
    rate: events[0].rate,
    amount: events[0].amount,
    url: events[0].source,
    title: events[0].note
  };
}

function prFetchLatestMlf1y_() {
  var events = prFetchRecentMlfEvents_(1);
  if (!events.length) throw new Error("未找到最新 1Y MLF 中标利率");

  return {
    date: events[0].date,
    rate: events[0].rate,
    amount: events[0].amount,
    url: events[0].source,
    title: events[0].note
  };
}

function prFetchLatestLpr_() {
  var events = prFetchRecentLprEvents_(1);
  if (!events.length) throw new Error("未找到最新 LPR");

  var out = {
    date: "",
    lpr_1y: "",
    lpr_5y: "",
    url: "",
    title: ""
  };

  for (var i = 0; i < events.length; i++) {
    var e = events[i];
    out.date = e.date || out.date;
    out.url = e.source || out.url;
    out.title = e.note || out.title;
    if (e.term === "1Y") out.lpr_1y = e.rate;
    if (e.term === "5Y+") out.lpr_5y = e.rate;
  }

  if (out.lpr_1y === "" && out.lpr_5y === "") {
    throw new Error("未解析到 LPR 数值");
  }
  return out;
}

/* =========================
 * 近期回补
 * ========================= */

function prFetchRecentOmoEvents_(limit) {
  var items = prListOmoArticles_();
  var out = [];
  var fetchedAt = prFormatPolicyDateTime_(new Date());

  for (var i = 0; i < items.length; i++) {
    if (out.length >= limit) break;

    var item = items[i];
    try {
      var html = prFetchHtml_(item.url);
      var parsed = prParseOmoDetail_(html);
      if (!parsed || parsed.rate == null) continue;

      out.push({
        date: parsed.date || item.date || "",
        type: "OMO",
        term: "7D",
        rate: parsed.rate,
        amount: parsed.amount,
        source: item.url,
        fetched_at: fetchedAt,
        note: item.title || ""
      });
    } catch (e) {
      Logger.log("prFetchRecentOmoEvents_ skip: " + item.url + " | " + e.message);
    }
  }

  out.sort(function(a, b) {
    if (a.date !== b.date) return a.date < b.date ? 1 : -1;
    return 0;
  });

  return prDedupePolicyRateEvents_(out);
}

function prFetchRecentMlfEvents_(limit) {
  var items = prListMlfArticles_();
  var out = [];
  var fetchedAt = prFormatPolicyDateTime_(new Date());

  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    try {
      var html = prFetchHtml_(item.url);
      var parsed = prParseMlfDetail_(html);
      if (!parsed || parsed.rate == null) continue;

      out.push({
        date: parsed.date || item.date || "",
        type: "MLF",
        term: "1Y",
        rate: parsed.rate,
        amount: parsed.amount,
        source: item.url,
        fetched_at: fetchedAt,
        note: item.title || ""
      });
    } catch (e) {
      Logger.log("prFetchRecentMlfEvents_ skip: " + item.url + " | " + e.message);
    }
  }

  out.sort(function(a, b) {
    return a.date < b.date ? 1 : a.date > b.date ? -1 : 0;
  });

  return prDedupePolicyRateEvents_(out).slice(0, limit);
}

function prFetchRecentLprEvents_(limit) {
  var items = prListLprArticles_();
  var out = [];
  var fetchedAt = prFormatPolicyDateTime_(new Date());

  for (var i = 0; i < items.length; i++) {
    var item = items[i];
    try {
      var html = prFetchHtml_(item.url);
      var parsed = prParseLprDetail_(html);
      if (!parsed) continue;

      if (parsed.lpr_1y !== "") {
        out.push({
          date: parsed.date || item.date || "",
          type: "LPR",
          term: "1Y",
          rate: parsed.lpr_1y,
          amount: "",
          source: item.url,
          fetched_at: fetchedAt,
          note: item.title || ""
        });
      }

      if (parsed.lpr_5y !== "") {
        out.push({
          date: parsed.date || item.date || "",
          type: "LPR",
          term: "5Y+",
          rate: parsed.lpr_5y,
          amount: "",
          source: item.url,
          fetched_at: fetchedAt,
          note: item.title || ""
        });
      }
    } catch (e) {
      Logger.log("prFetchRecentLprEvents_ skip: " + item.url + " | " + e.message);
    }
  }

  out.sort(function(a, b) {
    return a.date < b.date ? 1 : a.date > b.date ? -1 : 0;
  });

  return prDedupePolicyRateEvents_(out).slice(0, limit * 2);
}

/* =========================
 * 栏目列表
 * ========================= */

function prListOmoArticles_() {
  var html = prFetchHtml_(POLICY_RATE_SOURCE_CONFIG.omo_list_url);
  var anchors = prExtractAnchorsInOrder_(html, POLICY_RATE_SOURCE_CONFIG.omo_list_url);
  var out = [];
  var seen = {};

  for (var i = 0; i < anchors.length; i++) {
    var a = anchors[i];
    if (!/公开市场业务交易公告\s*\[\d{4}\]第\d+号/.test(a.title)) continue;
    if (a.url.indexOf("/125475/") < 0) continue;
    if (/\/125475\/index\.html$/i.test(a.url)) continue;

    var key = a.url;
    if (seen[key]) continue;
    seen[key] = true;

    out.push({
      title: a.title,
      url: a.url,
      date: "",
      seq: prExtractSeqNo_(a.title)
    });
  }

  return out;
}

function prListMlfArticles_() {
  var html = prFetchHtml_(POLICY_RATE_SOURCE_CONFIG.mlf_list_url);
  var anchors = prExtractAnchorsInOrder_(html, POLICY_RATE_SOURCE_CONFIG.mlf_list_url);
  var out = [];
  var seen = {};

  for (var i = 0; i < anchors.length; i++) {
    var a = anchors[i];
    if (!/中期借贷便利开展情况/.test(a.title)) continue;
    if (a.url.indexOf("/125873/") < 0) continue;
    if (/\/125873\/index\.html$/i.test(a.url)) continue;

    var key = a.url;
    if (seen[key]) continue;
    seen[key] = true;

    out.push({
      title: a.title,
      url: a.url,
      date: ""
    });
  }

  return out;
}

function prListLprArticles_() {
  var html = prFetchHtml_(POLICY_RATE_SOURCE_CONFIG.lpr_list_url);
  var anchors = prExtractAnchorsInOrder_(html, POLICY_RATE_SOURCE_CONFIG.lpr_list_url);
  var out = [];
  var seen = {};

  for (var i = 0; i < anchors.length; i++) {
    var a = anchors[i];
    if (!/贷款市场报价利率（LPR）公告/.test(a.title) && !/贷款市场报价利率\(LPR\)公告/.test(a.title)) continue;
    if (a.url.indexOf("/3876551/") < 0) continue;
    if (/\/3876551\/index\.html$/i.test(a.url)) continue;

    var key = a.url;
    if (seen[key]) continue;
    seen[key] = true;

    out.push({
      title: a.title,
      url: a.url,
      date: ""
    });
  }

  return out;
}

/* =========================
 * 详情页解析
 * ========================= */

function prParseOmoDetail_(html) {
  if (!html) return null;

  var text = prNormalizePolicyTextKeepLines_(prStripHtml_(html));
  var flat = prNormalizeWhitespace_(text);
  var date = prExtractArticleDate_(flat);

  var amount = "";
  var rate = null;

  // 首句抓规模：开展了485亿元7天期逆回购操作
  var mIntro = flat.match(/开展了?\s*([0-9]+(?:\.[0-9]+)?)\s*亿元\s*7\s*天期?逆回购操作/);
  if (!mIntro) {
    mIntro = flat.match(/([0-9]+(?:\.[0-9]+)?)\s*亿元\s*7\s*天期?逆回购操作/);
  }
  if (mIntro) {
    amount = Number(mIntro[1]);
  }

  // 仅解析“逆回购操作情况”之后的正文
  var idx = text.indexOf("逆回购操作情况");
  var body = idx >= 0 ? text.slice(idx) : text;

  var tokens = body
    .split("\n")
    .map(function(s) { return s.trim(); })
    .filter(function(s) { return s; });

  for (var i = 0; i < tokens.length; i++) {
    if (!/^7\s*天$/.test(tokens[i]) && !/^7\s*天期$/.test(tokens[i])) continue;

    for (var j = i + 1; j < Math.min(tokens.length, i + 8); j++) {
      if (rate == null && /%$/.test(tokens[j])) {
        var pct = tokens[j].replace(/\s+/g, "").replace("%", "");
        if (/^[0-9]+(?:\.[0-9]+)?$/.test(pct)) {
          rate = Number(pct);
          continue;
        }
      }

      if (amount === "" && /亿元$/.test(tokens[j])) {
        var amt = tokens[j].replace(/\s+/g, "").replace("亿元", "");
        if (/^[0-9]+(?:\.[0-9]+)?$/.test(amt)) {
          amount = Number(amt);
        }
      }

      if (rate != null && amount !== "") break;
    }

    if (rate != null) break;
  }

  // 兜底
  if (rate == null) {
    var mFlat = flat.match(/逆回购操作情况[\s\S]{0,120}?7\s*天[\s\S]{0,20}?([0-9]+(?:\.[0-9]+)?)\s*%/);
    if (mFlat) rate = Number(mFlat[1]);
  }

  if (rate == null) return null;

  return {
    date: date,
    rate: rate,
    amount: amount
  };
}

function prParseMlfDetail_(html) {
  if (!html) return null;

  var text = prNormalizeWhitespace_(prStripHtml_(html));
  var date = prExtractArticleDate_(text);
  var amount = "";
  var rate = null;

  var ma = text.match(/开展\s*([0-9]+(?:\.[0-9]+)?)\s*亿元中期借贷便利(?:（MLF）)?操作/);
  if (!ma) {
    ma = text.match(/([0-9]+(?:\.[0-9]+)?)\s*亿元中期借贷便利(?:（MLF）)?操作/);
  }
  if (ma) amount = Number(ma[1]);

  var mr = text.match(/期限\s*1\s*年[^。；;\n]*?中标利率\s*([0-9]+(?:\.[0-9]+)?)\s*%/);
  if (!mr) {
    mr = text.match(/1\s*年期[^。；;\n]*?中标利率\s*([0-9]+(?:\.[0-9]+)?)\s*%/);
  }
  if (!mr) {
    mr = text.match(/期限\s*1\s*年[^。；;\n]*?利率(?:为)?\s*([0-9]+(?:\.[0-9]+)?)\s*%/);
  }
  if (mr) rate = Number(mr[1]);

  if (rate == null) return null;

  return {
    date: date,
    amount: amount,
    rate: rate
  };
}

function prParseLprDetail_(html) {
  if (!html) return null;

  var text = prNormalizeWhitespace_(prStripHtml_(html));
  var date = prExtractArticleDate_(text);
  var lpr1 = "";
  var lpr5 = "";

  var m1 = text.match(/1\s*年期LPR为[:：]?\s*([0-9]+(?:\.[0-9]+)?)\s*%/);
  if (!m1) {
    m1 = text.match(/1\s*年期贷款市场报价利率为[:：]?\s*([0-9]+(?:\.[0-9]+)?)\s*%/);
  }
  if (m1) lpr1 = Number(m1[1]);

  var m5 = text.match(/5\s*年期以上LPR为[:：]?\s*([0-9]+(?:\.[0-9]+)?)\s*%/);
  if (!m5) {
    m5 = text.match(/5\s*年期以上贷款市场报价利率为[:：]?\s*([0-9]+(?:\.[0-9]+)?)\s*%/);
  }
  if (m5) lpr5 = Number(m5[1]);

  if (lpr1 === "" && lpr5 === "") return null;

  return {
    date: date,
    lpr_1y: lpr1,
    lpr_5y: lpr5
  };
}

/* =========================
 * 通用抓取/解析
 * ========================= */

function prFetchHtml_(url) {
  var lastError = null;

  for (var i = 0; i < 3; i++) {
    try {
      var resp = UrlFetchApp.fetch(url, {
        method: "get",
        muteHttpExceptions: true,
        followRedirects: true,
        headers: {
          "User-Agent": "Mozilla/5.0 (compatible; GAS policy-rate-bot/1.0)"
        }
      });

      var code = resp.getResponseCode();
      if (code >= 200 && code < 300) {
        return resp.getContentText("UTF-8");
      }
      lastError = new Error("请求失败 code=" + code + ", url=" + url);
    } catch (e) {
      lastError = e;
    }

    Utilities.sleep(400 + i * 400);
  }

  throw lastError || new Error("请求失败: " + url);
}

function prExtractAnchorsInOrder_(html, baseUrl) {
  var out = [];
  var re = /<a\b[^>]*href="([^"]+)"[^>]*>([\s\S]*?)<\/a>/ig;
  var m;

  while ((m = re.exec(html)) !== null) {
    var href = prHtmlDecode_(m[1]);
    var inner = m[2];
    var title = prCleanText_(prHtmlDecode_(prStripHtml_(inner)));
    if (!href || !title) continue;

    out.push({
      title: title,
      url: prToAbsoluteUrl_(href, baseUrl)
    });
  }

  return out;
}

function prExtractArticleDate_(text) {
  if (!text) return "";

  var m = text.match(/文章来源[:：]?\s*(\d{4}-\d{2}-\d{2})/);
  if (m) return m[1];

  m = text.match(/(20\d{2})年(\d{1,2})月(\d{1,2})日/);
  if (m) {
    return m[1] + "-" + prPad2_(Number(m[2])) + "-" + prPad2_(Number(m[3]));
  }

  return "";
}

function prExtractSeqNo_(title) {
  var m = String(title || "").match(/第(\d+)号/);
  return m ? Number(m[1]) : 0;
}

/* =========================
 * 表读写
 * ========================= */

function prEnsureRawPolicyRateSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_POLICY_RATE_RAW);
  if (!sheet) sheet = ss.insertSheet(SHEET_POLICY_RATE_RAW);

  var needInit = false;

  if (sheet.getLastRow() < 1) {
    needInit = true;
  } else {
    var existing = sheet.getRange(1, 1, 1, RAW_POLICY_RATE_HEADERS.length).getValues()[0];
    for (var i = 0; i < RAW_POLICY_RATE_HEADERS.length; i++) {
      if (String(existing[i] || "") !== RAW_POLICY_RATE_HEADERS[i]) {
        needInit = true;
        break;
      }
    }
  }

  if (needInit) {
    sheet.clear();
    sheet.getRange(1, 1, 1, RAW_POLICY_RATE_HEADERS.length).setValues([RAW_POLICY_RATE_HEADERS]);
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function prReadPolicyRateRows_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var vals = sheet.getRange(2, 1, lastRow - 1, RAW_POLICY_RATE_HEADERS.length).getValues();
  var rows = [];

  for (var i = 0; i < vals.length; i++) {
    var v = vals[i];
    if (!v[0]) continue;

    rows.push({
      _rowNum: i + 2,
      date: String(v[0] || ""),
      type: String(v[1] || ""),
      term: String(v[2] || ""),
      rate: v[3],
      amount: v[4],
      source: String(v[5] || ""),
      fetched_at: v[6],
      note: String(v[7] || "")
    });
  }

  rows.sort(function(a, b) {
    if (a.date !== b.date) return a.date < b.date ? -1 : 1;
    if (a.type !== b.type) return a.type < b.type ? -1 : 1;
    return a.term < b.term ? -1 : a.term > b.term ? 1 : 0;
  });

  return rows;
}

function prFindPolicyRateEventRow_(sheet, date, type, term) {
  var rows = prReadPolicyRateRows_(sheet);
  for (var i = 0; i < rows.length; i++) {
    if (rows[i].date === date && rows[i].type === type && rows[i].term === term) {
      return rows[i]._rowNum;
    }
  }
  return null;
}

function prUpsertPolicyRateEvent_(sheet, rowObj) {
  if (!rowObj || !rowObj.date || !rowObj.type || !rowObj.term) return;

  var rowNum = prFindPolicyRateEventRow_(sheet, rowObj.date, rowObj.type, rowObj.term);

  var arr = [[
    rowObj.date,
    rowObj.type,
    rowObj.term,
    rowObj.rate,
    rowObj.amount,
    rowObj.source,
    rowObj.fetched_at,
    rowObj.note
  ]];

  if (rowNum) {
    sheet.getRange(rowNum, 1, 1, RAW_POLICY_RATE_HEADERS.length).setValues(arr);
  } else {
    sheet.getRange(sheet.getLastRow() + 1, 1, 1, RAW_POLICY_RATE_HEADERS.length).setValues(arr);
  }
}

function prSortRawPolicyRateSheet_(sheet) {
  var lastRow = sheet.getLastRow();
  if (lastRow <= 2) return;

  sheet.getRange(2, 1, lastRow - 1, RAW_POLICY_RATE_HEADERS.length)
    .sort([
      { column: 1, ascending: true },
      { column: 2, ascending: true },
      { column: 3, ascending: true }
    ]);
}

/* =========================
 * 工具函数
 * ========================= */

function prDedupePolicyRateEvents_(events) {
  var out = [];
  var seen = {};

  for (var i = 0; i < events.length; i++) {
    var e = events[i];
    var key = e.date + "||" + e.type + "||" + e.term;
    if (seen[key]) continue;
    seen[key] = true;
    out.push(e);
  }

  return out;
}

function prFormatPolicyDateTime_(d) {
  return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
}

function prPad2_(n) {
  return n < 10 ? "0" + n : String(n);
}

function prNormalizePolicyTextKeepLines_(s) {
  return String(s || "")
    .replace(/\r/g, "\n")
    .replace(/\u00A0/g, " ")
    .replace(/[ \t\f\v]+/g, " ")
    .replace(/\n+/g, "\n")
    .trim();
}

function prCleanText_(s) {
  return String(s || "")
    .replace(/\u00A0/g, " ")
    .replace(/\s+/g, " ")
    .trim();
}

function prNormalizeWhitespace_(s) {
  return prCleanText_(s);
}

function prStripHtml_(html) {
  return String(html || "")
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<\/(p|div|tr|td|th|table|h1|h2|h3|h4|h5|li)\s*>/gi, "\n")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/gi, " ")
    .replace(/&#160;/gi, " ")
    .replace(/&ensp;|&emsp;/gi, " ");
}

function prHtmlDecode_(s) {
  return String(s || "")
    .replace(/&amp;/g, "&")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&quot;/g, "\"")
    .replace(/&#39;/g, "'")
    .replace(/&#x2F;/g, "/");
}

function prToAbsoluteUrl_(href, baseUrl) {
  if (/^https?:\/\//i.test(href)) return href;

  var originMatch = baseUrl.match(/^(https?:\/\/[^\/]+)/i);
  var origin = originMatch ? originMatch[1] : "";
  var base = baseUrl.replace(/\/[^\/]*$/, "/");

  if (href.indexOf("/") === 0) return origin + href;
  return base + href.replace(/^\.\//, "");
}