// ============================================================
// 115優高計畫書收件系統 - Google Apps Script
// 部署為 Web App：執行身分＝我、存取權限＝所有人（含匿名）
// ============================================================

var SPREADSHEET_ID = "1c64Qq8hOwgFt8Nt5-4_Ay9MGiuuoV-KhsAK49NgX5ww";

// 各子計畫的 Sheet 名稱與欄位標題
var SHEETS = {
  "a1": {
    name: "A-1課綱課程與社群",
    headers: ["時間戳記","姓名","科別","計畫名稱類型","計畫名稱",
              "分支計畫序號","課程類型","核心素養","校本能力指標",
              "課程領域(上)","課程領域(下)","課程面活動","社群面活動",
              "目標值","社群成員","社群活動描述","備註"]
  },
  "a2": {
    name: "A-2課綱課務推動",
    headers: ["時間戳記","姓名","科別","計畫名稱類型","計畫名稱",
              "計畫目標","負責人員","實施規劃與策略","目標值","備註"]
  },
  "b1-1": {
    name: "B1-1數位核心小組",
    headers: ["時間戳記","姓名","科別","計畫名稱類型","計畫名稱",
              "計畫目標","負責人員","實施規劃與策略","目標值","備註"]
  },
  "b1-2": {
    name: "B1-2數位課程與社群",
    headers: ["時間戳記","姓名","科別","計畫名稱類型","計畫名稱",
              "課程類型","課程領域","課程目標","教學平台",
              "課程實施規劃","社群發展策略","目標值","備註"]
  },
  "b1-3": {
    name: "B1-3數位遠距共授",
    headers: ["時間戳記","姓名","科別","計畫名稱類型","計畫名稱",
              "計畫目標","負責人員","實施規劃與策略","目標值","備註"]
  },
  "b2": {
    name: "B2-SEL社會情緒學習",
    headers: ["時間戳記","姓名","科別","計畫名稱類型","計畫名稱",
              "計畫目標","SEL核心小組","社群成員","社群工作項目",
              "融入課程名稱","課程領域","實施策略","目標值","備註"]
  },
  "b3": {
    name: "B3-交通安全教育",
    headers: ["時間戳記","姓名","科別","計畫名稱類型","計畫名稱",
              "計畫目標","負責人員","實施規劃與策略","目標值","備註"]
  },
  "c": {
    name: "C-學校特色發展",
    headers: ["時間戳記","姓名","科別","計畫名稱類型","計畫名稱",
              "計畫目標","負責人員","實施規劃與策略","目標值","備註"]
  }
};

// ── 主要進入點 ──────────────────────────────────────────────
function doPost(e) {
  var result = { ok: false, message: "" };
  try {
    var data = JSON.parse(e.postData.contents);
    var subPlan = (data.sub_plan_type || "").toLowerCase();

    if (!SHEETS[subPlan]) {
      result.message = "不明的子計畫類型：" + subPlan;
      return respond(result);
    }

    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = getOrCreateSheet(ss, subPlan);
    var row = buildRow(subPlan, data);
    sheet.appendRow(row);

    // 同時寫入「所有提交」總表
    var masterSheet = getOrCreateMasterSheet(ss);
    masterSheet.appendRow([
      new Date(),
      data.name || "",
      data.subject || "",
      subPlan,
      data.plan_name || "",
      data.plan_name_type || ""
    ]);

    result.ok = true;
    result.message = "已儲存";
  } catch (err) {
    result.message = err.toString();
  }
  return respond(result);
}

// GET 請求：回傳狀態（方便測試）
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, message: "115優高計畫書收件系統運作中" }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── 輔助函式 ─────────────────────────────────────────────────

function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet(ss, subPlan) {
  var cfg = SHEETS[subPlan];
  var sheet = ss.getSheetByName(cfg.name);
  if (!sheet) {
    sheet = ss.insertSheet(cfg.name);
    sheet.appendRow(cfg.headers);
    sheet.getRange(1, 1, 1, cfg.headers.length)
         .setBackground("#4472C4").setFontColor("#FFFFFF").setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getOrCreateMasterSheet(ss) {
  var name = "📋所有提交";
  var sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name, 0); // 放第一個
    sheet.appendRow(["時間戳記","姓名","科別","子計畫","計畫名稱","課程或社群"]);
    sheet.getRange(1, 1, 1, 6)
         .setBackground("#34A853").setFontColor("#FFFFFF").setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// 依子計畫類型組成試算表的一列
function buildRow(subPlan, d) {
  var ts = new Date();
  var common = [ts, d.name||"", d.subject||"", d.plan_name_type||"", d.plan_name||""];

  if (subPlan === "a1") {
    return common.concat([
      d.course_code || "",
      d.course_type || "",
      arrToStr(d.core_competencies),
      d.ability_indicators || "",
      d.domain_s1 || "",
      d.domain_s2 || "",
      arrToStr(d.course_activities),
      arrToStr(d.community_activities),
      d.goal_description || "",
      d.community_members || "",
      d.community_activity_desc || "",
      d.notes || ""
    ]);
  }
  if (subPlan === "b1-2") {
    return common.concat([
      d.course_type || "",
      d.course_domain || "",
      d.course_goal || "",
      arrToStr(d.teaching_platforms),
      d.implementation || "",
      d.community_strategy || "",
      d.goal_description || "",
      d.notes || ""
    ]);
  }
  if (subPlan === "b2") {
    return common.concat([
      d.plan_goal || "",
      d.sel_core_team || "",
      d.community_members || "",
      arrToStr(d.work_items),
      d.course_name || "",
      d.course_domain || "",
      d.implementation_strategy || "",
      d.goal_description || "",
      d.notes || ""
    ]);
  }
  // 通用（a2, b1-1, b1-3, b3, c）
  return common.concat([
    d.plan_goal || d.goal || "",
    d.members || "",
    d.strategy || "",
    d.goal_description || "",
    d.notes || ""
  ]);
}

function arrToStr(val) {
  if (!val) return "";
  if (Array.isArray(val)) return val.join("、");
  return String(val);
}

// ── 初始化：第一次執行時建立所有 Sheet ──────────────────────
function initSheets() {
  var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  getOrCreateMasterSheet(ss);
  Object.keys(SHEETS).forEach(function(k) { getOrCreateSheet(ss, k); });
  Logger.log("所有 Sheet 建立完成");
}
