// ============================================================
// 115優高計畫書 — 經費概算收件系統
// 獨立 GAS Web App（與計畫書 GAS 分開部署）
// 部署為 Web App：執行身分＝我、存取權限＝所有人（含匿名）
// ============================================================

// 第一次使用前：
// 1. 在 Google 試算表建立一份空白的新試算表
// 2. 複製試算表 ID 填入下方
var BUDGET_SPREADSHEET_ID = "請在此填入新建立的試算表ID";

// 各子計畫的分頁名稱
var PLAN_NAMES = {
  "a1":   "A-1課綱課程",
  "a2":   "A-2課務推動",
  "b1-1": "B1-1數位核心",
  "b1-2": "B1-2數位課程",
  "b1-3": "B1-3數位遠距",
  "b2":   "B2-SEL",
  "b3":   "B3-交通安全",
  "c":    "C-特色發展"
};

// ── 主要進入點 ──────────────────────────────────────────────
function doPost(e) {
  var result = { ok: false, message: "" };
  try {
    var data = JSON.parse(e.postData.contents);
    var subPlan = (data.sub_plan_type || "").toLowerCase();

    if (!PLAN_NAMES[subPlan]) {
      result.message = "不明的子計畫類型：" + subPlan;
      return respond(result);
    }

    var ss = SpreadsheetApp.openById(BUDGET_SPREADSHEET_ID);
    var planLabel = PLAN_NAMES[subPlan];
    var ts = new Date();
    var name = data.name || "";
    var subject = data.subject || "";
    var planName = data.plan_name || "";

    // ── 主表（主要費用項目）──────────────────────────────────
    var mainSheet = getOrCreateSheet(ss, planLabel + "主表",
      ["時間戳記","姓名","科別","計畫名稱","學期","費用類別","項目說明","單位","數量","單價(仟)","金額(仟)"]);
    (data.main_items || []).forEach(function(item) {
      mainSheet.appendRow([
        ts, name, subject, planName,
        item.semester     || "",
        item.category     || "",
        item.description  || "",
        item.unit         || "",
        item.quantity     || 0,
        item.unit_price   || 0,
        item.amount       || 0
      ]);
    });

    // ── 物品費清單 ───────────────────────────────────────────
    var goodsSheet = getOrCreateSheet(ss, planLabel + "物品費",
      ["時間戳記","姓名","科別","計畫名稱","學期","品名規格","單位","數量","單價(仟)","金額(仟)"]);
    (data.goods_items || []).forEach(function(item) {
      goodsSheet.appendRow([
        ts, name, subject, planName,
        item.semester  || "",
        item.name      || "",
        item.unit      || "",
        item.quantity  || 0,
        item.unit_price || 0,
        item.amount    || 0
      ]);
    });

    // ── 資料蒐集費清單 ───────────────────────────────────────
    var dataSheet = getOrCreateSheet(ss, planLabel + "資料蒐集費",
      ["時間戳記","姓名","科別","計畫名稱","學期","品名規格","單位","數量","單價(仟)","金額(仟)"]);
    (data.data_items || []).forEach(function(item) {
      dataSheet.appendRow([
        ts, name, subject, planName,
        item.semester   || "",
        item.name       || "",
        item.unit       || "",
        item.quantity   || 0,
        item.unit_price || 0,
        item.amount     || 0
      ]);
    });

    // ── 總覽表 ──────────────────────────────────────────────
    var masterSheet = getOrCreateSheet(ss, "📋所有提交",
      ["時間戳記","姓名","科別","子計畫","計畫名稱",
       "上學期主表(仟)","上學期物品費(仟)","上學期資蒐費(仟)",
       "下學期主表(仟)","下學期物品費(仟)","下學期資蒐費(仟)","總計(仟)"]);

    // 計算各分項金額
    function sumBySem(items, sem) {
      return (items || []).filter(function(i) {
        return i.semester === sem;
      }).reduce(function(acc, i) {
        return acc + (i.amount || 0);
      }, 0);
    }
    var s1Main  = sumBySem(data.main_items,  "上學期");
    var s2Main  = sumBySem(data.main_items,  "下學期");
    var s1Goods = sumBySem(data.goods_items, "上學期");
    var s2Goods = sumBySem(data.goods_items, "下學期");
    var s1Data  = sumBySem(data.data_items,  "上學期");
    var s2Data  = sumBySem(data.data_items,  "下學期");
    var grand   = s1Main + s2Main + s1Goods + s2Goods + s1Data + s2Data;

    masterSheet.appendRow([
      ts, name, subject, subPlan, planName,
      s1Main, s1Goods, s1Data,
      s2Main, s2Goods, s2Data,
      grand
    ]);

    result.ok = true;
    result.message = "已儲存，總計 " + grand.toFixed(3) + " 仟元";
  } catch (err) {
    result.message = err.toString();
  }
  return respond(result);
}

// GET 請求：回傳狀態（方便測試）
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ ok: true, message: "115優高計畫書經費收件系統運作中" }))
    .setMimeType(ContentService.MimeType.JSON);
}

// ── 輔助函式 ────────────────────────────────────────────────
function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getOrCreateSheet(ss, sheetName, headers) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    sheet.appendRow(headers);
    sheet.getRange(1, 1, 1, headers.length)
         .setBackground("#4472C4").setFontColor("#FFFFFF").setFontWeight("bold");
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// ── 初始化：建立所有分頁 ────────────────────────────────────
function initBudgetSheets() {
  var ss = SpreadsheetApp.openById(BUDGET_SPREADSHEET_ID);
  var baseHeaders = ["時間戳記","姓名","科別","計畫名稱","學期","品名規格","單位","數量","單價(仟)","金額(仟)"];
  var mainHeaders = ["時間戳記","姓名","科別","計畫名稱","學期","費用類別","項目說明","單位","數量","單價(仟)","金額(仟)"];

  Object.values(PLAN_NAMES).forEach(function(label) {
    getOrCreateSheet(ss, label + "主表", mainHeaders);
    getOrCreateSheet(ss, label + "物品費", baseHeaders);
    getOrCreateSheet(ss, label + "資料蒐集費", baseHeaders);
  });

  getOrCreateSheet(ss, "📋所有提交",
    ["時間戳記","姓名","科別","子計畫","計畫名稱",
     "上學期主表(仟)","上學期物品費(仟)","上學期資蒐費(仟)",
     "下學期主表(仟)","下學期物品費(仟)","下學期資蒐費(仟)","總計(仟)"]);

  Logger.log("所有分頁建立完成（共 " + (Object.keys(PLAN_NAMES).length * 3 + 1) + " 個分頁）");
}
