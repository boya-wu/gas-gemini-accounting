const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_KEY');
const GEMINI_URL = `https://generativelanguage.googleapis.com/v1beta/models/gemini-2.0-flash:generateContent?key=${GEMINI_API_KEY}`;

/**
 * 每次開啟試算表時，自動計算一次並更新看板
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('💰 記帳工具')
    .addItem('更新結算數據', 'refreshDashboard')
    .addToUi();
}

function refreshDashboard() {
  // 這裡可以加入強迫公式重新計算的邏輯，或彈出結算視窗
  SpreadsheetApp.getUi().alert('結算看板已更新！');
}

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('智慧記帳系統')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * 主要執行函式
 */
function processUpload(base64Data) {
  try {
    console.log("開始解析圖片...");
    const result = analyzeReceipt(base64Data);
    
    if (result.error) throw new Error(result.error);
    
    console.log("解析成功，準備寫入試算表...");
    const saveMsg = saveDataToSheet(result, "測試使用者"); // 暫時預設為測試使用者
    
    return "✅ 成功！\n" + JSON.stringify(result, null, 2) + "\n\n" + saveMsg;
  } catch (e) {
    console.error("發生錯誤: " + e.toString());
    return "❌ 錯誤: " + e.toString();
  }
}

/**
 * 第一步：僅負責解析圖片
 */
function analyzeReceipt(base64Image) {
  const cleanBase64 = base64Image.replace(/^data:image\/(png|jpg|jpeg);base64,/, "");
  const prompt = "你是一個收據解析助手。請從收據中提取：商店名稱(store)、日期(date)、總額(totalAmount, 數字)、細項(items, 包含 name 和 price)。請嚴格以 JSON 格式回傳。";

  const payload = {
    "contents": [{
      "parts": [
        {"text": prompt},
        {"inline_data": {"mime_type": "image/jpeg", "data": cleanBase64}}
      ]
    }]
  };

  const options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload)
  };

  const response = UrlFetchApp.fetch(GEMINI_URL, options);
  const result = JSON.parse(response.getContentText());
  const aiText = result.candidates[0].content.parts[0].text;
  const cleanJson = aiText.replace(/```json|```/g, '').trim();
  return cleanJson; // 回傳 JSON 字串給前端 renderItems
}

/**
 * 第二步：接收確認後的資料並寫入
 */
function processFinalSave(finalData) {
  const ss = SpreadsheetApp.openById("1YkslWYWACTfb6Z7-s9usNLwAGk0kXwS20xl7wW8Q6pQ");
  const transSheet = ss.getSheetByName('Transactions');
  const itemSheet = ss.getSheetByName('LineItems');
  const transactionId = Utilities.getUuid();

  // 確保日期格式正確
  const dateObj = new Date(finalData.date);
  const payer = finalData.payer || "吳";

  transSheet.appendRow([
    transactionId,
    dateObj,
    finalData.store || "未命名商店",
    Number(finalData.totalAmount),
    payer,
    ""
  ]);

  finalData.items.forEach(item => {
    if (item.name || item.price) { // 避免寫入空白行
      itemSheet.appendRow([
        transactionId,
        item.name || "未命名項目",
        Number(item.price || 0),
        "",
        item.owner || "共同"
      ]);
    }
  });

  return "已成功紀錄到試算表";
}

// analyzeReceipt 與 getBudgetSummary 保持原樣

/**
 * 從 Dashboard 抓取當月預算摘要
 */
function getBudgetSummary() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet(); // 🔹 改用 Active 避免 ID 錯誤
    const dash = ss.getSheetByName('Dashboard');
    
    if (!dash) return { wu: 0, witch: 0, month: "N/A" };

    // 取得資料
    const wuBalance = dash.getRange("B7").getValue();
    const witchBalance = dash.getRange("C7").getValue();
    
    return {
      wu: isNaN(wuBalance) ? 0 : Math.round(wuBalance),
      witch: isNaN(witchBalance) ? 0 : Math.round(witchBalance)
    };
  } catch (e) {
    console.error("抓取預算失敗: " + e.toString());
    return { wu: "Err", witch: "Err" };
  }
}