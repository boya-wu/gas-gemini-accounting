/**
 * 更新進階 Dashboard：修正變動支出計算邏輯
 */
function setupAdvancedDashboard() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let dashSheet = ss.getSheetByName('Dashboard');
  if (!dashSheet) dashSheet = ss.insertSheet('Dashboard');
  dashSheet.clear();

  // 1. 設定標題
  dashSheet.getRange("A1").setValue("📅 統計月份:").setFontWeight("bold");
  dashSheet.getRange("B1").setValue(Utilities.formatDate(new Date(), "GMT+8", "yyyy-MM"));

  // 2. 預算看板結構
  const budgetData = [
    ["💰 預算看板", "吳 (Wu)", "巫 (Witch)"],
    ["本月總收入 (A)", "", ""],
    ["預扣固定支出 (B)", "", ""],
    ["變動生活支出 (C)", "", ""],
    ["本月可用餘額", "", ""]
  ];
  dashSheet.getRange("A3:C7").setValues(budgetData).setBorder(true, true, true, true, true, true);
  dashSheet.getRange("A3:C3").setBackground("#e6f4ea").setFontWeight("bold");

  // --- 填入修正後的公式 ---

  // A. 收入 (FixedItems 抓取)
  dashSheet.getRange("B4").setFormula('=SUMIFS(FixedItems!C:C, FixedItems!A:A, "收入", FixedItems!D:D, "吳")');
  dashSheet.getRange("C4").setFormula('=SUMIFS(FixedItems!C:C, FixedItems!A:A, "收入", FixedItems!D:D, "巫")');

  // B. 固定支出分攤 (房租、保險、油錢)
  const fixedCost = 'SUMIFS(FixedItems!C:C, FixedItems!A:A, "固定支出", FixedItems!D:D, "';
  const annualCost = 'SUMIFS(FixedItems!C:C, FixedItems!A:A, "年度分攤", FixedItems!D:D, "';
  
  dashSheet.getRange("B5").setFormula(`=${fixedCost}吳") + (${fixedCost}共同")/2) + (${annualCost}吳")/12) + (${annualCost}共同")/24)`);
  dashSheet.getRange("C5").setFormula(`=${fixedCost}巫") + (${fixedCost}共同")/2) + (${annualCost}巫")/12) + (${annualCost}共同")/24)`);

  // C. 變動生活支出 (修正後的公式：使用 LineItems!F 欄月份輔助)
  // 邏輯：(個人品項總和) + (共同品項總和 / 2)
  const varWu = '=SUMIFS(LineItems!C:C, LineItems!E:E, "吳", LineItems!F:F, $B$1) + (SUMIFS(LineItems!C:C, LineItems!E:E, "共同", LineItems!F:F, $B$1) / 2)';
  const varWitch = '=SUMIFS(LineItems!C:C, LineItems!E:E, "巫", LineItems!F:F, $B$1) + (SUMIFS(LineItems!C:C, LineItems!E:E, "共同", LineItems!F:F, $B$1) / 2)';
  
  dashSheet.getRange("B6").setFormula(varWu);
  dashSheet.getRange("C6").setFormula(varWitch);

  // D. 餘額
  dashSheet.getRange("B7").setFormula('=B4-B5-B6');
  dashSheet.getRange("C7").setFormula('=C4-C5-C6');
  dashSheet.getRange("B7:C7").setFontWeight("bold").setBackground("#fff2cc");

  // 3. 墊付結算區塊
  const settleData = [
    ["🤝 墊付結算表", "吳 (Wu)", "巫 (Witch)"],
    ["實際墊付金額", "", ""],
    ["結算結果", "", ""]
  ];
  dashSheet.getRange("A9:C11").setValues(settleData).setBorder(true, true, true, true, true, true);
  
  // 墊付金額：比對 Transactions 裡誰付的錢
  dashSheet.getRange("B10").setFormula('=SUMPRODUCT(Transactions!D:D, (Transactions!E:E="吳") * (TEXT(Transactions!B:B, "yyyy-mm")=$B$1))');
  dashSheet.getRange("C10").setFormula('=SUMPRODUCT(Transactions!D:D, (Transactions!E:E="巫") * (TEXT(Transactions!B:B, "yyyy-mm")=$B$1))');
  
  // 結算結果 (墊付金額 - 應負擔變動支出)
  dashSheet.getRange("B11").setFormula('=B10-B6');
  dashSheet.getRange("C11").setFormula('=C10-C6');

  dashSheet.setColumnWidth(1, 150);
  dashSheet.setColumnWidths(2, 2, 180);
  
  SpreadsheetApp.getUi().alert("✅ Dashboard 公式已修正完畢！");
}