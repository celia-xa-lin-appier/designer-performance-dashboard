// ========================================
// 📝 在這裡填入你的設定
// ========================================
const SOURCES = [
  // --- 第 1 個表格 ---
  {
    id:          "1nfSmY4GeRLuy3YmOXhjiJ1cnfNkXa9Q_LDXyj5mUd-E",
    sheet:       "List",
    targetSheet: "Kathy",
  },
  // --- 第 2 個表格 ---
  {
    id:          "1nfSmY4GeRLuy3YmOXhjiJ1cnfNkXa9Q_LDXyj5mUd-E",
    sheet:       "統計",
    targetSheet: "Kathy's Summary",
  },
  // --- 第 3 個表格 ---
  {
    id:          "1S6WO4uedwmJ2aGpGPVK906ab6x0YO0GoxSbGmGV9uyU",
    sheet:       "List",
    targetSheet: "Lin",
  },
  // --- 第 4 個表格 ---
  {
    id:          "1S6WO4uedwmJ2aGpGPVK906ab6x0YO0GoxSbGmGV9uyU",
    sheet:       "統計",
    targetSheet: "Lin's Summary",
  },
  // --- 第 5 個表格 ---
  {
    id:          "1ZqI-v3RNYPX8s648VAfQxDoDvKaRqWJcvQf1YSDCeRI",
    sheet:       "List",
    targetSheet: "Min",
  },
  // --- 第 6 個表格 ---
  {
    id:          "1ZqI-v3RNYPX8s648VAfQxDoDvKaRqWJcvQf1YSDCeRI",
    sheet:       "統計",
    targetSheet: "Min's Summary",
  },
  
];

// ========================================
// ⏱️ 同步間隔設定（分鐘）
// ========================================
const SYNC_INTERVAL_MINUTES = 5;

// ========================================
// 以下不需要修改
// ========================================

function importWithFormat() {
  const targetSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  SOURCES.forEach(source => {
    try {
      const sourceSheet = SpreadsheetApp.openById(source.id).getSheetByName(source.sheet);
      const targetSheet = targetSpreadsheet.getSheetByName(source.targetSheet);

      if (!sourceSheet) { Logger.log(`❌ 找不到來源工作表：${source.sheet}`); return; }
      if (!targetSheet) { Logger.log(`❌ 找不到目標工作表：${source.targetSheet}`); return; }

      const lastRow = sourceSheet.getLastRow();
      const lastCol = sourceSheet.getLastColumn();
      if (lastRow === 0) { Logger.log(`⚠️ ${source.sheet} 沒有數據`); return; }

      const sourceRange = sourceSheet.getRange(1, 1, lastRow, lastCol);
      const sourceValues = sourceRange.getDisplayValues();
      const targetValues = targetSheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();

      if (JSON.stringify(sourceValues) === JSON.stringify(targetValues)) {
        Logger.log(`✅ ${source.sheet} 無變動，跳過`);
        return;
      }

      const targetRange = targetSheet.getRange(1, 1, lastRow, lastCol);
      targetRange.setValues(sourceValues);
      targetRange.setBackgrounds(sourceRange.getBackgrounds());
      targetRange.setFontColors(sourceRange.getFontColors());
      targetRange.setFontWeights(sourceRange.getFontWeights());
      targetRange.setFontSizes(sourceRange.getFontSizes());
      targetRange.setHorizontalAlignments(sourceRange.getHorizontalAlignments());
      targetRange.setVerticalAlignments(sourceRange.getVerticalAlignments());
      targetRange.setWrapStrategies(sourceRange.getWrapStrategies());

      Logger.log(`✅ ${source.sheet} 更新完成！${lastRow} 行 x ${lastCol} 欄`);

    } catch (e) {
      Logger.log(`❌ ${source.sheet} 發生錯誤：${e.message}`);
    }
  });
}

function setupTrigger() {
  ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
  ScriptApp.newTrigger("importWithFormat")
    .timeBased()
    .everyMinutes(SYNC_INTERVAL_MINUTES)
    .create();
  Logger.log(`觸發器設定完成！每 ${SYNC_INTERVAL_MINUTES} 分鐘同步一次`);
}