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
  const errors = [];

  SOURCES.forEach(source => {
    try {
      const sourceSheet = SpreadsheetApp.openById(source.id).getSheetByName(source.sheet);
      const targetSheet = targetSpreadsheet.getSheetByName(source.targetSheet);

      if (!sourceSheet) {
        Logger.log(`❌ 找不到來源工作表：${source.sheet}`);
        errors.push(`${source.targetSheet}：找不到來源工作表 "${source.sheet}"`);
        return;
      }
      if (!targetSheet) {
        Logger.log(`❌ 找不到目標工作表：${source.targetSheet}`);
        errors.push(`${source.targetSheet}：找不到目標工作表`);
        return;
      }

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
      errors.push(`${source.targetSheet} (${source.sheet})：${e.message}`);
    }
  });

  if (errors.length) notifySyncFailure(errors);
}

/**
 * Emails a summary when one or more SOURCES entries fail to sync, so a
 * broken sync doesn't go unnoticed the way it did before (Logger.log alone
 * is only visible to someone who thinks to check the execution log).
 * Throttled to at most one email per hour so a persistent failure across
 * many 5-minute runs doesn't flood the inbox.
 */
function notifySyncFailure(errors) {
  const props = PropertiesService.getScriptProperties();
  const throttleMs = 60 * 60 * 1000;
  const lastNotified = Number(props.getProperty('lastSyncFailureNotification') || 0);
  if (Date.now() - lastNotified < throttleMs) return;

  const recipient = 'celia.xa.lin@appier.com';
  const subject = `⚠️ Designer Calculator 同步失敗（${errors.length} 筆）`;
  const body = [
    '以下來源同步失敗：',
    '',
    ...errors,
    '',
    '請檢查來源試算表的權限或分頁名稱是否有變動。',
    '（此通知每小時最多寄送一次，避免同一個問題持續觸發時灌爆信箱。）'
  ].join('\n');

  try {
    MailApp.sendEmail(recipient, subject, body);
    props.setProperty('lastSyncFailureNotification', String(Date.now()));
  } catch (e) {
    Logger.log(`❌ 無法寄送同步失敗通知：${e.message}`);
  }
}

/**
 * Run once manually from the Apps Script editor to install the sync
 * trigger. Safe to re-run — only clears a prior importWithFormat trigger
 * first, so it never touches refreshDashboardCache's trigger in
 * Untitled.js (the previous version deleted *every* trigger in the
 * project, which silently wiped out unrelated triggers on every re-run).
 */
function setupTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'importWithFormat')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger("importWithFormat")
    .timeBased()
    .everyMinutes(SYNC_INTERVAL_MINUTES)
    .create();
  Logger.log(`觸發器設定完成！每 ${SYNC_INTERVAL_MINUTES} 分鐘同步一次`);
}