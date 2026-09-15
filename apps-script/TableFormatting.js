// =====================================================================
// 讓 Kathy1 / Lin1 / Min1 分頁的「狀態」「類型」欄位看起來接近設計師自己
// 表格（Google Sheets「表格」功能）的彩色標籤＋下拉選單效果。
//
// 為什麼不直接用「轉換為表格」：那個功能會強制把整張表的第 1 列當表格
// 標題列，而這三個分頁的第 1 列是同步腳本用來放「最後同步」橫幅的，
// 真正的欄位標題在第 2 列 —— 轉換後標題會變成同步橫幅、原本的欄位標題
// 反而變成資料的第一列，版面全亂。要修就得把橫幅搬到別的地方，牽動
// SYNC_BANNER_ROW / SYNC_DATA_START_ROW / DASHBOARD_HEADER_ROW 這些跨
// 檔案共用的列號設定，風險較高。
//
// 這裡改用「資料驗證下拉選單」＋「條件式格式」達到類似的視覺效果：
// - 不改動任何列號設定，跟同步／樞紐分析表/ dashboard 完全獨立
// - 條件式格式規則、資料驗證規則都是附掛在 Range 上的獨立物件，同步腳本
//   的 setValues()/setBackgrounds() 不會清掉它們，所以每 5 分鐘同步一次
//   也不會把顏色沖掉
// - 下拉選單清單、顏色都是「從目前資料裡實際出現過的值」動態算出來的，
//   之後這些分頁若出現新的狀態／類型文字，只要重新執行一次這個函式
//   就會自動納入
// =====================================================================

const CHIP_PALETTE = [
  { bg: '#e6f4ea', fg: '#137333' }, // 綠
  { bg: '#e8f0fe', fg: '#1a73e8' }, // 藍
  { bg: '#fef7e0', fg: '#a37800' }, // 黃
  { bg: '#fce8e6', fg: '#c5221f' }, // 紅
  { bg: '#f3e8fd', fg: '#8430ce' }, // 紫
  { bg: '#e0f2f1', fg: '#00695c' }, // 青
  { bg: '#fce8f3', fg: '#c4257c' }, // 粉
  { bg: '#fdf2e9', fg: '#b06000' }, // 橘
  { bg: '#f1f3f4', fg: '#5f6368' }  // 灰（備用）
];

// Google Sheets「表格」功能的彩色標籤顏色只是介面動態算出來顯示的效果，
// 不是存在儲存格裡的真實格式（實測過：用 Apps Script 讀 Kathy 獨立表格
// 那些儲存格的真實背景／字色，讀出來是白底黑字，跟畫面上看到的完全不
// 一樣），沒辦法用程式直接複製。這裡改成照著使用者截圖手動比對出來的
// 顏色，針對「狀態」欄位裡每一個具體的值指定顏色（而不是照出現順序輪
// 流套色票），盡量貼近 Kathy 表格下拉選單裡看到的樣子：
// - A 系列（審核通過的分享率門檻）用綠色，分享率門檻越高顏色越深
// - B / C 系列用藍色，越後面的階段顏色越深
// - 還在討論／退件的用黃色、紅色
// 之後如果 Kathy 的表格改了顏色，或狀態文字微調對不起來，把下面的值或
// 顏色改掉即可；沒列在這裡的新狀態值會自動 fallback 用 CHIP_PALETTE
// 依出現順序輪流上色。
const STATUS_COLOR_MAP = {
  'A1 In Progress':                                  { bg: '#f1f3f4', fg: '#5f6368' },
  'A2 Produced, Delivered & Slacked AM/BD':           { bg: '#fef7e0', fg: '#a37800' },
  'A3 Rejected after follow-up (with reasons)':       { bg: '#fce8e6', fg: '#c5221f' },
  'A4 Approved(share < 10%)':                         { bg: '#e6f4ea', fg: '#137333' },
  'A5 Approved + share > 10%':                        { bg: '#ceead6', fg: '#137333' },
  'A6 Approved + share > 25%':                        { bg: '#a8dab5', fg: '#0d652d' },
  'A7 Approved + share > 50%':                        { bg: '#81c995', fg: '#0d652d' },
  'A8 Approved + share > 70%':                        { bg: '#34a853', fg: '#ffffff' },
  'B1 Discovery Delivered':                           { bg: '#e8f0fe', fg: '#1a73e8' },
  'B2 Added to AA Pipeline':                          { bg: '#aecbfa', fg: '#1967d2' },
  'C1 Pitch Creatives Delivered':                     { bg: '#fef7e0', fg: '#a37800' },
  'C2 Client Committed to Online Campaign':           { bg: '#4285f4', fg: '#ffffff' },
  'Fail':                                             { bg: '#fce8e6', fg: '#c5221f' }
};

// 資料驗證下拉選單往下多留的緩衝列數，讓同步之後新增的任務列也自動有
// 下拉選單可用，不用每次都重新執行這個函式。
const CHIP_ROW_BUFFER = 200;

/**
 * 對單一分頁的「狀態」「類型」欄位套用下拉選單＋依值上色。
 * 可重複執行 —— 每次都會先清掉這兩欄舊的條件式格式規則再重建，
 * 不會疊加出重複規則。
 */
function applyStatusTypeChips(sheetName) {
  const ss = SpreadsheetApp.openById(DESIGNER_CALCULATOR_ID);
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    Logger.log(`⚠️ 找不到分頁：${sheetName}`);
    return;
  }

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < DASHBOARD_FIRST_DATA_ROW || lastCol < 1) {
    Logger.log(`⚠️ ${sheetName} 沒有資料`);
    return;
  }

  const header = sheet.getRange(DASHBOARD_HEADER_ROW, 1, 1, lastCol).getValues()[0];
  const statusCol = findColumn(header, ['狀態'], -1);
  const typeCol = findColumn(header, ['類型'], -1);

  if (statusCol === -1 && typeCol === -1) {
    Logger.log(`⚠️ ${sheetName} 找不到「狀態」或「類型」欄位`);
    return;
  }

  const validationLastRow = lastRow + CHIP_ROW_BUFFER;

  [
    { col: statusCol, label: '狀態' },
    { col: typeCol, label: '類型' }
  ].forEach(({ col, label }) => {
    if (col === -1) return;
    applyChipsToColumn(sheet, col, lastRow, validationLastRow, label);
  });

  Logger.log(`✅ ${sheetName} 的狀態／類型欄位已套上下拉選單＋顏色`);
}

function applyChipsToColumn(sheet, zeroBasedCol, lastRow, validationLastRow, label) {
  const col1 = zeroBasedCol + 1; // 轉成 1-based 供 getRange 使用
  const dataRange = sheet.getRange(DASHBOARD_FIRST_DATA_ROW, col1, lastRow - DASHBOARD_FIRST_DATA_ROW + 1, 1);
  const values = dataRange.getValues().flat().map(v => String(v || '').trim()).filter(v => v);

  const uniqueValues = [];
  const seen = {};
  values.forEach(v => {
    if (!seen[v]) {
      seen[v] = true;
      uniqueValues.push(v);
    }
  });

  if (!uniqueValues.length) {
    Logger.log(`（${label} 欄目前沒有資料，略過）`);
    return;
  }

  // --- 下拉選單（含緩衝列，讓未來新增的任務列也有下拉可用）---
  const validationRange = sheet.getRange(
    DASHBOARD_FIRST_DATA_ROW, col1,
    validationLastRow - DASHBOARD_FIRST_DATA_ROW + 1, 1
  );
  const rule = SpreadsheetApp.newDataValidation()
    .requireValueInList(uniqueValues, true)
    .setAllowInvalid(true) // 同步腳本寫回來源新值時不會被擋下
    .build();
  validationRange.setDataValidation(rule);

  // --- 依值上色（條件式格式）---
  const fullColumnRange = sheet.getRange(DASHBOARD_FIRST_DATA_ROW, col1, validationLastRow - DASHBOARD_FIRST_DATA_ROW + 1, 1);
  const existingRules = sheet.getConditionalFormatRules();
  const keptRules = existingRules.filter(r => {
    return !r.getRanges().some(rg =>
      rg.getSheet().getSheetId() === sheet.getSheetId() &&
      rg.getColumn() === col1 &&
      rg.getNumColumns() === 1
    );
  });

  let fallbackIndex = 0;
  const newRules = uniqueValues.map(value => {
    // 「狀態」欄優先用手動比對過的顏色；沒有對到的（或「類型」欄）就照
    // 出現順序輪流套色票。
    const color = (label === '狀態' && STATUS_COLOR_MAP[value])
      || CHIP_PALETTE[fallbackIndex++ % CHIP_PALETTE.length];
    return SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(value)
      .setBackground(color.bg)
      .setFontColor(color.fg)
      .setBold(true)
      .setRanges([fullColumnRange])
      .build();
  });

  sheet.setConditionalFormatRules(keptRules.concat(newRules));
  Logger.log(`  ${label}：${uniqueValues.length} 種值套用了下拉選單＋顏色`);
}

/** 依序對 Kathy1 / Lin1 / Min1 套用。 */
function applyStatusTypeChipsAll() {
  ['Kathy1', 'Lin1', 'Min1'].forEach(applyStatusTypeChips);
}
