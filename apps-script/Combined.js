// ========================================
// 📝 三人合併分頁設定
// ========================================
// 目標分頁名稱。這支程式會自己建立這個分頁（不存在的話），
// 並且完全由程式管理內容 —— 不要手動在裡面加東西，每次同步都會被覆蓋。
const COMBINED_TARGET_SHEET = 'All';

// 最前面那一欄的標題，用來標示每一列是誰的資料。
const COMBINED_LABEL_HEADER = 'Designer';

// 依哪一欄日期排序（1-based，以「來源表」的欄號為準，不含程式加上去的
// Designer 欄）。5 = E 欄 Start Date —— 用開始日而不是 F 欄 End Date，是
// 因為進行中的案子還沒有結束日，照 End Date 排會把它們全部擠到最下面。
const COMBINED_SORT_COLUMN = 5;

// true = 新的在上（遞減）。日期空白或看不懂的列一律排在最下面。
const COMBINED_SORT_DESCENDING = true;

// 只保留這一季的資料。填 'AUTO' 就會自動跟著當季走（10/1 一到就換成
// 2026Q4）—— 這裡刻意寫死 '2026Q3'，不然季初那幾天分頁會突然變空的，
// 看起來像壞掉。換季時手動改這一行就好。
const COMBINED_QUARTER = '2026Q3';

// 季度欄的標題。程式會先照標題找欄位，找不到才退回用 COMBINED_QUARTER_COLUMN，
// 這樣來源表中間插一欄也不會篩錯。
const COMBINED_QUARTER_HEADER = '季度';
const COMBINED_QUARTER_COLUMN = 9;   // 退路：I 欄

// 同步間隔（分鐘），跟 importWithFormat 一樣預設 5 分鐘。
const COMBINED_SYNC_INTERVAL_MINUTES = 5;

// 上一次寫入內容的雜湊值存放位置，用來判斷「沒變動就跳過」。
const COMBINED_HASH_PROPERTY = 'combinedSyncHash';

// ========================================
// 以下不需要修改
// ========================================

/**
 * Every attribute copied from source to merged tab, as a read/write pair.
 * Listing them once keeps the read, the pad, the change-detection signature
 * and the write from drifting apart as attributes are added or removed.
 *
 * 'values' is the odd one out twice over: it reads display values but writes
 * plain values (matching importWithFormat, so the merged tab shows exactly
 * what the per-designer tabs show), and it pads with a blank rather than by
 * repeating the row's last cell.
 */
const COMBINED_ATTRIBUTES = [
  { key: 'values',               read: 'getDisplayValues',        write: 'setValues' },
  { key: 'backgrounds',          read: 'getBackgrounds',          write: 'setBackgrounds' },
  { key: 'fontColors',           read: 'getFontColors',           write: 'setFontColors' },
  { key: 'fontWeights',          read: 'getFontWeights',          write: 'setFontWeights' },
  { key: 'fontSizes',            read: 'getFontSizes',            write: 'setFontSizes' },
  { key: 'horizontalAlignments', read: 'getHorizontalAlignments', write: 'setHorizontalAlignments' },
  { key: 'verticalAlignments',   read: 'getVerticalAlignments',   write: 'setVerticalAlignments' },
  { key: 'wrapStrategies',       read: 'getWrapStrategies',       write: 'setWrapStrategies' }
];

/**
 * Builds one merged tab holding all three designers' List sheets, filtered to
 * one quarter, sorted by start date with the designers interleaved, and a
 * leading column naming whose row each one is.
 *
 * Filtering on the sheets' own 季度 column rather than on a date range means
 * a task that straddles a quarter boundary lands wherever it was actually
 * booked for scoring, instead of wherever its start date happens to fall.
 *
 * Reads the three source spreadsheets directly (the same ones SOURCES in
 * Code (DO NOT EDIT).js lists) rather than concatenating the already-synced
 * Kathy/Lin/Min tabs, so this tab is never a copy-of-a-copy running one
 * extra sync interval behind the source.
 */
function syncCombinedSheet() {
  const errors = [];
  const blocks = [];

  // SOURCES lives in Code (DO NOT EDIT).js. Apps Script shares one global
  // scope across all files in the project, and this runs at call time (long
  // after every file has been evaluated), so the reference is safe.
  SOURCES.forEach(source => {
    try {
      const sourceSheet = SpreadsheetApp.openById(source.id).getSheetByName(source.sheet);
      if (!sourceSheet) {
        Logger.log(`❌ 找不到來源工作表：${source.sheet}`);
        errors.push(`${COMBINED_TARGET_SHEET} ← ${source.targetSheet}：找不到來源工作表 "${source.sheet}"`);
        return;
      }

      const lastRow = sourceSheet.getLastRow();
      const lastCol = sourceSheet.getLastColumn();
      if (lastRow < 1 || lastCol < 1) {
        Logger.log(`⚠️ ${source.sheet} 沒有數據`);
        return;
      }

      const range = sourceSheet.getRange(1, 1, lastRow, lastCol);
      // Display values are what gets written, but the sort needs the real
      // underlying values — a date cell's display text depends on whatever
      // number format that sheet happens to use.
      const block = { designer: source.targetSheet, cols: lastCol, rawValues: range.getValues() };
      COMBINED_ATTRIBUTES.forEach(attr => { block[attr.key] = range[attr.read](); });
      blocks.push(block);

    } catch (e) {
      Logger.log(`❌ ${source.sheet} 發生錯誤：${e.message}`);
      errors.push(`${COMBINED_TARGET_SHEET} ← ${source.targetSheet} (${source.sheet})：${e.message}`);
    }
  });

  // Reuses the throttled failure mail from Code (DO NOT EDIT).js so a broken
  // merge is as visible as a broken per-designer sync.
  if (errors.length) notifySyncFailure(errors);
  if (!blocks.length) {
    Logger.log('⚠️ 沒有任何來源可用，合併分頁保持原狀');
    return;
  }

  // Interleaving the designers means every row has to be the same width, so
  // narrower sources get padded out to the widest one. The header likewise
  // comes from the widest source, so it always spans every column written.
  const width = blocks.reduce((max, block) => Math.max(max, block.cols), 0);
  const headerBlock = blocks.filter(block => block.cols === width)[0];

  const quarter = combinedTargetQuarter();
  const body = [];
  blocks.forEach(block => {
    const quarterIndex = combinedQuarterIndex(block.values[0]);
    for (let i = 1; i < block.values.length; i++) {
      const rowQuarter = String(block.values[i][quarterIndex] || '').trim().toUpperCase();
      if (rowQuarter !== quarter) continue;
      body.push({
        sortKey: combinedSortKey(block.rawValues[i], block.values[i]),
        cells: buildCombinedRow(block, i, block.designer, width)
      });
    }
  });

  if (!body.length) {
    Logger.log(`⚠️ 沒有任何 ${quarter} 的資料，${COMBINED_TARGET_SHEET} 只會有標題列`);
  }

  // Array.prototype.sort is stable, so rows sharing a start date — and the
  // undated rows parked at the bottom — keep their original per-designer
  // order instead of shuffling on every sync.
  body.sort((a, b) => {
    if (a.sortKey === null && b.sortKey === null) return 0;
    if (a.sortKey === null) return 1;
    if (b.sortKey === null) return -1;
    return COMBINED_SORT_DESCENDING ? b.sortKey - a.sortKey : a.sortKey - b.sortKey;
  });

  const rows = [buildCombinedRow(headerBlock, 0, COMBINED_LABEL_HEADER, width)]
    .concat(body.map(entry => entry.cells));

  const grid = {};
  COMBINED_ATTRIBUTES.forEach(attr => { grid[attr.key] = rows.map(row => row[attr.key]); });

  // Skip the rewrite when nothing changed, the same way importWithFormat
  // does — this runs every few minutes and a no-op clear()+rewrite would
  // otherwise churn the sheet (and any pivot table reading it) constantly.
  // Wrap strategies are left out of the signature: they come back as enum
  // objects that don't survive JSON.stringify, and they never change on
  // their own without some other attribute changing too.
  const signature = Utilities.base64Encode(
    Utilities.computeDigest(
      Utilities.DigestAlgorithm.MD5,
      JSON.stringify(
        COMBINED_ATTRIBUTES
          .filter(attr => attr.key !== 'wrapStrategies')
          .map(attr => grid[attr.key])
      )
    )
  );

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const props = PropertiesService.getScriptProperties();
  let target = spreadsheet.getSheetByName(COMBINED_TARGET_SHEET);

  if (target && props.getProperty(COMBINED_HASH_PROPERTY) === signature) {
    Logger.log(`✅ ${COMBINED_TARGET_SHEET} 無變動，跳過`);
    return;
  }

  if (!target) target = spreadsheet.insertSheet(COMBINED_TARGET_SHEET);

  const totalRows = rows.length;
  const totalCols = width + 1;
  if (target.getMaxRows() < totalRows) {
    target.insertRowsAfter(target.getMaxRows(), totalRows - target.getMaxRows());
  }
  if (target.getMaxColumns() < totalCols) {
    target.insertColumnsAfter(target.getMaxColumns(), totalCols - target.getMaxColumns());
  }

  // Full clear rather than an overwrite: a shrinking source (rows deleted in
  // someone's List) would otherwise leave stale rows and their colours
  // behind below the new content.
  target.clear();

  const targetRange = target.getRange(1, 1, totalRows, totalCols);
  COMBINED_ATTRIBUTES.forEach(attr => { targetRange[attr.write](grid[attr.key]); });
  target.setFrozenRows(1);

  props.setProperty(COMBINED_HASH_PROPERTY, signature);
  Logger.log(`✅ ${COMBINED_TARGET_SHEET} 更新完成！${quarter} 共 ${totalRows - 1} 筆（${totalRows} 行 x ${totalCols} 欄）`);
}

/** 要保留的季度，例如 '2026Q3'。COMBINED_QUARTER 設成 'AUTO' 時依當下日期推算。 */
function combinedTargetQuarter() {
  if (String(COMBINED_QUARTER).toUpperCase() !== 'AUTO') {
    return String(COMBINED_QUARTER).trim().toUpperCase();
  }

  const timeZone = Session.getScriptTimeZone();
  const now = new Date();
  const year = Utilities.formatDate(now, timeZone, 'yyyy');
  const month = Number(Utilities.formatDate(now, timeZone, 'MM'));
  return `${year}Q${Math.ceil(month / 3)}`;
}

/** 季度欄在來源表裡的索引，照標題找；找不到就退回設定的欄號。 */
function combinedQuarterIndex(headerRow) {
  for (let i = 0; i < headerRow.length; i++) {
    if (String(headerRow[i]).trim() === COMBINED_QUARTER_HEADER) return i;
  }
  return COMBINED_QUARTER_COLUMN - 1;
}

/**
 * Turns row `index` of a source block into one merged-tab row: padded out to
 * `width` source columns, with the designer label pushed on the front.
 *
 * Padding and the label cell both reuse values that came back from a get*()
 * call on a real cell — the row's last cell for the padding, its first cell
 * for the label — so every value handed to the matching set*() is one the
 * API is already known to accept, and the label sits inside the row's own
 * banding instead of punching a white gap down the left edge.
 */
function buildCombinedRow(block, index, label, width) {
  const row = {};

  COMBINED_ATTRIBUTES.forEach(attr => {
    const cells = block[attr.key][index].slice(0, width);
    const filler = attr.key === 'values' ? '' : cells[cells.length - 1];
    while (cells.length < width) cells.push(filler);
    row[attr.key] = [attr.key === 'values' ? label : cells[0]].concat(cells);
  });

  return row;
}

/**
 * Sort key for one row, in epoch millis, or null when the row has no usable
 * date (blank, "TBD", and anything else unparseable — those sort to the
 * bottom rather than pretending to be some particular date).
 */
function combinedSortKey(rawRow, displayRow) {
  const index = COMBINED_SORT_COLUMN - 1;

  // normalizeDate() lives in Untitled.js — same project, same global scope.
  // It already handles Date objects, "2026年9月1日", and plain date strings.
  const iso = normalizeDate(rawRow[index] || displayRow[index]);
  if (!iso) return null;

  const time = new Date(`${iso}T00:00:00`).getTime();
  return isNaN(time) ? null : time;
}

/**
 * Run once manually from the Apps Script editor to install the merge
 * trigger. Safe to re-run — only clears a prior syncCombinedSheet trigger,
 * so it never touches importWithFormat's or refreshDashboardCache's.
 */
function setupCombinedTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'syncCombinedSheet')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('syncCombinedSheet')
    .timeBased()
    .everyMinutes(COMBINED_SYNC_INTERVAL_MINUTES)
    .create();
  Logger.log(`觸發器設定完成！每 ${COMBINED_SYNC_INTERVAL_MINUTES} 分鐘合併一次`);
}
