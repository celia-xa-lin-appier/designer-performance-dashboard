/**
 * 「Q3 / Q2 Designer status」的樞紐分析表是用「欄位位置」綁定欄位的，所以只要
 * 設計師的來源表插一欄（例如 2026-09-14 多出 playable/video/image count 三欄），
 * 季度篩選和總分值欄就會指到別的欄位，整張表變空白卻不會報錯。
 *
 * auditPivots() 只讀不寫，印出每張表目前綁到哪一欄；
 * fixPivots() 則照下面的設定，用「標題名稱」重新綁定 24 張表，
 * 來源分頁或範圍不對的會就地砍掉重建（Apps Script 沒有 setSourceDataRange）。
 * 來源表以後再插欄，重跑一次 fixPivots() 就好。
 */
const PIVOT_TABS = {
  'Q3 Designer status': '2026Q3',
  'Q2 Designer status': '2026Q2'
};

// 錨點欄位 → 該讀哪個分頁。三個分頁都是同步寫入的那一份（Kathy 分頁 2026-09-14
// 被改名成 Kathy1；Lin / Min 的原始分頁被保護成唯讀，同步早就改寫到 Lin1 / Min1）。
// 資料從第 2 列開始，因為第 1 列被 Code.gs 拿去放「最後同步：…」橫幅了。
const PIVOT_DESIGNER_BY_COLUMN = {
  A: 'Kathy1', D: 'Kathy1',
  J: 'Lin1',  M: 'Lin1',
  R: 'Min1',  U: 'Min1',
  // Q2 分頁的 Min 區塊比 Q3 往右一欄（S / V），兩種版面都涵蓋進來。
  S: 'Min1',  V: 'Min1'
};

/** 錨點 → 這張表該長什麼樣（欄位一律用標題名稱，不用欄號）。 */
function pivotSpecFor(a1) {
  const match = String(a1).match(/^([A-Z]+)(\d+)$/);
  if (!match) return null;
  const column = match[1];
  const row = Number(match[2]);
  const designer = PIVOT_DESIGNER_BY_COLUMN[column];
  if (!designer) return null;

  if (row === 2) {
    return { designer, rows: ['狀態'], cols: ['類型'], value: { name: '總分', fn: 'SUM' }, display: '分數分布' };
  }
  if (row === 17) {
    return { designer, rows: ['狀態'], cols: ['類型'], value: { name: '狀態', fn: 'COUNTA' }, display: '數量分布' };
  }
  if (row === 31) {
    // 每位設計師第 31 列有兩張小表：左邊照類型數，右邊照狀態數。
    const byType = (column === 'A' || column === 'J' || column === 'R' || column === 'S');
    return byType
      ? { designer, rows: ['類型'], cols: [], value: { name: '類型', fn: 'COUNTA' } }
      : { designer, rows: ['狀態'], cols: [], value: { name: '狀態', fn: 'COUNTA' } };
  }
  return null;
}

/** 只讀不寫：印出每張樞紐分析表目前的來源、列/欄/值/篩選綁到哪一欄。 */
function auditPivots() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  Object.keys(PIVOT_TABS).forEach(tabName => {
    const sheet = ss.getSheetByName(tabName);
    if (!sheet) { Logger.log('❌ 找不到分頁：' + tabName); return; }
    const pivots = sheet.getPivotTables();
    Logger.log('===== ' + tabName + '：' + pivots.length + ' 張樞紐分析表');
    pivots.forEach(pt => {
      const src = pt.getSourceDataRange();
      const srcSheet = src.getSheet();
      const header = srcSheet.getRange(src.getRow(), 1, 1, srcSheet.getLastColumn()).getDisplayValues()[0];
      const nameOf = col => (header[col - 1] || ('col' + col));
      Logger.log([
        pt.getAnchorCell().getA1Notation(),
        'src=' + srcSheet.getName() + '!' + src.getA1Notation(),
        'rows=[' + pt.getRowGroups().map(g => nameOf(g.getSourceDataColumn())).join(', ') + ']',
        'cols=[' + pt.getColumnGroups().map(g => nameOf(g.getSourceDataColumn())).join(', ') + ']',
        'values=[' + pt.getPivotValues().map(v => nameOf(v.getSourceDataColumn()) + '/' + v.getSummarizedBy()).join(', ') + ']',
        'filters=[' + pt.getFilters().map(f => nameOf(f.getSourceDataColumn())).join(', ') + ']'
      ].join(' | '));
    });
  });
}

/** 把兩個 status 分頁的樞紐分析表全部依標題名稱重新綁定。可重複執行。 */
function fixPivots() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const notes = [];

  Object.keys(PIVOT_TABS).forEach(tabName => {
    const quarter = PIVOT_TABS[tabName];
    const label = quarter.slice(4);           // '2026Q3' → 'Q3'
    const sheet = ss.getSheetByName(tabName);
    if (!sheet) { notes.push('❌ 找不到分頁：' + tabName); return; }

    sheet.getPivotTables().forEach(pivot => {
      const anchorCell = pivot.getAnchorCell();
      const anchor = anchorCell.getA1Notation();
      const spec = pivotSpecFor(anchor);
      if (!spec) { notes.push('⚠️ ' + tabName + ' ' + anchor + '：沒有對應設定，略過'); return; }

      try {
        const srcSheet = ss.getSheetByName(spec.designer);
        if (!srcSheet) throw new Error('找不到來源分頁 ' + spec.designer);
        // 第 1 列是同步時間橫幅，標題列在第 2 列，所以來源範圍從第 2 列起算。
        const wanted = srcSheet.getRange(
          SYNC_DATA_START_ROW, 1,
          srcSheet.getMaxRows() - SYNC_DATA_START_ROW + 1,
          srcSheet.getLastColumn()
        );

        let table = pivot;
        const current = pivot.getSourceDataRange();
        if (current.getSheet().getName() !== spec.designer ||
            current.getA1Notation() !== wanted.getA1Notation()) {
          pivot.remove();
          table = anchorCell.createPivotTable(wanted);
          notes.push('🔁 ' + tabName + ' ' + anchor + '：來源改成 ' + spec.designer + '!' + wanted.getA1Notation());
        }

        applyPivotSpec(table, srcSheet, spec, quarter, label);
        notes.push('✅ ' + tabName + ' ' + anchor + ' → ' + spec.designer + '，' + quarter);
      } catch (e) {
        notes.push('❌ ' + tabName + ' ' + anchor + '：' + e.message);
      }
    });
  });

  Logger.log(notes.join('\n'));
  return notes;
}

/** 清掉既有的列/欄/值/篩選，再照 spec 用標題名稱重建。 */
function applyPivotSpec(table, srcSheet, spec, quarter, label) {
  const header = srcSheet.getRange(SYNC_DATA_START_ROW, 1, 1, srcSheet.getLastColumn()).getDisplayValues()[0];
  const columnOf = name => {
    const index = header.indexOf(name);
    if (index === -1) throw new Error(srcSheet.getName() + ' 找不到標題「' + name + '」');
    return index + 1;
  };

  table.getRowGroups().forEach(group => group.remove());
  table.getColumnGroups().forEach(group => group.remove());
  table.getPivotValues().forEach(value => value.remove());
  table.getFilters().forEach(filter => filter.remove());

  spec.rows.forEach(name => table.addRowGroup(columnOf(name)));
  spec.cols.forEach(name => table.addColumnGroup(columnOf(name)));

  const value = table.addPivotValue(
    columnOf(spec.value.name),
    SpreadsheetApp.PivotTableSummarizeFunction[spec.value.fn]
  );
  if (spec.display) value.setDisplayName(label + ' ' + spec.display);

  table.addFilter(
    columnOf('季度'),
    SpreadsheetApp.newFilterCriteria().setVisibleValues([quarter]).build()
  );
}
