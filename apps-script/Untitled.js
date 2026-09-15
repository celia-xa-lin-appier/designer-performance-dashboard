const DESIGNERS = ['Kathy', 'Lin', 'Min'];

// Dashboard 讀 Designer Calculator 裡同步好的 Kathy1 / Lin1 / Min1 三個分頁，
// 而不是三位設計師自己的表。好處是畫面上的數字跟 Q3 / Q2 Designer status 的
// 樞紐分析表一定是同一份資料，權限也只需要這一個檔案。
//
// 代價是多了一層依賴：Code (DO NOT EDIT).gs 的同步一旦停掉（2026-09 就整整
// 停過一週沒人發現），dashboard 會跟著停在舊資料上而不會報錯。判斷資料新不新
// 的依據是這三個分頁第 1 列的「最後同步：…」橫幅。
const DESIGNER_CALCULATOR_ID = '1E144XQoWjzOpUMnDgZidejV4LOiUGZBT5hoPlp3nzko';
const DESIGNER_SOURCES = {
  Kathy: { spreadsheetId: DESIGNER_CALCULATOR_ID, sheetName: 'Kathy1' },
  Lin: { spreadsheetId: DESIGNER_CALCULATOR_ID, sheetName: 'Lin1' },
  Min: { spreadsheetId: DESIGNER_CALCULATOR_ID, sheetName: 'Min1' }
};

// 那三個分頁第 1 列被同步程式拿去放「最後同步」橫幅，標題列在第 2 列，
// 資料從第 3 列開始。刻意寫死數字而不是引用 Code.gs 的 SYNC_DATA_START_ROW，
// 因為跨檔案的常數在載入順序上不保證先被定義。
const DASHBOARD_HEADER_ROW = 2;
const DASHBOARD_FIRST_DATA_ROW = 3;
// Dashboard data is refreshed at most once every 120 seconds.
const CACHE_TTL_SECONDS = 120;
const CACHE_KEY_PREFIX = 'designer_dashboard_v2';
const CACHE_CHUNK_SIZE = 80000;

// Per Irene: B2 "Added to AA Pipeline" stops earning points from this
// quarter onward (inclusive). See the B2 override in getDashboardData().
const B2_SCORE_REMOVED_FROM = { year: 2026, quarter: 3 };

function isQuarterOnOrAfter(quarterText, boundary) {
  const match = String(quarterText || '').trim().match(/^(\d{4})Q([1-4])$/);
  if (!match) return false;
  const year = Number(match[1]);
  const q = Number(match[2]);
  return year > boundary.year || (year === boundary.year && q >= boundary.quarter);
}

/**
 * Serves the dashboard page itself. Deploy as a Web App with access set to
 * "Anyone within [your domain]" — company policy blocks true anonymous
 * access, so the page must be opened by a signed-in Google account in the
 * domain.
 *
 * ?mode=data serves the same JSON getDashboardPayload() returns, but via a
 * plain HTTP response instead of the google.script.run bridge. The page
 * tries fetch(selfUrl + '?mode=data') first (faster when it works — no
 * RPC-bridge overhead) and falls back to google.script.run automatically
 * if that fetch fails for any reason (e.g. the sandboxed iframe this page
 * runs in isn't actually same-origin with this URL).
 *
 * The dashboard's JS is loaded from a data: URI (<?!= scriptDataUri ?> in
 * index.html) instead of living in an inline <script> tag, and instead of
 * a separate <script src> request. Two dead ends led here:
 *  - An inline <script> tag: HtmlService's IFRAME sandbox reliably
 *    corrupts a single character right around the midpoint of a large
 *    inline <script> block's content (some internal two-part relay that
 *    doesn't respect token boundaries) — "Unexpected identifier" at
 *    whatever token happens to straddle that midpoint, at any script size
 *    once the content is non-trivial.
 *  - <script src="...?mode=script"> served via ContentService: the
 *    request 302s through script.googleusercontent.com and the browser
 *    ends up seeing Content-Type: application/binary instead of a JS
 *    type, so it refuses to execute the response as a script.
 * A data: URI sidesteps both — no separate request, and no inline
 * <script> text content for the sandbox to relocate.
 */
function doGet(e) {
  if (e && e.parameter && e.parameter.mode === 'data') {
    let payload;
    try {
      payload = getDashboardPayload();
    } catch (error) {
      payload = {
        designers: DESIGNERS,
        tasks: [],
        error: error && error.message ? error.message : String(error)
      };
    }
    return ContentService.createTextOutput(JSON.stringify(payload))
      .setMimeType(ContentService.MimeType.JSON);
  }

  // script.html holds the dashboard JS base64-encoded, not as plain text —
  // HtmlService.createHtmlOutputFromFile() validates file content as HTML
  // and throws "Malformed HTML content" on plain JS full of stray < and >
  // (comparisons, string literals containing tag-like text). Base64 has
  // none of those characters, so it round-trips through that API safely,
  // and conveniently is already in the exact form a data: URI needs.
  const scriptB64 = HtmlService.createHtmlOutputFromFile('script').getContent();

  const template = HtmlService.createTemplateFromFile('index');
  template.scriptDataUri = 'data:application/javascript;base64,' + scriptB64;
  return template.evaluate()
    .setTitle('Performance Framework')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Called from the page via google.script.run instead of fetch(), so it
 * runs as the signed-in viewer inside the Apps Script sandbox rather than
 * as an anonymous cross-origin request.
 */
function getDashboardPayload() {
  const cachedJson = getCachedDashboardJson();
  if (cachedJson) {
    return JSON.parse(cachedJson);
  }

  // Prevent multiple visitors from rebuilding the same cache simultaneously.
  const lock = LockService.getScriptLock();
  const hasLock = lock.tryLock(10000);

  try {
    // Another request may have completed the cache while this request waited.
    const secondCheck = getCachedDashboardJson();
    if (secondCheck) {
      return JSON.parse(secondCheck);
    }

    const data = getDashboardData();
    setCachedDashboardJson(JSON.stringify(data));
    return data;
  } finally {
    if (hasLock) lock.releaseLock();
  }
}

/**
 * Reads only rows 2:lastRow and columns A:O.
 * Uses one getValues() call per designer sheet.
 */
function getDashboardData() {
  const tasks = [];

  DESIGNERS.forEach(designer => {
    const source = DESIGNER_SOURCES[designer];
    if (!source) return;
    const sheet = SpreadsheetApp.openById(source.spreadsheetId).getSheetByName(source.sheetName);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow < DASHBOARD_FIRST_DATA_ROW || lastCol < 1) return;

    // Columns are located by header text, not by fixed position. On
    // 2026-09-07 the source sheets had an empty "第 1 欄" deleted and
    // "第 2 欄" renamed to "總分", shifting every column from K rightward one
    // to the left. This code was reading O for the total score, so it started
    // reading the 國家 text ("RU", "US") instead — parseNumber() turned that
    // into 0 and the whole dashboard silently showed a score of 0.00 while
    // still counting the right number of tasks. Matching on the header means
    // a column being inserted, deleted or reordered can no longer do that.
    const header = sheet.getRange(DASHBOARD_HEADER_ROW, 1, 1, lastCol).getValues()[0];
    const col = {
      task:        findColumn(header, ['任務'], 0),
      type:        findColumn(header, ['類型'], 1),
      am:          findColumn(header, ['AM'], 2),
      status:      findColumn(header, ['狀態'], 3),
      startDate:   findColumn(header, ['開始日期'], 4),
      endDate:     findColumn(header, ['結束日期'], 5),
      quarter:     findColumn(header, ['季度'], 8),
      typeScore:   findColumn(header, ['類型分數'], 11),
      statusScore: findColumn(header, ['狀態分數'], 12),
      totalScore:  findColumn(header, ['總分', '第2欄'], 13),
      // The free-text video spec column is labeled "size" (not "數量") in
      // the live sheets, and there's a separate "count" column next to it
      // that the sheet itself already resolves to a number per row. Prefer
      // that computed number; only fall back to parsing the size text when
      // count is blank.
      videoSpec:  findColumn(header, ['size', '數量'], -1),
      videoCount: findColumn(header, ['count'], -1)
    };

    const values = sheet.getRange(
      DASHBOARD_FIRST_DATA_ROW, 1,
      lastRow - DASHBOARD_FIRST_DATA_ROW + 1, lastCol
    ).getValues();

    values.forEach(row => {
      const task = String(row[col.task] || '').trim();
      const type = String(row[col.type] || '').trim();
      const am = String(row[col.am] || '').trim();
      const status = String(row[col.status] || '').trim();
      const startDateValue = normalizeDate(row[col.startDate]);
      const rawEndDate = row[col.endDate];
      const endDateValue = normalizeDate(rawEndDate);
      const endDate = formatEndDate(rawEndDate, endDateValue);
      const quarter = String(row[col.quarter] || '').trim();
      const videoSpec = String(row[col.videoSpec] || '').trim();
      const rawVideoCount = col.videoCount === -1 ? '' : row[col.videoCount];
      const videoCount = (rawVideoCount !== '' && rawVideoCount !== null && rawVideoCount !== undefined)
        ? parseNumber(rawVideoCount)
        : parseVideoCount(row[col.videoSpec]);

      const mScore = parseNumber(row[col.typeScore]);
      let nScore = parseNumber(row[col.statusScore]);
      let oScore = parseNumber(row[col.totalScore]);
      const statusCode = extractStatusCode(status);

      // Per Irene: B2 "Added to AA Pipeline" no longer earns points from
      // 2026Q3 onward. The source sheets' own score formulas may still show
      // the old value, so this is enforced here rather than depending on
      // every sheet being updated correctly.
      if (statusCode === 'B2' && isQuarterOnOrAfter(quarter, B2_SCORE_REMOVED_FROM)) {
        nScore = 0;
        oScore = 0;
      }

      if (!task && !status && !endDateValue) return;

      tasks.push({
        designer,
        task,
        am,
        type,
        normalizedType: normalizeType(type),
        status,
        statusCode,
        startDateValue,
        endDate,
        endDateValue,
        quarter,
        mScore,
        nScore,
        oScore,
        videoSpec,
        videoCount
      });
    });
  });

  return {
    designers: DESIGNERS,
    tasks,
    updatedAt: Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd'T'HH:mm:ssXXX"
    )
  };
}

/**
 * CacheService limits the size of each key, so the compressed JSON is split
 * into chunks. This keeps caching reliable even when the dashboard grows.
 */
function getCachedDashboardJson() {
  const cache = CacheService.getScriptCache();
  const metaText = cache.get(`${CACHE_KEY_PREFIX}:meta`);
  if (!metaText) return null;

  try {
    const meta = JSON.parse(metaText);
    const keys = [];

    for (let i = 0; i < meta.chunkCount; i++) {
      keys.push(`${CACHE_KEY_PREFIX}:chunk:${i}`);
    }

    const chunks = cache.getAll(keys);
    const encoded = keys.map(key => chunks[key] || '').join('');
    if (!encoded) return null;

    const compressed = Utilities.base64Decode(encoded);
    const blob = Utilities.ungzip(Utilities.newBlob(compressed));
    return blob.getDataAsString('UTF-8');
  } catch (error) {
    clearDashboardCache();
    return null;
  }
}

function setCachedDashboardJson(json) {
  const cache = CacheService.getScriptCache();
  const compressed = Utilities.gzip(
    Utilities.newBlob(json, 'application/json', 'dashboard.json')
  );
  const encoded = Utilities.base64Encode(compressed.getBytes());
  const chunkCount = Math.ceil(encoded.length / CACHE_CHUNK_SIZE);
  const entries = {};

  for (let i = 0; i < chunkCount; i++) {
    entries[`${CACHE_KEY_PREFIX}:chunk:${i}`] = encoded.slice(
      i * CACHE_CHUNK_SIZE,
      (i + 1) * CACHE_CHUNK_SIZE
    );
  }

  cache.putAll(entries, CACHE_TTL_SECONDS);
  cache.put(
    `${CACHE_KEY_PREFIX}:meta`,
    JSON.stringify({ chunkCount }),
    CACHE_TTL_SECONDS
  );
}

/**
 * Keeps the cache warm in the background so a visitor's google.script.run
 * call almost always hits a cache entry instead of waiting on a live
 * Sheets read. Installed as a time-driven trigger by setupCacheTrigger();
 * not called by the page itself.
 */
function refreshDashboardCache() {
  const lock = LockService.getScriptLock();
  const hasLock = lock.tryLock(10000);
  if (!hasLock) return;

  try {
    const data = getDashboardData();
    setCachedDashboardJson(JSON.stringify(data));
  } finally {
    lock.releaseLock();
  }
}

/**
 * Run once manually from the Apps Script editor to install the background
 * refresh trigger. Safe to re-run — only clears a prior
 * refreshDashboardCache trigger first, so it never stacks duplicates and
 * never touches importWithFormat's trigger in Code (DO NOT EDIT).js.
 */
function setupCacheTrigger() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'refreshDashboardCache')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('refreshDashboardCache')
    .timeBased()
    .everyMinutes(1)
    .create();
}

function clearDashboardCache() {
  const cache = CacheService.getScriptCache();
  const metaText = cache.get(`${CACHE_KEY_PREFIX}:meta`);
  const keys = [`${CACHE_KEY_PREFIX}:meta`];

  if (metaText) {
    try {
      const meta = JSON.parse(metaText);
      for (let i = 0; i < meta.chunkCount; i++) {
        keys.push(`${CACHE_KEY_PREFIX}:chunk:${i}`);
      }
    } catch (error) {
      // Ignore malformed cache metadata.
    }
  }

  cache.removeAll(keys);
}

function normalizeType(value) {
  const text = String(value || '').trim();
  if (!text) return '';

  // 先把寫法差異抹平再查表 —— 大小寫、斜線兩側的空白、連字號。這樣
  // 「Interactive/Multi Video」「Interactive / Multi-video」會落在同一個 key，
  // 不必每出現一種寫法就補一條。（2026-09 有 9 筆 Interactive/Multi Video
  // 因為只差結尾那個 " Video" 就被判成無法辨識的類型。）
  const key = text
    .toLowerCase()
    .replace(/[-\u2013\u2014]/g, ' ')
    .replace(/\s*\/\s*/g, '/')
    .replace(/\s+/g, ' ')
    .trim();

  const mapping = {
    'branding': 'Branding',
    'video': 'Video',
    'ai stories': 'AI stories',
    'ai story': 'AI stories',
    'playable': 'Playable',
    'interactive/multi': 'Interactive/Multi',
    'interactive/multi video': 'Interactive/Multi',
    'static/banner': 'Static/Banner',
    'other': 'other'
  };

  return mapping[key] || '';
}

function extractStatusCode(status) {
  if (!status) return 'No Status';

  const text = String(status).trim();
  const match = text.match(/^(A[1-8]|B[1-2]|C[1-2]|F)/i);

  if (match) return match[1].toUpperCase();
  if (text.toLowerCase() === 'fail') return 'F';

  return text;
}

/**
 * Index of the first column whose header exactly matches one of `names`, or
 * `fallback` when none of them is present. Whitespace is stripped before
 * comparing, so "第 2 欄" and "第2欄" count as the same header. The match is
 * exact rather than a substring test, so looking for 類型 never lands on
 * 類型分數.
 */
function findColumn(header, names, fallback) {
  const wanted = names.map(name => name.replace(/\s+/g, ''));

  for (let i = 0; i < header.length; i++) {
    const text = String(header[i] || '').replace(/\s+/g, '');
    if (text && wanted.indexOf(text) !== -1) return i;
  }

  return fallback;
}

/**
 * How many finished videos a 任務 row represents, read from the 數量 column.
 *
 * The column is free text rather than a number, and two notations are in use:
 *
 *   "1. 16:9 15s 2. 16:9 6s 3. 9:16 15s 4. 9:16 6s"  -> 4  (numbered list)
 *   "1. 9:16 14s 2. 16:9 15s x2"                     -> 3  (item 2 delivered twice)
 *   "9:16, 16:9 6s & 15s"                            -> 4  (every ratio x every length)
 *   "4"                                              -> 4  (already counted by hand)
 *   ""                                               -> 0  (nothing claimed)
 *
 * The numbered-list form wins when it is present because it is explicit. The
 * cross-product form is only a fallback for rows written the way the request
 * was originally phrased, so those rows still count instead of silently
 * reading as a single video.
 */
function parseVideoCount(value) {
  if (value === null || value === undefined || value === '') return 0;
  if (typeof value === 'number') return isNaN(value) ? 0 : value;

  const text = String(value).trim();
  if (!text) return 0;

  // A cell someone already totalled by hand.
  if (/^\d+(\.\d+)?$/.test(text)) return Number(text);

  // Numbered list. The punctuation has to be followed by whitespace or the end
  // of the cell, so a decimal like "1.5" is not mistaken for item 1.
  const markers = /(?:^|[\s;,、])(\d+)\s*[.．、)](?=\s|$)/g;
  const starts = [];
  let match;
  while ((match = markers.exec(text)) !== null) {
    starts.push(markers.lastIndex);
  }

  if (starts.length) {
    let total = 0;
    for (let i = 0; i < starts.length; i++) {
      const item = text.slice(starts[i], i + 1 < starts.length ? starts[i + 1] : text.length);
      total += itemMultiplier(item);
    }
    return total;
  }

  // Cross product: "9:16, 16:9" x "6s & 15s". Ratios and lengths are counted as
  // sets so "6s, 6s" does not inflate the total.
  const ratios = uniqueMatches(text, /\b\d{1,2}\s*:\s*\d{1,2}\b/g);
  const lengths = uniqueMatches(text, /\b\d{1,3}\s*s\b/gi);

  if (ratios.length || lengths.length) {
    return Math.max(ratios.length, 1) * Math.max(lengths.length, 1);
  }

  // Text that describes something, but nothing this code recognises. Treating
  // it as one video is closer than treating it as none.
  return itemMultiplier(text);
}

/**
 * The "x2" on the end of a list item, or 1 when there is none.
 *
 * 尺寸跟份數在這一欄是同一種寫法 ——「320x480 x1」「1280 x 720 x1」——
 * 所以先把尺寸剔掉再找份數，否則「320x480」會被讀成 ×480，一列橫幅素材
 * 就能把一個人的影片總數灌高好幾百支。
 */
function itemMultiplier(text) {
  const sizes = /\d{2,4}\s*[x×X]\s*\d{2,4}/g;
  const match = /[x×X]\s*(\d+)/.exec(String(text).replace(sizes, ' '));
  if (!match) return 1;
  const count = Number(match[1]);
  return count > 0 ? count : 1;
}

/** Distinct matches of `pattern`, compared with whitespace stripped. */
function uniqueMatches(text, pattern) {
  const seen = {};
  const out = [];
  let match;
  while ((match = pattern.exec(text)) !== null) {
    const key = match[0].replace(/\s+/g, '').toLowerCase();
    if (!seen[key]) {
      seen[key] = true;
      out.push(key);
    }
  }
  return out;
}

function parseNumber(value) {
  if (value === null || value === undefined || value === '') return 0;

  if (typeof value === 'number') {
    return Number.isFinite(value) ? value : 0;
  }

  const cleaned = String(value)
    .replace(/,/g, '')
    .replace(/%/g, '')
    .trim();

  const num = Number(cleaned);
  return Number.isFinite(num) ? num : 0;
}

function formatEndDate(rawValue, normalizedValue) {
  if (!rawValue) return '';

  if (rawValue instanceof Date && !isNaN(rawValue.getTime())) {
    return normalizedValue;
  }

  return String(rawValue).trim();
}

function normalizeDate(value) {
  if (!value) return '';

  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone(),
      'yyyy-MM-dd'
    );
  }

  const text = String(value).trim();
  const zhMatch = text.match(/(\d{2,4})年\s*(\d{1,2})月\s*(\d{1,2})日/);

  if (zhMatch) {
    let year = Number(zhMatch[1]);
    if (year < 100) year += 2000;

    const month = Number(zhMatch[2]);
    const day = Number(zhMatch[3]);

    return `${year}-${String(month).padStart(2, '0')}-${String(day).padStart(2, '0')}`;
  }

  const date = new Date(text);
  if (isNaN(date.getTime())) return '';

  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    'yyyy-MM-dd'
  );
}
