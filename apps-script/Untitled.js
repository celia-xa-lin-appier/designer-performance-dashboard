const DESIGNERS = ['Kathy', 'Lin', 'Min'];

// The dashboard reads each designer's tasks straight from her own
// spreadsheet — not from the synced copy in Designer Calculator. Reading
// has never needed write access, so this is immune to a designer's tab in
// Designer Calculator being blocked by a protected range there (as
// happened with Min's). Code (DO NOT EDIT).js keeps syncing List data into
// Designer Calculator on its own schedule, independent of this — that
// copy now exists for the "Q3 Designer status" pivot table and manual
// review, not for the dashboard itself.
const DESIGNER_SOURCES = {
  Kathy: { spreadsheetId: '1nfSmY4GeRLuy3YmOXhjiJ1cnfNkXa9Q_LDXyj5mUd-E', sheetName: 'List' },
  Lin: { spreadsheetId: '1S6WO4uedwmJ2aGpGPVK906ab6x0YO0GoxSbGmGV9uyU', sheetName: 'List' },
  Min: { spreadsheetId: '1ZqI-v3RNYPX8s648VAfQxDoDvKaRqWJcvQf1YSDCeRI', sheetName: 'List' }
};

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
    if (lastRow < 2) return;

    // A:O = 15 columns. Row 1 is the header, so start at row 2.
    const values = sheet.getRange(2, 1, lastRow - 1, 15).getValues();

    values.forEach(row => {
      // Sheet columns:
      // A Task, B Type, C AM, D Status, E Start Date, F End Date, I Quarter,
      // M Multiplier, N Status Score, O Total Score.
      const task = String(row[0] || '').trim();
      const type = String(row[1] || '').trim();
      const am = String(row[2] || '').trim();
      const status = String(row[3] || '').trim();
      const startDateValue = normalizeDate(row[4]);
      const rawEndDate = row[5];
      const endDateValue = normalizeDate(rawEndDate);
      const endDate = formatEndDate(rawEndDate, endDateValue);
      const quarter = String(row[8] || '').trim();

      const mScore = parseNumber(row[12]);
      let nScore = parseNumber(row[13]);
      let oScore = parseNumber(row[14]);
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
        oScore
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

  const key = text.toLowerCase().replace(/\s+/g, ' ');
  const mapping = {
    'branding': 'Branding',
    'video': 'Video',
    'ai stories': 'AI stories',
    'ai story': 'AI stories',
    'playable': 'Playable',
    'interactive/multi': 'Interactive/Multi',
    'interactive / multi': 'Interactive/Multi',
    'static/banner': 'Static/Banner',
    'static / banner': 'Static/Banner',
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
