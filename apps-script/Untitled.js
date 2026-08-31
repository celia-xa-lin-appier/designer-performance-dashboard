const SPREADSHEET_ID = '1E144XQoWjzOpUMnDgZidejV4LOiUGZBT5hoPlp3nzko';
const DESIGNERS = ['Kathy', 'Lin', 'Min'];

// Dashboard data is refreshed at most once every 120 seconds.
const CACHE_TTL_SECONDS = 120;
const CACHE_KEY_PREFIX = 'designer_dashboard_v2';
const CACHE_CHUNK_SIZE = 80000;

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

  return HtmlService.createHtmlOutputFromFile('index')
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
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  const tasks = [];

  DESIGNERS.forEach(designer => {
    const sheet = ss.getSheetByName(designer);
    if (!sheet) return;

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;

    // A:O = 15 columns. Row 1 is the header, so start at row 2.
    const values = sheet.getRange(2, 1, lastRow - 1, 15).getValues();

    values.forEach(row => {
      // Sheet columns:
      // A Task, B Type, D Status, F End Date, I Quarter,
      // M Multiplier, N Status Score, O Total Score.
      const task = String(row[0] || '').trim();
      const type = String(row[1] || '').trim();
      const status = String(row[3] || '').trim();
      const rawEndDate = row[5];
      const endDateValue = normalizeDate(rawEndDate);
      const endDate = formatEndDate(rawEndDate, endDateValue);
      const quarter = String(row[8] || '').trim();

      const mScore = parseNumber(row[12]);
      const nScore = parseNumber(row[13]);
      const oScore = parseNumber(row[14]);

      if (!task && !status && !endDateValue) return;
      if (oScore === 0) return;

      tasks.push({
        designer,
        task,
        type,
        normalizedType: normalizeType(type),
        status,
        statusCode: extractStatusCode(status),
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
