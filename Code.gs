// ============================================================
//  Stock Learning Portal — Google Apps Script Backend
//  Code.gs
// ============================================================

var SHEET_ID = ''; // Leave blank — script uses the bound spreadsheet

// ── Entry point ──────────────────────────────────────────────
function doGet(e) {
  var page = e && e.parameter && e.parameter.page;

  // Get the real deployed exec URL (works in both dev and production)
  var execUrl = ScriptApp.getService().getUrl();

  if (page === 'calculator') {
    return HtmlService.createHtmlOutputFromFile('Calculator')
      .setTitle('Trade Calculator — Stock Learning Portal')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  if (page === 'reflection') {
    return HtmlService.createHtmlOutputFromFile('Reflection')
      .setTitle('Trade Reflection — Stock Learning Portal')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  if (page === 'candles') {
    return HtmlService.createHtmlOutputFromFile('Candles')
      .setTitle('Candlestick — Stock Learning Portal')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  if (page === 'sectors') {
    return HtmlService.createHtmlOutputFromFile('Sectors')
      .setTitle('TSE Sector Map — Stock Learning Portal')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  if (page === 'swingrules') {
    return HtmlService.createHtmlOutputFromFile('SwingRules')
      .setTitle('Swing Trading Rules — Stock Learning Portal')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  if (page === 'chartpatterns') {
    return HtmlService.createHtmlOutputFromFile('ChartPatterns')
      .setTitle('Chart Patterns — Stock Learning Portal')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  if (page === 'indicators') {
    return HtmlService.createHtmlOutputFromFile('Indicators')
      .setTitle('Technical Indicators — Stock Learning Portal')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  if (page === 'principles') {
    return HtmlService.createHtmlOutputFromFile('Principles')
      .setTitle('Investing Principles — Stock Learning Portal')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  if (page === 'watchlist') {
    return HtmlService.createHtmlOutputFromFile('Watchlist')
      .setTitle('Watchlist — Stock Learning Portal')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }

  // Main portal — inject the real URL so tiles can open other pages correctly
  var template = HtmlService.createTemplateFromFile('Index');
  template.baseUrl = execUrl;
  return template.evaluate()
    .setTitle('Stock Learning Portal')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ── Get spreadsheet ──────────────────────────────────────────
function getSpreadsheet() {
  if (SHEET_ID) return SpreadsheetApp.openById(SHEET_ID);
  return SpreadsheetApp.getActiveSpreadsheet();
}

// ── Fetch all data (called by frontend) ─────────────────────
function getAllData() {
  try {
    var ss = getSpreadsheet();
    var terms = getTerms(ss);
    return {
      categories: getCategories(ss),
      terms:      terms,
      meta: { lastUpdated: new Date().toISOString(), totalTerms: terms.length }
    };
  } catch (e) {
    return { error: e.message };
  }
}

// ── Read Categories sheet ────────────────────────────────────
function getCategories(ss) {
  var sheet = ss.getSheetByName('Categories');
  if (!sheet) return [];
  var rows = sheet.getDataRange().getValues();
  var headers = rows[0];
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    if (!row[0]) continue;
    var obj = {};
    headers.forEach(function(h, idx) { obj[h] = row[idx]; });
    result.push(obj);
  }
  result.sort(function(a, b) { return (a.sort_order || 0) - (b.sort_order || 0); });
  return result;
}

// ── Read Terms sheet ─────────────────────────────────────────
function getTerms(ss) {
  var sheet = ss.getSheetByName('Terms');
  if (!sheet) return [];
  var rows = sheet.getDataRange().getValues();
  var headers = rows[0];
  var result = [];
  for (var i = 1; i < rows.length; i++) {
    var row = rows[i];
    if (!row[0]) continue;
    var obj = {};
    headers.forEach(function(h, idx) { obj[h] = row[idx]; });
    obj.tags = obj.tags ? String(obj.tags).split(',').map(function(t) { return t.trim(); }) : [];
    result.push(obj);
  }
  return result;
}

// ── Save progress ────────────────────────────────────────────
function saveProgress(userId, termId, status) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Progress');
    if (!sheet) {
      sheet = ss.insertSheet('Progress');
      sheet.appendRow(['userId','termId','status','timestamp']);
    }
    var data = sheet.getDataRange().getValues();
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === userId && data[i][1] === termId) {
        sheet.getRange(i + 1, 3).setValue(status);
        sheet.getRange(i + 1, 4).setValue(new Date().toISOString());
        return { success: true };
      }
    }
    sheet.appendRow([userId, termId, status, new Date().toISOString()]);
    return { success: true };
  } catch (e) {
    return { error: e.message };
  }
}

// ── Get progress for a user ──────────────────────────────────
function getProgress(userId) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Progress');
    if (!sheet) return {};
    var data = sheet.getDataRange().getValues();
    var result = {};
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] === userId) result[data[i][1]] = data[i][2];
    }
    return result;
  } catch (e) {
    return {};
  }
}


// ============================================================
//  TRADE REFLECTION FUNCTIONS
// ============================================================

// Save a new reflection
function saveReflection(data) {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Reflections');
    
    if (!sheet) {
      sheet = ss.insertSheet('Reflections');
      var headers = [
        'timestamp','code','stock_name','entry_date','exit_date',
        'entry_price','exit_price','shares','result','pl_yen','pl_pct',
        'reason','pattern','planned_entry','stop_loss','target_exit',
        'went_right','went_wrong','lessons'
      ];
      sheet.appendRow(headers);
      sheet.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#E6F1FB');
      sheet.setFrozenRows(1);
    }
    
    var row = [
      data.timestamp, data.code, data.stock_name, data.entry_date, data.exit_date,
      data.entry_price, data.exit_price, data.shares, data.result, data.pl_yen, data.pl_pct,
      data.reason, data.pattern, data.planned_entry, data.stop_loss, data.target_exit,
      data.went_right, data.went_wrong, data.lessons
    ];
    
    sheet.appendRow(row);
    return { success: true };
  } catch (e) {
    return { success: false, error: e.message };
  }
}

// Get all reflections - returns JSON string to avoid GAS serialization issues
function getReflections() {
  try {
    var ss = getSpreadsheet();
    var sheet = ss.getSheetByName('Reflections');

    if (!sheet) return JSON.stringify({ reflections: [] });

    var data = sheet.getDataRange().getValues();
    if (data.length < 2) return JSON.stringify({ reflections: [] });

    var headers = data[0].map(function(h) { return String(h); });
    var reflections = [];

    for (var i = data.length - 1; i >= 1; i--) {
      var row = data[i];
      var obj = {};
      headers.forEach(function(h, idx) {
        var val = row[idx];
        if (val instanceof Date) {
          obj[h] = Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        } else {
          obj[h] = val === null || val === undefined ? '' : val;
        }
      });
      reflections.push(obj);
    }

    return JSON.stringify({ reflections: reflections });
  } catch (e) {
    return JSON.stringify({ reflections: [], error: e.message });
  }
}


// ============================================================
//  WATCHLIST FUNCTIONS
// ============================================================

var TTL_PRICE = 300; // current price cache — 5 minutes

function formatDate_(value) {
  if (value instanceof Date) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(value);
}

function getWatchlistSheet() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('Watchlist');
  if (!sheet) {
    sheet = ss.insertSheet('Watchlist');
    sheet.appendRow(['symbol','name','noticePrice','noticeDate','notes']);
    sheet.getRange(1,1,1,5).setFontWeight('bold').setBackground('#E6F1FB');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getAllWatchlist() {
  var sheet = getWatchlistSheet();
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  return sheet.getRange(2, 1, lastRow - 1, 5).getValues().map(function(row) {
    return {
      symbol:      row[0] || '',
      name:        row[1] || '',
      noticePrice: Number(row[2]) || 0,
      noticeDate:  row[3] ? formatDate_(row[3]) : '',
      notes:       row[4] || ''
    };
  });
}

function addWatchlistItem(item) {
  var existing = getAllWatchlist();
  for (var i = 0; i < existing.length; i++) {
    if (existing[i].symbol === item.symbol) throw new Error(item.symbol + ' is already on your watchlist');
  }
  getWatchlistSheet().appendRow([
    item.symbol || '',
    item.name || '',
    item.noticePrice || 0,
    item.noticeDate || formatDate_(new Date()),
    item.notes || ''
  ]);
  return getAllWatchlist();
}

function deleteWatchlistItem(index) {
  var sheet = getWatchlistSheet();
  var row = index + 2;
  if (row < 2 || row > sheet.getLastRow()) throw new Error('Invalid index: ' + index);
  sheet.deleteRow(row);
  return getAllWatchlist();
}

// ── Weekly price history ──────────────────────────────────────

function getWatchlistHistorySheet() {
  var ss = getSpreadsheet();
  var sheet = ss.getSheetByName('WatchlistHistory');
  if (!sheet) {
    sheet = ss.insertSheet('WatchlistHistory');
    sheet.appendRow(['date','symbol','name','price','noticePrice','changePct']);
    sheet.getRange(1,1,1,6).setFontWeight('bold').setBackground('#E6F1FB');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

// Attach to a weekly time trigger (Apps Script editor -> Triggers -> + Add
// Trigger -> Function: recordWatchlistPrices -> Time-driven -> Week timer).
function recordWatchlistPrices() {
  var watchlist = getAllWatchlist();
  if (watchlist.length === 0) return;

  var symbols = watchlist.map(function(w) { return w.symbol; });
  var priceMap = fetchAllPrices(symbols);
  var today = formatDate_(new Date());
  var sheet = getWatchlistHistorySheet();

  watchlist.forEach(function(w) {
    var price = priceMap[w.symbol];
    if (!price) return;
    var changePct = w.noticePrice > 0 ? Math.round(((price - w.noticePrice) / w.noticePrice) * 10000) / 100 : 0;
    sheet.appendRow([today, w.symbol, w.name, price, w.noticePrice, changePct]);
  });
}

function getWatchlistHistory() {
  var sheet = getWatchlistHistorySheet();
  var lastRow = sheet.getLastRow();
  if (lastRow <= 1) return [];
  return sheet.getRange(2, 1, lastRow - 1, 6).getValues().map(function(row) {
    return {
      date:        row[0] ? formatDate_(row[0]) : '',
      symbol:      row[1] || '',
      name:        row[2] || '',
      price:       Number(row[3]) || 0,
      noticePrice: Number(row[4]) || 0,
      changePct:   Number(row[5]) || 0
    };
  });
}

// ── Yahoo Finance helpers ─────────────────────────────────────

// Fetches current price for each symbol in one batched call (via
// UrlFetchApp.fetchAll), with a short cache to avoid refetching on
// every page load.
function fetchAllPrices(symbols) {
  if (!symbols || symbols.length === 0) return {};

  var cache = CacheService.getScriptCache();
  var results = {};
  var toFetch = [];

  symbols.forEach(function(sym) {
    var cached = cache.get('yf_price_' + sym);
    if (cached) {
      results[sym] = Number(cached);
    } else {
      toFetch.push(sym);
    }
  });

  if (toFetch.length > 0) {
    var requests = toFetch.map(function(sym) {
      return {
        url: 'https://query1.finance.yahoo.com/v8/finance/chart/' + sym + '?interval=1d&range=1d',
        headers: { 'User-Agent': 'Mozilla/5.0', 'Accept': 'application/json' },
        muteHttpExceptions: true
      };
    });

    var responses = UrlFetchApp.fetchAll(requests);
    responses.forEach(function(res, i) {
      var sym = toFetch[i];
      try {
        if (res.getResponseCode() === 200) {
          var data = JSON.parse(res.getContentText());
          var price = data.chart && data.chart.result && data.chart.result[0] &&
                      data.chart.result[0].meta && data.chart.result[0].meta.regularMarketPrice;
          if (price) {
            results[sym] = price;
            try { cache.put('yf_price_' + sym, String(price), TTL_PRICE); } catch (e) {}
          }
        }
      } catch (e) {}
    });
  }

  return results;
}

// Looks up a symbol's display name (used when adding a new watchlist item
// so the user only has to type the stock code, not the full name).
function fetchYahooQuote(symbol) {
  var url = 'https://query1.finance.yahoo.com/v8/finance/chart/' + symbol + '?interval=1d&range=1d';
  try {
    var res = UrlFetchApp.fetch(url, {
      muteHttpExceptions: true,
      headers: { 'User-Agent': 'Mozilla/5.0', 'Accept': 'application/json' }
    });
    if (res.getResponseCode() !== 200) return { name: symbol };
    var data = JSON.parse(res.getContentText());
    var meta = data.chart && data.chart.result && data.chart.result[0] && data.chart.result[0].meta;
    var name = meta ? (meta.longName || meta.shortName || symbol) : symbol;
    return { name: name };
  } catch (e) {
    return { name: symbol };
  }
}