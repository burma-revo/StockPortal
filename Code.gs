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