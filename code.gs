/**
 * SOSBOARD v2.0+ - Intelligent Organic Content Operating System
 * Backend Logic (Merged & Optimized)
 */

// --- MAIN ENTRY POINTS ---

function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('SOSBOARD v2.0 - Intelligent Dashboard')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  // Case-insensitive fix: convert to exact file names in the project
  const files = {
    'Frontend': 'Frontend',
    'frontend': 'Frontend',
    'Styles': 'index', // Styles are usually in index or separate
  };
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

/**
 * Initialize System
 */
function initSOSBOARD() {
  const result = setupAllTriggers();
  logAudit('INIT', 'System initialized by ' + Session.getActiveUser().getEmail());
  return result;
}

// --- DATABASE OPERATIONS (Sheets) ---

const TARGET_SS_ID = '1NA0kJyEypr25RwqeNjM_B5iLIdH4OEDLh3qU6GwHg_c';
const DB_SHEET_NAME = 'Database_JOY';
const SETTINGS_SHEET_NAME = 'Settings_JOY';
const GROWTH_SHEET_NAME = 'Growth_JOY';
const LIBRARY_SHEET_NAME = 'Content_Library';
const AUDIT_SHEET_NAME = 'Audit_Log';
const VERSION_SHEET_NAME = 'Version_History';

function getTargetId() {
  return PropertiesService.getScriptProperties().getProperty('TARGET_SPREADSHEET_ID') || TARGET_SS_ID;
}

function getRequestSheet(name = DB_SHEET_NAME) {
  let ss;
  try {
    ss = SpreadsheetApp.openById(getTargetId());
  } catch (e) {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    if (name === DB_SHEET_NAME) {
      sheet.appendRow([
        'id', 'startDate', 'label', 'dateRange', 
        'reach', 'engagement', 'views', 'profileActivity',
        'followersPct', 'nonFollowersPct', 'storiesPct', 'reelsPct', 'feedsPct',
        'win_type', 'win_title', 'win_link', 'win_what', 'win_why', 'win_todo',
        'lose_type', 'lose_title', 'lose_link', 'lose_what', 'lose_why', 'lose_todo',
        'topic_tags', 'content_score', 'post_hour', 'post_day', 'version_id'
      ]);
    }
  }
  return sheet;
}

function getWeeklyData() {
  try {
    const sheet = getRequestSheet(DB_SHEET_NAME);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return getDummyWeeklyData();
    
    const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const getIdx = (name) => headers.indexOf(name);
    
    return data.map(row => ({
      id: String(row[getIdx('id')]),
      startDate: row[getIdx('startDate')],
      label: row[getIdx('label')],
      dateRange: row[getIdx('dateRange')],
      metrics: {
        reach: Number(row[getIdx('reach')]) || 0,
        engagement: Number(row[getIdx('engagement')]) || 0,
        views: Number(row[getIdx('views')]) || 0,
        profileActivity: Number(row[getIdx('profileActivity')]) || 0,
        followersPct: Number(row[getIdx('followersPct')]) || 0,
        nonFollowersPct: Number(row[getIdx('nonFollowersPct')]) || 0,
        storiesPct: Number(row[getIdx('storiesPct')]) || 0,
        reelsPct: Number(row[getIdx('reelsPct')]) || 0,
        feedsPct: Number(row[getIdx('feedsPct')]) || 0
      },
      winningContent: {
        type: row[getIdx('win_type')],
        title: row[getIdx('win_title')],
        link: row[getIdx('win_link')],
        analysis: {
          what: row[getIdx('win_what')],
          why: row[getIdx('win_why')],
          todo: row[getIdx('win_todo')]
        }
      },
      losingContent: {
        type: row[getIdx('lose_type')],
        title: row[getIdx('lose_title')],
        link: row[getIdx('lose_link')],
        analysis: {
          what: row[getIdx('lose_what')],
          why: row[getIdx('lose_why')],
          todo: row[getIdx('lose_todo')]
        }
      },
      topicTags: JSON.parse(row[getIdx('topic_tags')] || '[]'),
      contentScore: Number(row[getIdx('content_score')] || 0),
      postHour: Number(row[getIdx('post_hour')] || 19),
      postDay: row[getIdx('post_day')] || 'Wednesday',
      versionId: row[getIdx('version_id')]
    }));
  } catch (e) {
    Logger.log('Error getWeeklyData: ' + e.toString());
    return [];
  }
}

function saveWeeklyData(jsonData) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const sheet = getRequestSheet(DB_SHEET_NAME);
    const headers = [
      'id', 'startDate', 'label', 'dateRange', 
      'reach', 'engagement', 'views', 'profileActivity',
      'followersPct', 'nonFollowersPct', 'storiesPct', 'reelsPct', 'feedsPct',
      'win_type', 'win_title', 'win_link', 'win_what', 'win_why', 'win_todo',
      'lose_type', 'lose_title', 'lose_link', 'lose_what', 'lose_why', 'lose_todo',
      'topic_tags', 'content_score', 'post_hour', 'post_day', 'version_id'
    ];
    
    sheet.clear();
    sheet.appendRow(headers);
    
    const rows = jsonData.map(item => [
      item.id, item.startDate, item.label, item.dateRange,
      item.metrics.reach, item.metrics.engagement, item.metrics.views, item.metrics.profileActivity,
      item.metrics.followersPct, item.metrics.nonFollowersPct, item.metrics.storiesPct, item.metrics.reelsPct, item.metrics.feedsPct,
      item.winningContent.type, item.winningContent.title, item.winningContent.link, item.winningContent.analysis.what, item.winningContent.analysis.why, item.winningContent.analysis.todo,
      item.losingContent.type, item.losingContent.title, item.losingContent.link, item.losingContent.analysis.what, item.losingContent.analysis.why, item.losingContent.analysis.todo,
      JSON.stringify(item.topicTags || []), item.contentScore || 0, item.postHour || 19, item.postDay || 'Wednesday', item.versionId || Date.now().toString()
    ]);
    
    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }
    
    // Sync to library and log audit
    saveToContentLibrary(jsonData);
    logAudit('SAVE_DATA', `Saved ${rows.length} weeks of data`);
    
    return "Success";
  } catch (e) {
    return "Error: " + e.toString();
  } finally {
    lock.releaseLock();
  }
}

// --- AI OPERATIONS (Gemini) ---

function getGeminiKey() {
  return PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY') || '';
}

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;
    const data = payload.data;
    let result = {};
    
    switch(action) {
      case 'getWeeklyData':
        result = getWeeklyData();
        break;
      case 'saveWeeklyData':
        result = saveWeeklyData(data);
        break;
      case 'generateAIInsight':
        result = generateAIInsight(data);
        break;
      case 'getContentLibrary':
        result = getContentLibrary();
        break;
      default:
        result = { error: 'Unknown action' };
    }
    
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ error: error.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function generateAIInsight(weekDataArray) {
  try {
    const apiKey = getGeminiKey();
    const recent = (weekDataArray || []).slice(-8);
    if (recent.length === 0) return { error: 'NO_DATA' };
    
    const summary = recent.map(w => {
      const er = w.metrics.reach > 0 ? ((w.metrics.engagement / w.metrics.reach) * 100).toFixed(2) : 0;
      return `${w.label}: Reach=${w.metrics.reach}, ER=${er}%, Win="${w.winningContent.title}"`;
    }).join('\n');
    
    const prompt = `Analisis data performa JOY (8 minggu terakhir):\n${summary}\nBerikan JSON: { "summary": "", "win_pattern": "", "avoid": "", "recommendations": [], "next_week_focus": "", "score": 0 }`;
    
    const payload = {
      contents: [{ role: 'user', parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.5, responseMimeType: "application/json" }
    };
    
    const resp = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=' + apiKey, {
      method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true
    });
    
    if (resp.getResponseCode() !== 200) return { error: 'API_ERROR' };
    const result = JSON.parse(resp.getContentText());
    const text = result.candidates[0].content.parts[0].text;
    return { success: true, data: JSON.parse(text) };
  } catch (e) { return { error: e.toString() }; }
}

function generateContentScore(draftData) {
  try {
    const apiKey = getGeminiKey();
    const prompt = `Audit draf konten JOY:\nPlatform: ${draftData.platform}\nJudul: ${draftData.title}\nCaption: ${draftData.caption}\nBerikan JSON: { "score": 0, "strengths": [], "weaknesses": [], "recommendations": [] }`;
    
    const payload = {
      contents: [{ role: 'user', parts: [{ text: prompt }] }],
      generationConfig: { temperature: 0.5, responseMimeType: "application/json" }
    };
    
    const resp = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=' + apiKey, {
      method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true
    });
    
    const result = JSON.parse(resp.getContentText());
    return { success: true, data: JSON.parse(result.candidates[0].content.parts[0].text) };
  } catch (e) { return { error: e.toString() }; }
}

function askLLMAgent(question) {
  try {
    const apiKey = getGeminiKey();
    const weekData = getWeeklyData();
    let context = weekData.slice(-5).map(w => `${w.label}: Reach ${w.metrics.reach}`).join(' | ');
    const prompt = `Kamu asisten SOSBOARD JOY. Konteks: ${context}. Tanya: ${question}`;
    
    const payload = { contents: [{ role: 'user', parts: [{ text: prompt }] }] };
    const resp = UrlFetchApp.fetch('https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash-latest:generateContent?key=' + apiKey, {
      method: 'post', contentType: 'application/json', payload: JSON.stringify(payload), muteHttpExceptions: true
    });
    
    const result = JSON.parse(resp.getContentText());
    return { success: true, text: result.candidates[0].content.parts[0].text };
  } catch (e) { return { error: e.toString() }; }
}

// --- CONTENT LIBRARY ---

function getContentLibrary() {
  try {
    const sheet = getRequestSheet(LIBRARY_SHEET_NAME);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return { success: true, data: [] };
    const data = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    return {
      success: true,
      data: data.map(r => ({
        id: r[0], weekId: r[1], platform: r[2], type: r[3], title: r[4], link: r[5],
        tags: JSON.parse(r[6] || '[]'), score: r[7], metrics: JSON.parse(r[8] || '{}')
      }))
    };
  } catch (e) { return { error: e.toString() }; }
}

function saveToContentLibrary(weeklyData) {
  try {
    const sheet = getRequestSheet(LIBRARY_SHEET_NAME);
    const existingIds = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues().flat() : [];
    
    weeklyData.forEach(w => {
      const winId = w.id + '_WIN';
      if (!existingIds.includes(winId) && w.winningContent.title) {
        sheet.appendRow([
          winId, w.id, 'IG', w.winningContent.type, w.winningContent.title, w.winningContent.link,
          JSON.stringify(w.topicTags), w.contentScore, JSON.stringify(w.metrics)
        ]);
      }
    });
  } catch (e) {}
}

// --- TRIGGERS & UTILS ---

function setupAllTriggers() {
  try {
    ScriptApp.getProjectTriggers().forEach(t => ScriptApp.deleteTrigger(t));
    ScriptApp.newTrigger('sendWeeklyReminder').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(9).create();
    ScriptApp.newTrigger('checkAndSendAnomalyAlerts').timeBased().everyDays(1).atHour(10).create();
    return "Triggers setup successfully";
  } catch (e) { return "Trigger error: " + e.toString(); }
}

function sendWeeklyReminder() {
  const email = getAdminEmail();
  MailApp.sendEmail(email, "📅 SOSBOARD - Weekly Reminder", "Update data performa minggu ini!");
}

function checkAndSendAnomalyAlerts() {
  const data = getWeeklyData();
  if (data.length < 2) return;
  const cur = data[data.length-1], prev = data[data.length-2];
  if (cur.metrics.reach < prev.metrics.reach * 0.7) {
    MailApp.sendEmail(getAdminEmail(), "⚠️ SOSBOARD Alert", `Reach turun drastis di ${cur.label}!`);
  }
}

function getAdminEmail() {
  return Session.getActiveUser().getEmail() || 'muhammadnurulqolbi.id@gmail.com';
}

function logAudit(action, details) {
  try {
    const sheet = getRequestSheet(AUDIT_SHEET_NAME);
    sheet.appendRow([new Date(), Session.getActiveUser().getEmail(), action, details]);
  } catch (e) {}
}

function getVersionHistory() {
  try {
    const sheet = getRequestSheet(AUDIT_SHEET_NAME);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];
    const data = sheet.getRange(Math.max(2, lastRow - 50), 1, Math.min(50, lastRow - 1), 4).getValues();
    return data.map(r => ({ timestamp: r[0], user: r[1], action: r[2], details: r[3] })).reverse();
  } catch (e) { return []; }
}

function getTargets() {
  try {
    const sheet = getRequestSheet(SETTINGS_SHEET_NAME);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return {};
    const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    const targets = {};
    data.forEach(r => { targets[r[0]] = { reachPct: r[1], engagementPct: r[2] }; });
    return targets;
  } catch (e) { return {}; }
}

function saveTargets(targets) {
  try {
    const sheet = getRequestSheet(SETTINGS_SHEET_NAME);
    sheet.clear();
    sheet.appendRow(['Month', 'ReachTarget', 'EngTarget']);
    const rows = Object.keys(targets).map(k => [k, targets[k].reachPct, targets[k].engagementPct]);
    if (rows.length > 0) sheet.getRange(2, 1, rows.length, 3).setValues(rows);
    return "Success";
  } catch (e) { return e.toString(); }
}

function getGrowthData() {
  try {
    const sheet = getRequestSheet(GROWTH_SHEET_NAME);
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return [];
    const data = sheet.getRange(2, 1, lastRow - 1, 6).getValues();
    return data.map(r => ({ month: r[0], instagram: r[1], tiktok: r[2], linkedin: r[3], posts: r[4], totalFollowers: r[5] }));
  } catch (e) { return []; }
}

function saveGrowthData(data) {
  try {
    const sheet = getRequestSheet(GROWTH_SHEET_NAME);
    sheet.clear();
    sheet.appendRow(['Month', 'Instagram', 'TikTok', 'LinkedIn', 'Posts', 'TotalFollowers']);
    const rows = data.map(r => [r.month, r.instagram, r.tiktok, r.linkedin, r.posts, r.totalFollowers]);
    if (rows.length > 0) sheet.getRange(2, 1, rows.length, 6).setValues(rows);
    return "Success";
  } catch (e) { return e.toString(); }
}

function getLinkPreview(url) {
  try {
    const resp = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
    if (resp.getResponseCode() !== 200) return { error: 'Fetch failed' };
    const html = resp.getContentText();
    const title = html.match(/<title>(.*?)<\/title>/)?.[1] || 'No Title';
    const image = html.match(/<meta property="og:image" content="(.*?)"/)?.[1] || '';
    return { success: true, title, image };
  } catch (e) { return { error: e.toString() }; }
}

function getDummyWeeklyData() {
  return [
    {
      id: 'w1', startDate: new Date().toISOString(), label: 'W1', dateRange: 'May 1-7',
      metrics: { reach: 5000, engagement: 500, views: 10000, profileActivity: 50, followersPct: 70, nonFollowersPct: 30, storiesPct: 20, reelsPct: 50, feedsPct: 30 },
      winningContent: { type: 'IGR', title: 'Winning Post', link: '', analysis: { what: 'Video', why: 'Hook', todo: 'Repeat' } },
      losingContent: { type: 'IGF', title: 'Losing Post', link: '', analysis: { what: 'Static', why: 'No hook', todo: 'Avoid' } },
      topicTags: ['Product'], contentScore: 8, postHour: 19, postDay: 'Wednesday', versionId: 'v1'
    }
  ];
}
