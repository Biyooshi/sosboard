/**
 * BACKEND GOOGLE APPS SCRIPT
 * Dashboard Monitoring Content JOY
 */

const DB_SHEET_NAME = 'Database_JOY';
const SETTINGS_SHEET_NAME = 'Settings_JOY';
const GROWTH_SHEET_NAME = 'Growth_JOY';

// SPREADSHEET ID - Target spreadsheet untuk data dummy
const TARGET_SPREADSHEET_ID = '1NA0kJyEypr25RwqeNjM_B5iLIdH4OEDLh3qU6GwHg_c';
// API Key loaded from Script Properties (to prevent public exposure on GitHub)
const GEMINI_API_KEY = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');

function doGet() {
  return HtmlService.createTemplateFromFile('index')
      .evaluate()
      .setTitle('Dashboard Monitoring Content JOY')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// --- HELPER: Dynamic Column Mapping ---
function getColumnIndex(headerRow, columnName) {
  const index = headerRow.indexOf(columnName);
  if (index === -1) {
    Logger.log(`⚠️ Column '${columnName}' not found in header row`);
    return -1;
  }
  return index;
}

// --- API: GET DATA (OPTIMIZED WITH ERROR HANDLING) ---
function getWeeklyData() {
  try {
    Logger.log("📊 [getWeeklyData] Starting data fetch...");
    const startTime = new Date().getTime();
    
    // START FIX: Use TARGET SPREADSHEET ID for reading data
    let ss, sheet;
    try {
      ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
      sheet = ss.getSheetByName(DB_SHEET_NAME);
      Logger.log("✅ [getWeeklyData] Using TARGET spreadsheet");
    } catch (e) {
      Logger.log("⚠️ [getWeeklyData] Cannot access target spreadsheet, using active spreadsheet: " + e.toString());
      ss = SpreadsheetApp.getActiveSpreadsheet();
      sheet = ss.getSheetByName(DB_SHEET_NAME);
    }
    // END FIX
    
    if (!sheet) {
      Logger.log("⚠️ [getWeeklyData] Sheet not found, creating new sheet...");
      sheet = ss.insertSheet(DB_SHEET_NAME);
      const headers = [
        'id', 'startDate', 'label', 'dateRange', 
        'reach', 'engagement', 'views', 'profileActivity',
        'followersPct', 'nonFollowersPct', 'storiesPct', 'reelsPct', 'feedsPct',
        'win_type', 'win_title', 'win_link', 'win_what', 'win_why', 'win_todo',
        'lose_type', 'lose_title', 'lose_link', 'lose_what', 'lose_why', 'lose_todo'
      ];
      sheet.appendRow(headers);
      Logger.log("✅ [getWeeklyData] New sheet created, returning empty array");
      return []; // Return empty array - let user add data through UI
    }
    
    // PERFORMANCE OPTIMIZATION: Only fetch data rows, not entire sheet
    const lastRow = sheet.getLastRow();
    
    // SELF-SEEDING CHECK: If sheet is empty (or has only header), auto-populate with dummy data
    if (lastRow <= 1) {
      Logger.log("ℹ️ [getWeeklyData] Sheet is empty. Auto-populating dummy data...");
      try {
        const dummyData = getDummyWeeklyData4Months();
        saveWeeklyDataToTargetSpreadsheet(ss, dummyData);
        Logger.log("✅ [getWeeklyData] Auto-population successful. Returning new data.");
        return dummyData;
      } catch (seedError) {
        Logger.log("❌ [getWeeklyData] Auto-population failed: " + seedError.toString());
        return [];
      }
    }
    
    const numRows = lastRow - 1; 
    const numCols = sheet.getLastColumn();
    Logger.log(`📈 [getWeeklyData] Fetching ${numRows} rows x ${numCols} columns`);
    
    const headerRow = sheet.getRange(1, 1, 1, numCols).getValues()[0];
    const data = sheet.getRange(2, 1, numRows, numCols).getValues();
    Logger.log(`✅ [getWeeklyData] Data fetched successfully: ${data.length} rows`); 
    
    const getIdx = (colName) => getColumnIndex(headerRow, colName);
    
    const parsedData = data.map((row, rowIndex) => {
      try {
        const idIdx = getIdx('id');
        const id = idIdx >= 0 ? String(row[idIdx] || '').trim() : '';
        if (!id) return null; // Skip empty rows
        
        return {
          id: id,
          startDate: String(row[getIdx('startDate')] || '').trim(),
          label: String(row[getIdx('label')] || '').trim(),
          dateRange: String(row[getIdx('dateRange')] || '').trim(),
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
            type: String(row[getIdx('win_type')] || 'IGR').trim(),
            title: String(row[getIdx('win_title')] || '').trim(),
            link: String(row[getIdx('win_link')] || '').trim(),
            analysis: {
              what: String(row[getIdx('win_what')] || '').trim(),
              why: String(row[getIdx('win_why')] || '').trim(),
              todo: String(row[getIdx('win_todo')] || '').trim()
            }
          },
          losingContent: {
            type: String(row[getIdx('lose_type')] || 'IGR').trim(),
            title: String(row[getIdx('lose_title')] || '').trim(),
            link: String(row[getIdx('lose_link')] || '').trim(),
            analysis: {
              what: String(row[getIdx('lose_what')] || '').trim(),
              why: String(row[getIdx('lose_why')] || '').trim(),
              todo: String(row[getIdx('lose_todo')] || '').trim()
            }
          }
        };
      } catch (e) {
        Logger.log(`❌ [getWeeklyData] Error parsing row ${rowIndex + 2}: ${e.toString()}`);
        return null;
      }
    }).filter(item => item !== null);
    
    const endTime = new Date().getTime();
    Logger.log(`✅ [getWeeklyData] Completed in ${endTime - startTime}ms. Returning ${parsedData.length} records`);
    return parsedData;
  } catch (e) {
    Logger.log(`❌ [getWeeklyData] CRITICAL ERROR: ${e.toString()}`);
    return [];
  }
}

// --- API: SAVE DATA ---
function saveWeeklyData(jsonData) {
  try {
    Logger.log("💾 [saveWeeklyData] Starting save operation...");
    
    let targetSS, targetSheet;
    try {
      targetSS = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
      targetSheet = targetSS.getSheetByName(DB_SHEET_NAME);
      if (!targetSheet) targetSheet = targetSS.insertSheet(DB_SHEET_NAME);
    } catch (e) {
      targetSS = SpreadsheetApp.getActiveSpreadsheet();
      targetSheet = targetSS.getSheetByName(DB_SHEET_NAME);
      if (!targetSheet) targetSheet = targetSS.insertSheet(DB_SHEET_NAME);
    }
    
    let sheet = targetSheet;

    const expectedHeaders = [
      'id', 'startDate', 'label', 'dateRange', 
      'reach', 'engagement', 'views', 'profileActivity',
      'followersPct', 'nonFollowersPct', 'storiesPct', 'reelsPct', 'feedsPct',
      'win_type', 'win_title', 'win_link', 'win_what', 'win_why', 'win_todo',
      'lose_type', 'lose_title', 'lose_link', 'lose_what', 'lose_why', 'lose_todo'
    ];
    
    const currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const headersMatch = currentHeaders.length === expectedHeaders.length && currentHeaders[0] === 'id';
    
    if (!headersMatch) {
      sheet.clear();
      sheet.appendRow(expectedHeaders);
    }

    if (!jsonData || jsonData.length === 0) {
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
      return "Data Cleared";
    }

    const lastRow = sheet.getLastRow();
    if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);

    const rows = jsonData.map(item => [
      String(item.id || '').trim(),
      String(item.startDate || '').trim(),
      String(item.label || '').trim(),
      String(item.dateRange || '').trim(),
      Number(item.metrics?.reach) || 0,
      Number(item.metrics?.engagement) || 0,
      Number(item.metrics?.views) || 0,
      Number(item.metrics?.profileActivity) || 0,
      Number(item.metrics?.followersPct) || 0,
      Number(item.metrics?.nonFollowersPct) || 0,
      Number(item.metrics?.storiesPct) || 0,
      Number(item.metrics?.reelsPct) || 0,
      Number(item.metrics?.feedsPct) || 0,
      String(item.winningContent?.type || 'IGR').trim(),
      String(item.winningContent?.title || '').trim(),
      String(item.winningContent?.link || '').trim(),
      String(item.winningContent?.analysis?.what || '').trim(),
      String(item.winningContent?.analysis?.why || '').trim(),
      String(item.winningContent?.analysis?.todo || '').trim(),
      String(item.losingContent?.type || 'IGR').trim(),
      String(item.losingContent?.title || '').trim(),
      String(item.losingContent?.link || '').trim(),
      String(item.losingContent?.analysis?.what || '').trim(),
      String(item.losingContent?.analysis?.why || '').trim(),
      String(item.losingContent?.analysis?.todo || '').trim()
    ]);

    if (rows.length > 0) {
      sheet.getRange(2, 1, rows.length, expectedHeaders.length).setValues(rows);
      SpreadsheetApp.flush();
      return "Success - Saved " + rows.length + " weeks";
    }
    return "Success - No data to save";
  } catch (e) {
    return "Error: " + e.toString();
  }
}

function saveWeeklyDataToTargetSpreadsheet(targetSS, jsonData) {
  try {
    let sheet = targetSS.getSheetByName(DB_SHEET_NAME);
    if (!sheet) sheet = targetSS.insertSheet(DB_SHEET_NAME);
    
    const expectedHeaders = [
      'id', 'startDate', 'label', 'dateRange', 
      'reach', 'engagement', 'views', 'profileActivity',
      'followersPct', 'nonFollowersPct', 'storiesPct', 'reelsPct', 'feedsPct',
      'win_type', 'win_title', 'win_link', 'win_what', 'win_why', 'win_todo',
      'lose_type', 'lose_title', 'lose_link', 'lose_what', 'lose_why', 'lose_todo'
    ];
    
    sheet.getRange(1, 1, 1, expectedHeaders.length).setValues([expectedHeaders]);
    if (!jsonData || jsonData.length === 0) return "Data Cleared";

    const rows = jsonData.map(item => [
      String(item.id || '').trim(),
      String(item.startDate || '').trim(),
      String(item.label || '').trim(),
      String(item.dateRange || '').trim(),
      Number(item.metrics?.reach) || 0,
      Number(item.metrics?.engagement) || 0,
      Number(item.metrics?.views) || 0,
      Number(item.metrics?.profileActivity) || 0,
      Number(item.metrics?.followersPct) || 0,
      Number(item.metrics?.nonFollowersPct) || 0,
      Number(item.metrics?.storiesPct) || 0,
      Number(item.metrics?.reelsPct) || 0,
      Number(item.metrics?.feedsPct) || 0,
      String(item.winningContent?.type || 'IGR').trim(),
      String(item.winningContent?.title || '').trim(),
      String(item.winningContent?.link || '').trim(),
      String(item.winningContent?.analysis?.what || '').trim(),
      String(item.winningContent?.analysis?.why || '').trim(),
      String(item.winningContent?.analysis?.todo || '').trim(),
      String(item.losingContent?.type || 'IGR').trim(),
      String(item.losingContent?.title || '').trim(),
      String(item.losingContent?.link || '').trim(),
      String(item.losingContent?.analysis?.what || '').trim(),
      String(item.losingContent?.analysis?.why || '').trim(),
      String(item.losingContent?.analysis?.todo || '').trim()
    ]);

    if (rows.length > 0) {
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) sheet.deleteRows(2, lastRow - 1);
      sheet.getRange(2, 1, rows.length, expectedHeaders.length).setValues(rows);
      SpreadsheetApp.flush();
      return "Success";
    }
    return "Success - No rows";
  } catch (e) {
    throw e;
  }
}

// --- API: SETTINGS ---
function getTargets() {
  try {
    let ss, sheet;
    try {
      ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
      sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
    } catch (e) {
      ss = SpreadsheetApp.getActiveSpreadsheet();
      sheet = ss.getSheetByName(SETTINGS_SHEET_NAME);
    }
    if (!sheet) return createDefaultTargetsSheet(ss);
    
    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) return createDefaultTargetsSheet(ss, sheet);
    
    const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
    const targets = {};
    const monthNames = ['January','February','March','April','May','June','July','August','September','October','November','December'];

    data.forEach(row => {
      let rawMonth = row[0];
      let monthKey = "";
      if (rawMonth instanceof Date) {
        monthKey = `${monthNames[rawMonth.getMonth()]} ${rawMonth.getFullYear()}`;
      } else {
        const str = String(rawMonth).trim();
        if (str.includes('GMT') || str.match(/[A-Z][a-z]{2} [A-Z][a-z]{2} \d{2}/)) {
           try {
             const d = new Date(str);
             if (!isNaN(d.getTime())) monthKey = `${monthNames[d.getMonth()]} ${d.getFullYear()}`;
             else monthKey = str;
           } catch(e) { monthKey = str; }
        } else {
          monthKey = str;
        }
      }
      if (monthKey) {
        targets[monthKey] = {
          reachPct: Number(row[1]) || 15,
          engagementPct: Number(row[2]) || 12
        };
      }
    });
    return targets;
  } catch (e) {
    return {};
  }
}

function createDefaultTargetsSheet(ss, existingSheet = null) {
   const sheet = existingSheet || ss.insertSheet(SETTINGS_SHEET_NAME);
   sheet.clear();
   sheet.appendRow(['Month', 'ReachTarget', 'EngTarget']);
   const defaults = [];
   const months = ['January','February','March','April','May','June','July','August','September','October','November','December'];
   for (let y = 2025; y <= 2026; y++) {
     const startM = (y === 2025) ? 9 : 0; 
     const endM = 11;
     for (let m = startM; m <= endM; m++) {
       defaults.push([`${months[m]} ${y}`, 15, 12]);
     }
   }
   if (defaults.length > 0) sheet.getRange(2, 1, defaults.length, 3).setValues(defaults);
   SpreadsheetApp.flush();
   const t = {};
   defaults.forEach(d => { t[d[0]] = { reachPct: d[1], engagementPct: d[2] }; });
   return t;
}

function saveTargets(monthlyTargets) {
  try {
    let targetSS, targetSheet;
    try {
      targetSS = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
      targetSheet = targetSS.getSheetByName(SETTINGS_SHEET_NAME);
      if (!targetSheet) targetSheet = targetSS.insertSheet(SETTINGS_SHEET_NAME);
    } catch (e) {
      targetSS = SpreadsheetApp.getActiveSpreadsheet();
      targetSheet = targetSS.getSheetByName(SETTINGS_SHEET_NAME);
      if (!targetSheet) targetSheet = targetSS.insertSheet(SETTINGS_SHEET_NAME);
    }
    
    const sheet = targetSheet;
    const monthNames = ['January','February','March','April','May','June','July','August','September','October','November','December'];
    let cleanRows = [];

    Object.keys(monthlyTargets).forEach(key => {
       let dateObj = null;
       const strKey = String(key).trim();
       const parts = strKey.split(' ');
       const firstPart = parts[0];
       
       if (monthNames.includes(firstPart) && parts.length === 2 && !isNaN(parts[1])) {
          dateObj = new Date(parseInt(parts[1]), monthNames.indexOf(firstPart), 1);
       } else {
          const parsed = new Date(strKey);
          if (!isNaN(parsed.getTime())) dateObj = parsed;
       }
       
       if (dateObj) {
          const cleanLabel = `${monthNames[dateObj.getMonth()]} ${dateObj.getFullYear()}`;
          const val = monthlyTargets[key];
          cleanRows.push({
             timestamp: dateObj.getTime(), 
             label: cleanLabel,
             reach: Number(val.reachPct) || 15,
             eng: Number(val.engagementPct) || 12
          });
       }
    });

    cleanRows.sort((a, b) => a.timestamp - b.timestamp);
    const uniqueMap = new Map();
    cleanRows.forEach(item => uniqueMap.set(item.label, [item.label, item.reach, item.eng]));
    const finalRows = Array.from(uniqueMap.values());
    
    if (finalRows.length > 0) {
      sheet.clear();
      sheet.appendRow(['Month', 'ReachTarget', 'EngTarget']);
      sheet.getRange(2, 1, finalRows.length, 3).setNumberFormat("@"); 
      sheet.getRange(2, 1, finalRows.length, 3).setValues(finalRows);
      SpreadsheetApp.flush();
      return "Success";
    }
    return "No valid data to save";
  } catch (e) {
    throw e;
  }
}

function createPdfReport() {
  const ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
  const ssId = TARGET_SPREADSHEET_ID;
  const sheet = ss.getSheetByName(DB_SHEET_NAME);
  if (!sheet) throw new Error("Database sheet not found");
  
  const sheetId = sheet.getSheetId();
  const url = "https://docs.google.com/spreadsheets/d/" + ssId + "/export" +
              "?format=pdf&size=legal&portrait=false&gridlines=true&fitw=true&gid=" + sheetId;
  const params = {
    method: "GET",
    headers: { "Authorization": "Bearer " + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true
  };
  const response = UrlFetchApp.fetch(url, params);
  const blob = response.getBlob().setName("JOY_Content_Report.pdf");
  const file = DriveApp.createFile(blob);
  return file.getUrl();
}

// --- API: GROWTH DATA ---
function getGrowthData() {
  try {
    let ss, sheet;
    try {
      ss = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
      sheet = ss.getSheetByName('Growth_JOY');
    } catch (e) {
      ss = SpreadsheetApp.getActiveSpreadsheet();
      sheet = ss.getSheetByName('Growth_JOY');
    }
    if (!sheet) {
      sheet = ss.insertSheet('Growth_JOY');
      sheet.appendRow(['Month', 'Instagram', 'TikTok', 'LinkedIn', 'Posts', 'TotalFollowers']);
      return [];
    }

    const lastRow = sheet.getLastRow();
    if (lastRow <= 1) {
      try {
        const dummyData = getDummyGrowthData4Months();
        saveGrowthDataToTargetSpreadsheet(ss, dummyData);
        return dummyData;
      } catch (e) { return []; }
    }
    
    const numRows = lastRow - 1;
    const numCols = sheet.getLastColumn();
    const data = sheet.getRange(2, 1, numRows, numCols).getValues();
    const headers = sheet.getRange(1, 1, 1, numCols).getValues()[0];
    const hasPosts = headers.length > 5 || (headers.length === 5 && headers[4] === 'Posts');
    
    const growthData = data.map(row => {
      try {
        const monthStr = String(row[0]);
        let month = monthStr;
        if (monthStr.includes('GMT') || monthStr.includes('T00:00:00')) {
          try {
            const date = new Date(monthStr);
            const monthNames = ['January','February','March','April','May','June','July','August','September','October','November','December'];
            month = `${monthNames[date.getMonth()]} ${date.getFullYear()}`;
          } catch(e) { month = monthStr; }
        }
        return {
          month: month,
          instagram: Number(row[1]) || 0,
          tiktok: Number(row[2]) || 0,
          linkedin: Number(row[3]) || 0,
          posts: hasPosts ? (Number(row[4]) || 0) : 0,
          totalFollowers: hasPosts ? (Number(row[5]) || 0) : (Number(row[4]) || 0)
        };
      } catch (e) { return null; }
    }).filter(item => item !== null);
    
    return growthData;
  } catch (e) { return []; }
}

function saveGrowthData(growthData) {
  try {
    let targetSS, targetSheet;
    try {
      targetSS = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);
      targetSheet = targetSS.getSheetByName(GROWTH_SHEET_NAME);
      if (!targetSheet) targetSheet = targetSS.insertSheet(GROWTH_SHEET_NAME);
    } catch (e) {
      targetSS = SpreadsheetApp.getActiveSpreadsheet();
      targetSheet = targetSS.getSheetByName(GROWTH_SHEET_NAME);
      if (!targetSheet) targetSheet = targetSS.insertSheet(GROWTH_SHEET_NAME);
    }
    
    let sheet = targetSheet;
    sheet.clear();
    sheet.appendRow(['Month', 'Instagram', 'TikTok', 'LinkedIn', 'Posts', 'TotalFollowers']);
    
    if (!growthData || growthData.length === 0) return "Data Cleared";
    
    const rows = growthData.map(item => [
      item.month, item.instagram, item.tiktok, item.linkedin, item.posts || 0, item.totalFollowers
    ]);
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    SpreadsheetApp.flush();
    return "Success";
  } catch (e) { return "Error: " + e.toString(); }
}

function saveGrowthDataToTargetSpreadsheet(targetSS, growthData) {
  try {
    let sheet = targetSS.getSheetByName(GROWTH_SHEET_NAME);
    if (!sheet) {
      sheet = targetSS.insertSheet(GROWTH_SHEET_NAME);
      sheet.appendRow(['Month', 'Instagram', 'TikTok', 'LinkedIn', 'Posts', 'TotalFollowers']);
    }
    sheet.clear();
    sheet.appendRow(['Month', 'Instagram', 'TikTok', 'LinkedIn', 'Posts', 'TotalFollowers']);
    if (!growthData || growthData.length === 0) return "Data Cleared";
    
    const rows = growthData.map(item => [
      item.month, item.instagram, item.tiktok, item.linkedin, item.posts || 0, item.totalFollowers
    ]);
    sheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
    SpreadsheetApp.flush();
    return "Success";
  } catch (e) { throw e; }
}

// --- AI FEATURES (GEMINI API) ---

function callGeminiAPI(prompt) {
  if (!GEMINI_API_KEY) return { error: "API Key not found" };
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key=${GEMINI_API_KEY}`;
  const payload = {
    contents: [{ parts: [{ text: prompt }] }],
    generationConfig: { temperature: 0.7 }
  };
  const options = {
    method: "post",
    contentType: "application/json",
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };
  try {
    const response = UrlFetchApp.fetch(url, options);
    const json = JSON.parse(response.getContentText());
    if (json.error) return { error: json.error.message };
    const text = json.candidates[0].content.parts[0].text;
    return { success: true, text: text };
  } catch (e) {
    return { error: e.toString() };
  }
}

function generateAIInsight(weeklyData) {
  if (!weeklyData || weeklyData.length === 0) return { success: false, data: null };
  const recentData = weeklyData.slice(-4);
  const prompt = `You are a social media expert analyzing the following weekly performance data for the last ${recentData.length} weeks.
Data: ${JSON.stringify(recentData)}
Provide an analysis in JSON format exactly like this:
{
  "summary": "Brief overall performance summary.",
  "win_pattern": "What type of content is winning right now?",
  "avoid": "What to avoid based on losing content?",
  "recommendations": ["Rec 1", "Rec 2"]
}
Return ONLY valid JSON.`;
  const result = callGeminiAPI(prompt);
  if (result.error) return { success: false, error: result.error };
  try {
    let rawStr = result.text.replace(/```json/g, "").replace(/```/g, "").trim();
    return { success: true, data: JSON.parse(rawStr) };
  } catch (e) {
    return { success: false, error: "Failed to parse JSON" };
  }
}

function auditContent(hook, caption) {
  const prompt = `You are an expert social media copywriter.
Please audit the following content draft:
Hook/Title: ${hook}
Caption: ${caption}
Provide an audit in JSON format exactly like this:
{
  "score": 8,
  "feedback": "Your overall feedback.",
  "hook_analysis": "Feedback on hook.",
  "caption_analysis": "Feedback on caption clarity and CTA.",
  "improvements": ["Idea 1", "Idea 2"]
}
Return ONLY valid JSON.`;
  const result = callGeminiAPI(prompt);
  if (result.error) return { success: false, error: result.error };
  try {
    let rawStr = result.text.replace(/```json/g, "").replace(/```/g, "").trim();
    return { success: true, data: JSON.parse(rawStr) };
  } catch (e) {
    return { success: false, error: "Failed to parse JSON" };
  }
}

function askAIAgent(question, contextData) {
  const prompt = `You are an AI assistant for a social media dashboard.
Context Data: ${JSON.stringify(contextData)}
User Question: ${question}
Provide a short, helpful, and insightful answer based on the data provided. Answer directly without JSON formatting.`;
  const result = callGeminiAPI(prompt);
  if (result.error) return { success: false, answer: "Sorry, I encountered an error: " + result.error };
  return { success: true, answer: result.text };
}

function getDummyGrowthData4Months() {
  return [
    { month: 'October 2025', instagram: 1200, tiktok: 800, linkedin: 300, posts: 15, totalFollowers: 2300 },
    { month: 'November 2025', instagram: 1350, tiktok: 950, linkedin: 350, posts: 18, totalFollowers: 2650 },
    { month: 'December 2025', instagram: 1500, tiktok: 1200, linkedin: 400, posts: 12, totalFollowers: 3100 },
    { month: 'January 2026', instagram: 1650, tiktok: 1400, linkedin: 450, posts: 20, totalFollowers: 3500 }
  ];
}

function getDummyWeeklyData4Months() {
  return [
    {
      id: 'week-1-oct-2025', startDate: '2025-10-01', label: 'WEEK 1 OCT', dateRange: 'OCT 01 - OCT 04', 
      metrics: { reach: 45000, engagement: 5000, views: 60000, profileActivity: 900, followersPct: 18.5, nonFollowersPct: 81.5, storiesPct: 22, reelsPct: 55, feedsPct: 23 }, 
      winningContent: { type: 'IGR', title: 'Tips Digital Marketing 2026', link: 'https://www.instagram.com/p/Cx123abc/', analysis: { what: 'Reels prediksi', why: 'Konten edukatif', todo: 'Buat part 2' } }, 
      losingContent: { type: 'IGS', title: 'Promo Flash Sale', link: 'https://www.instagram.com/p/Cx456def/', analysis: { what: 'Story jualan', why: 'Terlalu banyak teks', todo: 'Gunakan visual' } }
    },
    {
      id: 'week-2-oct-2025', startDate: '2025-10-05', label: 'WEEK 2 OCT', dateRange: 'OCT 05 - OCT 11', 
      metrics: { reach: 132000, engagement: 19800, views: 195000, profileActivity: 3500, followersPct: 21.5, nonFollowersPct: 78.5, storiesPct: 17, reelsPct: 70, feedsPct: 13 }, 
      winningContent: { type: 'IGF', title: 'Infografis Carousel Strategy Marketing 2026', link: 'https://www.instagram.com/p/Cx789ghi/', analysis: { what: 'Carousel 10 slide', why: 'Save-able content', todo: 'Repurpose' } }, 
      losingContent: { type: 'LKDN', title: 'Corporate Update', link: 'https://www.linkedin.com/posts/activity-123', analysis: { what: 'Teks panjang', why: 'Membosankan', todo: 'Persingkat caption' } }
    }
  ];
}