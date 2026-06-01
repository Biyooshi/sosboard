/**
 * BACKEND GOOGLE APPS SCRIPT
 * Dashboard Monitoring Content JOY
 */

const DB_SHEET_NAME = 'Database_JOY';
const SETTINGS_SHEET_NAME = 'Settings_JOY';
const GROWTH_SHEET_NAME = 'Growth_JOY';

// SPREADSHEET ID - Target spreadsheet untuk data dummy
const TARGET_SPREADSHEET_ID = '1NA0kJyEypr25RwqeNjM_B5iLIdH4OEDLh3qU6GwHg_c';

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
        'reach', 'engagement', 'views', 'profileActi
<truncated 45216 bytes>
, losingContent: { type: 'IGF', title: 'Year End Sale', thumbnail: '', link: 'https://www.instagram.com/p/Cz666yza/', analysis: { what: 'Promo repetitif', why: 'Fatigue', todo: 'Kurangi frekuensi' } } },
    { id: 'week-5-dec-2025', startDate: '2025-12-28', label: 'WEEK 5 DEC', dateRange: 'DEC 28 - DEC 31', metrics: { reach: 60000, engagement: 9000, views: 80000, profileActivity: 1000, followersPct: 20.0, nonFollowersPct: 80.0, storiesPct: 20, reelsPct: 60, feedsPct: 20 }, winningContent: { type: 'IGS', title: 'Last Post 2025', thumbnail: '', link: 'https://www.instagram.com/p/Cz777abc/', analysis: { what: 'Closing year', why: 'Emotional', todo: 'Prepare 2026' } }, losingContent: { type: 'IGF', title: 'Random Post', thumbnail: '', link: 'https://www.instagram.com/p/Cz888def/', analysis: { what: 'Filler content', why: 'Low effort', todo: 'Plan better' } } },

    // JANUARY 2026
    { id: 'week-1-jan-2026', startDate: '2026-01-01', label: 'WEEK 1 JAN', dateRange: 'JAN 01 - JAN 03', metrics: { reach: 50000, engagement: 7000, views: 60000, profileActivity: 1000, followersPct: 21.0, nonFollowersPct: 79.0, storiesPct: 18, reelsPct: 68, feedsPct: 14 }, winningContent: { type: 'IGR', title: 'New Year Goals', thumbnail: '', link: 'https://www.instagram.com/p/Ca111abc/', analysis: { what: 'Tips planning', why: 'Actionable', todo: 'Follow up' } }, losingContent: { type: 'IGS', title: 'Generic Post', thumbnail: '', link: 'https://www.instagram.com/p/Ca222def/', analysis: { what: 'Happy New Year generic', why: 'No value', todo: 'Personalize' } } },
    { id: 'week-2-jan-2026', startDate: '2026-01-04', label: 'WEEK 2 JAN', dateRange: 'JAN 04 - JAN 10', metrics: { reach: 135000, engagement: 20250, views: 195000, profileActivity: 3600, followersPct: 22.0, nonFollowersPct: 78.0, storiesPct: 17, reelsPct: 69, feedsPct: 14 }, winningContent: { type: 'IGF', title: 'Success S
<truncated 153298 bytes>

NOTE: The output was truncated because it was too long. Use a more targeted query or a smaller range to get the information you need.