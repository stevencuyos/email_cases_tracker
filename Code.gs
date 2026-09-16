/**
 * CRITICAL CONFIGURATION:
 * This script MUST be deployed with:
 * - Execute as: "User accessing the web app"
 * - Who has access: "Anyone within [Your Organization]"
 */

// --- UTILITIES ---

// Helper function to securely log errors to a dedicated tab without crashing the script
function logError(functionName, errorMessage, userLdap) {
  try {
    const lock = LockService.getScriptLock();
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let errorSheet = ss.getSheetByName('Error_Logs');
    if (!errorSheet) {
      errorSheet = ss.insertSheet('Error_Logs');
      errorSheet.appendRow(['Timestamp', 'Function', 'User', 'Error Message']);
      errorSheet.getRange(1, 1, 1, 4).setFontWeight("bold").setBackground("#fce8e6");
      errorSheet.setFrozenRows(1);
    }
    errorSheet.appendRow([new Date(), functionName, userLdap || 'Unknown', errorMessage]);
    lock.releaseLock();
  } catch (e) {
    console.error("Failed to write to Error_Logs tab: " + e);
  }
}

// Ultra-fast date string converter to replace slow Utilities.formatDate calls in loops
// Converts a JS Date object to "M/d/yyyy" format using local timezone context
function toDateStringFast(dateObj) {
  if (!(dateObj instanceof Date)) {
    // If it's already a string, attempt a naive cleanup
    return String(dateObj).split('T')[0]; 
  }
  return Utilities.formatDate(dateObj, Session.getScriptTimeZone(), "M/d/yyyy");
}

// --- COLUMN MAP FOR Raw_Cases (0-indexed, matches getValues() output) ---
const RAW_COLS = {
  TIMESTAMP: 0, DATE: 1, INTERVAL: 2, AGENT: 3, NAME: 4, SITE: 5, LOB: 6, WORKFLOW: 7,
  SHIFT_TYPE: 8, CASE_TYPE: 9, TOTAL: 10, VALID: 11, FLAGGED: 12, CASE_IDS: 13,
  AUDIT_NOTES: 14, INTERVAL_ACTIVITY: 15, OT_TYPE: 16
};

// --- MAIN APPLICATION LOGIC ---

// 1. Serve the Web App Interface
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Case Tracking Portal')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 2. Auto-Initialize the Database Structure
function initializeDatabase() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    const sheetsConfig = {
      'Raw_Cases': ['Timestamp', 'Date', 'Interval', 'Agent', 'Name', 'Site', 'LOB', 'Workflow', 'Shift Type', 'Case Type', 'Total', 'Valid', 'Flagged', 'Case IDs', 'Audit Notes', 'Interval Activity', 'OT Type'],
      'Index_CaseIDs': ['Case ID', 'Type Logged', 'Date Logged', 'Agent'],
      'Audit Queue': ['Status', 'Timestamp', 'Agent', 'Site', 'Case Type', 'Total Logged', 'Flagged IDs', 'Audit Reason', 'Resolution', 'RawRowRef'],
      'Error_Logs': ['Timestamp', 'Function', 'User', 'Error Message']
    };

    for (const [sheetName, headers] of Object.entries(sheetsConfig)) {
      let sheet = ss.getSheetByName(sheetName);
      if (!sheet) {
        sheet = ss.insertSheet(sheetName);
        sheet.appendRow(headers);
        sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f3f3f3");
        sheet.setFrozenRows(1);
      }
    }
  } catch (e) {
    console.error("Initialization error: " + e);
  } finally {
    lock.releaseLock();
  }
}

// 3. Fetch User Profile from Masterlist Tab (With Cache)
function getUserProfile() {
  const email = Session.getActiveUser().getEmail();
  const currentLdap = email ? email.split('@')[0] : 'unknown_agent';
  
  const cache = CacheService.getUserCache();
  const cachedProfile = cache.get('userProfile_' + currentLdap);
  if (cachedProfile) {
    return JSON.parse(cachedProfile);
  }
  
  let profile = {
    ldap: currentLdap,
    name: currentLdap,
    site: 'Unknown',
    lob: '',
    workflow: '',
    role: 'Agent',
    isManager: false
  };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = ss.getSheetByName('Masterlist');
    
    if (masterSheet && masterSheet.getLastRow() > 1) {
      // Fetch from Col A (1) all the way to Col AS (45)
      const masterData = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 45).getValues();
      
      let isManager = false;
      
      for (let i = 0; i < masterData.length; i++) {
        const rowLdap = masterData[i][0]; // Col A (LDAP)
        const rowSup = masterData[i][28]; // Col AC (Supervisor LDAP)
        const rowMgr = masterData[i][36]; // Col AK (Manager LDAP)

        // Dynamic Manager Auth: If user is listed as a Sup or Mgr for ANY agent, grant access
        if (!isManager && (rowSup === currentLdap || rowMgr === currentLdap)) {
          isManager = true;
        }

        // Grab their personal demographic details
        if (rowLdap && rowLdap.toString().toLowerCase() === currentLdap.toLowerCase()) {
          profile.name = masterData[i][1] || currentLdap; // Col B (Name)
          profile.lob = masterData[i][21] || '';          // Col V (LOB)
          profile.workflow = masterData[i][22] || '';     // Col W (Workflow)
          profile.site = masterData[i][44] || 'Unknown';  // Col AS (Site)
        }
      }
      
      profile.isManager = isManager;
      if (isManager) profile.role = 'Leadership';
    }
    
    // Store in cache for 1 hour (3600 seconds)
    cache.put('userProfile_' + currentLdap, JSON.stringify(profile), 3600);
    
  } catch (e) {
    logError('getUserProfile', e.toString(), currentLdap);
  }
  
  return profile;
}

// Clears cached profile/agent-list data so access changes (new manager, new masterlist row) take effect immediately.
function clearAccessCache() {
  const email = Session.getActiveUser().getEmail();
  const currentLdap = email ? email.split('@')[0] : 'unknown_agent';
  CacheService.getUserCache().remove('userProfile_' + currentLdap);
  CacheService.getScriptCache().removeAll(['allAgentsList', 'emailAgentsList']);
  return true;
}

// Throws unless the current user is a manager. Call at the top of any manager-only function.
function requireManagerOrThrow() {
  const profile = getUserProfile();
  if (!profile.isManager) {
    throw new Error("Access denied: manager permissions required.");
  }
  return profile;
}

// 4. Process Submission with LockService (Concurrency Control)
function submitCases(formObject) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
  } catch (e) {
    return { success: false, error: "System is busy processing other submissions. Please try again in a few seconds." };
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss.getSheetByName('Raw_Cases')) initializeDatabase();

    const rawSheet = ss.getSheetByName('Raw_Cases');
    const indexSheet = ss.getSheetByName('Index_CaseIDs');
    const auditSheet = ss.getSheetByName('Audit Queue');

    const timestamp = new Date();
    const dateStr = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "M/d/yyyy");
    const intervalStr = Utilities.formatDate(timestamp, Session.getScriptTimeZone(), "h:00 a");
    
    const userProfile = getUserProfile();
    const ldap = userProfile.ldap;
    const site = userProfile.site; 
    
    const shiftType = formObject.shiftType;
    const otType = formObject.otType || '';
    const caseType = formObject.caseType;
    const intervalActivity = formObject.intervalActivity || 'Normal Production';
    const rawText = formObject.caseIdsText || '';

    const CASE_ID_PATTERN = /^\d-\d{7,14}$/;
    let rawIds = rawText.split(/[\n,;\s]+/).map(id => id.trim()).filter(id => id !== '');
    let uniqueIds = [...new Set(rawIds)];
    let malformedIds = uniqueIds.filter(id => !CASE_ID_PATTERN.test(id));
    uniqueIds = uniqueIds.filter(id => CASE_ID_PATTERN.test(id));
    let validIds = [];
    let flaggedIds = [];
    let reasonCounts = {};

    const indexData = indexSheet.getLastRow() > 1 ? indexSheet.getRange(2, 1, indexSheet.getLastRow() - 1, 4).getValues() : [];

    uniqueIds.forEach(id => {
      let isFlagged = false;
      let reason = "";
      let history = indexData.filter(row => row[0] == id);

      if (caseType === 'Regular Email (Take Next)' && history.length > 0) {
        isFlagged = true;
        reason = "Previously logged in system";
      } else if (caseType === 'Reopened Cases') {
        let todayHistory = history.filter(row => row[2] == dateStr && row[1] === 'Reopened Cases' && row[3] === ldap);
        if (todayHistory.length > 0) {
          isFlagged = true;
          reason = "Already reopened today";
        }
      }

      if (isFlagged) {
        flaggedIds.push(id);
        reasonCounts[reason] = (reasonCounts[reason] || 0) + 1;
      } else {
        validIds.push(id);
      }
    });

    let totalCount = uniqueIds.length;
    let validCount = validIds.length;
    let flaggedCount = flaggedIds.length;
    let auditNotes = flaggedCount > 0
      ? "⚠️ " + Object.entries(reasonCounts).map(([reason, count]) => `${reason} (${count})`).join(', ')
      : "Clean";

    let rawRowNumber = rawSheet.getLastRow() + 1;
    // Writes LOB, Workflow, and Name into the raw sheet, appended Interval Activity at the end
    rawSheet.appendRow([timestamp, dateStr, intervalStr, ldap, userProfile.name, site, userProfile.lob, userProfile.workflow, shiftType, caseType, totalCount, validCount, flaggedCount, uniqueIds.join(', '), auditNotes, intervalActivity, otType]);
    rawSheet.getRange(rawRowNumber, 3).setNumberFormat('@'); // Prevent Sheets from auto-converting this to a Date

    let indexDataToAppend = uniqueIds.map(id => [id, caseType, dateStr, ldap]);
    if (indexDataToAppend.length > 0) {
      indexSheet.getRange(indexSheet.getLastRow() + 1, 1, indexDataToAppend.length, 4).setValues(indexDataToAppend);
    }

    if (flaggedCount > 0) {
      auditSheet.appendRow(["🔴 PENDING", timestamp, ldap, site, caseType, totalCount, flaggedIds.join(', '), auditNotes, "", rawRowNumber]);
    }
    
    return { success: true, valid: validCount, flagged: flaggedCount, rejected: malformedIds.length };
    
  } catch (error) {
    const user = Session.getActiveUser().getEmail() || 'Unknown';
    logError('submitCases', error.toString(), user);
    return { success: false, error: "System encountered an error processing your cases. Please try again." };
  } finally {
    lock.releaseLock();
  }
}

// 5. Fetch Dashboard Data
function getDashboardData() {
  try {
    requireManagerOrThrow();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const auditSheet = ss.getSheetByName('Audit Queue');
    const rawSheet = ss.getSheetByName('Raw_Cases');

    // A. Get Pending Audits
    let audits = [];
    if (auditSheet && auditSheet.getLastRow() > 1) {
      // FIXED: Now fetching 10 columns instead of 9 to grab the shifted RawRowRef
      const auditData = auditSheet.getRange(2, 1, auditSheet.getLastRow() - 1, 10).getValues();
      audits = auditData.map((r, i) => ({
        row: i + 2, 
        status: r[0],
        timestamp: (r[1] instanceof Date) ? Utilities.formatDate(r[1], Session.getScriptTimeZone(), "h:mm a") : String(r[1]),
        agent: r[2],
        site: r[3],
        caseType: r[4],
        totalLogged: r[5],
        flaggedIds: r[6],
        reason: r[7],
        rawRowRef: r[9] // FIXED: Shifted from index 8 to index 9 (Column J)
      })).filter(a => String(a.status).includes('PENDING'));
    }

    // B. Get Today's Metrics (Grouped by Agent)
    let metrics = {};
    if (rawSheet && rawSheet.getLastRow() > 1) {
      const todayStr = toDateStringFast(new Date());
      const rawData = rawSheet.getRange(2, 1, rawSheet.getLastRow() - 1, 15).getValues(); 

      rawData.forEach(r => {
        let rowDateStr = toDateStringFast(r[RAW_COLS.DATE]);

        if (rowDateStr === todayStr) { 
          const agent = r[RAW_COLS.AGENT] ? r[RAW_COLS.AGENT].toString().trim().toLowerCase() : '';
          const site = r[RAW_COLS.SITE]; 
          const valid = Number(r[RAW_COLS.VALID]) || 0; 
          const flagged = Number(r[RAW_COLS.FLAGGED]) || 0;
          
          if (!metrics[agent]) {
            metrics[agent] = { agent: agent, site: site, totalValid: 0, totalFlagged: 0 };
          }
          metrics[agent].totalValid += valid;
          metrics[agent].totalFlagged += flagged;
        }
      });
    }

    return {
      audits: audits,
      metrics: Object.values(metrics).sort((a, b) => b.totalValid - a.totalValid) 
    };
  } catch (e) {
    const user = Session.getActiveUser().getEmail() || 'Unknown';
    logError('getDashboardData', e.toString(), user);
    throw new Error("Failed to load dashboard data. Please try refreshing."); // Throws to the frontend withFailureHandler
  }
}

// 6. Resolve Soft Audits (Approve/Reject)
function resolveAudit(auditRow, rawRowRef, resolution) {
  return resolveAuditsBulk([{ auditRow: auditRow, rawRowRef: rawRowRef }], resolution);
}

// 6b. Resolve Bulk Audits
function resolveAuditsBulk(auditsToProcess, resolution) {
  const profile = getUserProfile();
  if (!profile.isManager) {
    return { success: false, error: "Access denied: manager permissions required." };
  }

  if (!auditsToProcess || auditsToProcess.length === 0) {
    return { success: true }; // Nothing to do
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const auditSheet = ss.getSheetByName('Audit Queue');
    const rawSheet = ss.getSheetByName('Raw_Cases');

    // To avoid timeouts with many API calls, we could fetch all data and write back,
    // but these are sparse updates. Using a RangeList for simple status updates helps,
    // though we still need to calculate math for approvals.

    // We'll organize updates by row to minimize calls.
    const auditUpdates = []; // {row, status, resolution}
    const rawUpdates = [];   // {row, valid, flagged, note}

    for (const audit of auditsToProcess) {
      const { auditRow, rawRowRef } = audit;

      if (resolution === 'Approve') {
        auditUpdates.push({ row: auditRow, col1: "🟢 APPROVED", col9: "Approved" });

        // Need to read current values to do the math for Approve
        // It's still a read per row, but we can do it quickly.
        const validCount = rawSheet.getRange(rawRowRef, 12).getValue();
        const flaggedCount = rawSheet.getRange(rawRowRef, 13).getValue();
        rawUpdates.push({ row: rawRowRef, col12: validCount + flaggedCount, col13: 0, col15: "✅ Resolved by Manager" });
      } else {
        auditUpdates.push({ row: auditRow, col1: "⚫ REJECTED", col9: "Rejected" });
        rawUpdates.push({ row: rawRowRef, col12: null, col13: 0, col15: "❌ Rejected by Manager (Duplicate/Fraud)" });
      }
    }

    // Apply updates efficiently
    auditUpdates.forEach(u => {
      auditSheet.getRange(u.row, 1).setValue(u.col1);
      auditSheet.getRange(u.row, 9).setValue(u.col9);
    });

    rawUpdates.forEach(u => {
      if (u.col12 !== null) rawSheet.getRange(u.row, 12).setValue(u.col12);
      rawSheet.getRange(u.row, 13).setValue(u.col13);
      rawSheet.getRange(u.row, 15).setValue(u.col15);
    });

    return { success: true };
  } catch (e) {
    const user = Session.getActiveUser().getEmail() || 'Unknown';
    logError('resolveAuditsBulk', e.toString(), user);
    return { success: false, error: "Failed to resolve audits. Please check your connection and try again." };
  } finally {
    lock.releaseLock();
  }
}

// 7. Fetch User's Submissions for "My Submissions" Tab
function getMySubmissions(dateStr, targetLdap) {
  const realEmail = Session.getActiveUser().getEmail();
  logError('ENTRY_CHECK_V2', 'ACTUAL_LOGGED_IN_EMAIL=' + realEmail + ' | dateStr=' + dateStr + ' targetLdap=' + targetLdap, 'entry');
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rawSheet = ss.getSheetByName('Raw_Cases');
    const userProfile = getUserProfile();
    
    // If targetLdap is provided and the user is a manager, use it. Otherwise default to their own ldap.
    const queryLdap = (userProfile.isManager && targetLdap) ? targetLdap.toLowerCase() : userProfile.ldap.toLowerCase();
    
    let submissions = [];

    if (rawSheet && rawSheet.getLastRow() > 1) {
      const rawData = rawSheet.getRange(2, 1, rawSheet.getLastRow() - 1, 16).getValues();
      
      const dedupeMap = {};

      rawData.forEach(r => {
        let rowDateStr = toDateStringFast(r[1]);
        
        // r[3] is Agent LDAP
        if (rowDateStr === dateStr && r[3] && r[3].toString().trim().toLowerCase() === queryLdap.trim()) {
          const intervalVal = (r[2] instanceof Date)
            ? Utilities.formatDate(r[2], Session.getScriptTimeZone(), "h:00 a")
            : r[2];
          const caseIdsVal = r[13];

          // Normalize Case IDs for the dedupe key only (order-independent), so "A, B, C" and
          // "C, A, B" are recognized as the same submission. The original caseIdsVal (unsorted)
          // is still what gets displayed to the user.
          const normalizedIds = caseIdsVal
            ? caseIdsVal.split(',').map(id => id.trim().toLowerCase()).filter(id => id !== '').sort().join(',')
            : '';

          // Dedupe key: same interval + case type + same set of Case IDs (regardless of order)
          // = same physical submission. If it was resubmitted (and possibly resolved by a
          // manager), the later row in the sheet is authoritative, so it naturally overwrites
          // the earlier entry below.
          const dedupeKey = intervalVal + '|' + r[9] + '|' + normalizedIds;

          dedupeMap[dedupeKey] = {
            interval: intervalVal,        // Col C
            activity: r[15] || 'Normal Production', // Col P
            caseType: r[9],               // Col J
            validCount: r[11],            // Col L
            caseIds: caseIdsVal           // Col N
          };
        }
      });

      submissions = Object.values(dedupeMap);
    }
    
    return submissions;
  } catch (e) {
    const user = Session.getActiveUser().getEmail() || 'Unknown';
    logError('getMySubmissions', e.toString(), user);
    throw new Error("Unable to fetch submissions data. Check network and retry.");
  }
}

// 8. Fetch All Agent LDAPs for Autocomplete (With Cache)
function getAllAgents() {
  try {
    const cache = CacheService.getScriptCache();
    const cachedAgents = cache.get('allAgentsList');
    if (cachedAgents) {
      return JSON.parse(cachedAgents);
    }
  
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = ss.getSheetByName('Masterlist');
    let agents = new Set();
    
    if (masterSheet && masterSheet.getLastRow() > 1) {
      // Col A is LDAP
      const data = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 1).getValues();
      data.forEach(row => {
        const ldap = row[0] ? row[0].toString().trim().toLowerCase() : '';
        if (ldap) agents.add(ldap);
      });
    }
    
    let sortedAgents = Array.from(agents).sort();
    // Cache for 4 hours
    cache.put('allAgentsList', JSON.stringify(sortedAgents), 14400);
    return sortedAgents;
    
  } catch(e) {
    const user = Session.getActiveUser().getEmail() || 'Unknown';
    logError('getAllAgents', e.toString(), user);
    return [];
  }
}

// --- 10. Fetch Analytics Data ---
function getAnalyticsData(daysToFetch = 7) {
  try {
    requireManagerOrThrow();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rawSheet = ss.getSheetByName('Raw_Cases');
    const auditSheet = ss.getSheetByName('Audit Queue');

    const now = new Date();
    const todayStr = toDateStringFast(now);

    // Create cutoff date based on requested range (00:00:00 of the target day)
    const cutoffDate = new Date();
    cutoffDate.setDate(now.getDate() - (daysToFetch - 1));
    cutoffDate.setHours(0, 0, 0, 0);

    let data = {
      kpis: {
        totalValidToday: 0,
        pendingAudits: 0,
        activeStaffToday: 0,
        activeStaffThisHour: 0
      },
      heatmap: [], // Array of { name: 'Day', data: [{ x: 'Hour', y: count }] }
      workflows: {},
      sites: {},
      leaderboard: {}
    };

    // 1. Pending Audits KPI
    if (auditSheet && auditSheet.getLastRow() > 1) {
      const auditData = auditSheet.getRange(2, 1, auditSheet.getLastRow() - 1, 1).getValues();
      data.kpis.pendingAudits = auditData.filter(r => String(r[0]).includes('PENDING')).length;
    }

    if (!rawSheet || rawSheet.getLastRow() <= 1) return data;

    const rawData = rawSheet.getRange(2, 1, rawSheet.getLastRow() - 1, 12).getValues(); // Up to Col 11 (Valid)

    // Format helpers
    const currentHourStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "h:00 a");

    let activeAgentsToday = new Set();
    let activeAgentsThisHour = new Set();

    // Heatmap structure prep
    const heatmapDataMap = {}; // { '9/24/2026': { '9:00 AM': 10, ... } }

    rawData.forEach(r => {
      const rowDateObj = (r[RAW_COLS.DATE] instanceof Date) ? r[RAW_COLS.DATE] : new Date(r[RAW_COLS.DATE]);

      // Only process data within the cutoff range
      if (rowDateObj < cutoffDate) return;

      const rowDateStr = toDateStringFast(rowDateObj);
      const rowInterval = (r[RAW_COLS.INTERVAL] instanceof Date) ? Utilities.formatDate(r[RAW_COLS.INTERVAL], Session.getScriptTimeZone(), "h:00 a") : r[RAW_COLS.INTERVAL];
      const agent = r[RAW_COLS.AGENT] ? r[RAW_COLS.AGENT].toString().trim().toLowerCase() : 'unknown';
      const site = r[RAW_COLS.SITE] || 'Unknown';
      const workflow = r[RAW_COLS.CASE_TYPE] || 'Unknown';
      const validCases = Number(r[RAW_COLS.VALID]) || 0;

      // KPI: Today's Metrics
      if (rowDateStr === todayStr && validCases > 0) {
        data.kpis.totalValidToday += validCases;
        activeAgentsToday.add(agent);
        if (rowInterval === currentHourStr) {
          activeAgentsThisHour.add(agent);
        }
      }

      // Populate Heatmap Data
      if (!heatmapDataMap[rowDateStr]) heatmapDataMap[rowDateStr] = {};
      heatmapDataMap[rowDateStr][rowInterval] = (heatmapDataMap[rowDateStr][rowInterval] || 0) + validCases;

      // Populate Donut (Workflows)
      data.workflows[workflow] = (data.workflows[workflow] || 0) + validCases;

      // Populate Bar Chart (Sites)
      data.sites[site] = (data.sites[site] || 0) + validCases;

      // Populate Leaderboard
      if (validCases > 0) {
        if (!data.leaderboard[agent]) data.leaderboard[agent] = 0;
        data.leaderboard[agent] += validCases;
      }
    });

    data.kpis.activeStaffToday = activeAgentsToday.size;
    data.kpis.activeStaffThisHour = activeAgentsThisHour.size;

    // Transform Heatmap data for ApexCharts
    // Sort dates ascending
    const sortedDates = Object.keys(heatmapDataMap).sort((a, b) => new Date(a) - new Date(b));
    const allHours = ["12:00 AM","1:00 AM","2:00 AM","3:00 AM","4:00 AM","5:00 AM","6:00 AM","7:00 AM","8:00 AM","9:00 AM","10:00 AM","11:00 AM","12:00 PM","1:00 PM","2:00 PM","3:00 PM","4:00 PM","5:00 PM","6:00 PM","7:00 PM","8:00 PM","9:00 PM","10:00 PM","11:00 PM"];

    // Reverse the dates so newest is on top of the Y-axis for standard heatmap look
    sortedDates.reverse().forEach(dateStr => {
      // Get a short friendly name (e.g. "Mon, Sep 24")
      const d = new Date(dateStr);
      const friendlyName = Utilities.formatDate(d, Session.getScriptTimeZone(), "EEE, MMM d");

      const daySeries = { name: friendlyName, data: [] };
      allHours.forEach(hour => {
        daySeries.data.push({
          x: hour,
          y: heatmapDataMap[dateStr][hour] || 0
        });
      });
      data.heatmap.push(daySeries);
    });

    return data;
  } catch (e) {
    const user = Session.getActiveUser().getEmail() || 'Unknown';
    logError('getAnalyticsData', e.toString(), user);
    throw new Error("Unable to fetch analytics data. Check network and retry.");
  }
}

// 9. Fetch Data for Interval View
function getIntervalData(dateStr, intervalHourStr) {
  try {
    requireManagerOrThrow();
    // dateStr format expected: "9/15/2026"
    // intervalHourStr format expected: "16:00" (24-hour format string)
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const shiftSheet = ss.getSheetByName('Agent Shifts');
    const rawSheet = ss.getSheetByName('Raw_Cases');
    const masterSheet = ss.getSheetByName('Masterlist'); 
    
    let targetHour = parseInt(intervalHourStr.split(':')[0], 10);
    let agentsInInterval = {};
    
    // 1. Build a whitelist of "Email" Channel Agents from the Masterlist (With Cache)
    const cache = CacheService.getScriptCache();
    let emailAgentsList = cache.get('emailAgentsList');
    let emailAgents = new Set();
    
    if (emailAgentsList) {
      emailAgents = new Set(JSON.parse(emailAgentsList));
    } else if (masterSheet && masterSheet.getLastRow() > 1) {
      const masterData = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 24).getValues();
      for (let i = 0; i < masterData.length; i++) {
        const ldap = masterData[i][0] ? masterData[i][0].toString().toLowerCase() : '';
        const channel = masterData[i][23] ? masterData[i][23].toString().toLowerCase() : ''; 
        
        if (channel === 'email') {
          emailAgents.add(ldap);
        }
      }
      cache.put('emailAgentsList', JSON.stringify(Array.from(emailAgents)), 14400);
    }

    // 2. Process Regular Shifts from 'Agent Shifts'
    if (shiftSheet && shiftSheet.getLastRow() > 2) {
      const shiftData = shiftSheet.getDataRange().getValues();
      const headers = shiftData[1]; // Row 2 holds the actual dates
      
      // Find matching date column (Starts checking from Col I / Index 8)
      let dateColIdx = -1;
      for (let c = 8; c < headers.length; c++) {
        let cellDate = headers[c];
        let formattedCellDate = "";
        try {
          if (cellDate instanceof Date) {
            formattedCellDate = toDateStringFast(cellDate);
          } else if (cellDate) {
             formattedCellDate = toDateStringFast(new Date(cellDate));
          }
        } catch(e) {}
        
        if (formattedCellDate === dateStr) {
          dateColIdx = c;
          break;
        }
      }
      
      if (dateColIdx !== -1) {
        // Loop through agents starting from Row 3
        for (let r = 2; r < shiftData.length; r++) {
          const ldap = shiftData[r][0] ? shiftData[r][0].toString().trim().toLowerCase() : '';
          
          // NEW FILTER: Skip agent if they are not in the Email channel (unless they do OT later)
          if (!emailAgents.has(ldap)) {
            continue; 
          }

          const site = shiftData[r][4]; // Col E
          const shiftVal = shiftData[r][dateColIdx]; 
          
          if (shiftVal && shiftVal !== "OFF" && shiftVal !== "VL" && shiftVal !== "LOA" && shiftVal !== "AWOL") {
            let startHour = -1;
            
            if (shiftVal instanceof Date) {
              startHour = shiftVal.getHours();
            } else if (typeof shiftVal === 'string' && shiftVal.includes(':')) {
              startHour = parseInt(shiftVal.split(':')[0], 10);
            } else if (typeof shiftVal === 'number') {
              startHour = Math.round(shiftVal * 24); // Convert decimal time
            }
            
            if (startHour !== -1) {
              let endHour = startHour + 9;
              let isOnShift = false;
              
              // Handle shifts crossing midnight
              if (endHour <= 24) {
                isOnShift = (targetHour >= startHour && targetHour < endHour);
              } else {
                isOnShift = (targetHour >= startHour || targetHour < (endHour - 24));
              }
              
              if (isOnShift) {
                let formattedSOS = startHour + ":00";
                let eosActual = endHour > 24 ? endHour - 24 : endHour;
                let formattedEOS = eosActual + ":00";
                
                agentsInInterval[ldap] = {
                  ldap: ldap,
                  sos: formattedSOS,
                  eos: formattedEOS,
                  site: site,
                  isOT: false,
                  casesLogged: 0,
                  regularCount: 0,
                  manualCount: 0,
                  reopenedCount: 0
                };
              }
            }
          }
        }
      }
    }
    
    // 3. Process Overtime & Live Metrics from 'Raw_Cases'
    if (rawSheet && rawSheet.getLastRow() > 1) {
      const rawData = rawSheet.getRange(2, 1, rawSheet.getLastRow() - 1, 17).getValues(); 
      // Format target hour to match Raw_Cases "h:00 a" format (e.g. "4:00 PM").
      // We still use Utilities here because it runs exactly once per function call, not in a loop.
      let targetDateObj = new Date();
      targetDateObj.setHours(targetHour, 0, 0, 0);
      const intervalLabel = Utilities.formatDate(targetDateObj, Session.getScriptTimeZone(), "h:00 a");
      
      rawData.forEach(r => {
        let rowDateStr = toDateStringFast(r[RAW_COLS.DATE]);
        const rowInterval = (r[RAW_COLS.INTERVAL] instanceof Date)
          ? Utilities.formatDate(r[RAW_COLS.INTERVAL], Session.getScriptTimeZone(), "h:00 a")
          : r[RAW_COLS.INTERVAL];
        
        if (rowDateStr === dateStr && rowInterval === intervalLabel) {
          const ldap = r[RAW_COLS.AGENT] ? r[RAW_COLS.AGENT].toString().trim().toLowerCase() : '';
          const site = r[RAW_COLS.SITE];
          const isOvertime = r[RAW_COLS.SHIFT_TYPE] === "Overtime";
          const otType = r[RAW_COLS.OT_TYPE] || '';
          const validCases = Number(r[RAW_COLS.VALID]) || 0;
          const intervalActivity = r[RAW_COLS.INTERVAL_ACTIVITY] || 'Normal Production';
          
          // If OT agent isn't on the shift list, add them dynamically regardless of their Channel
          if (!agentsInInterval[ldap] && isOvertime) {
            agentsInInterval[ldap] = {
              ldap: ldap,
              sos: "OT",
              eos: "OT",
              site: site,
              isOT: true,
              otType: otType,
              casesLogged: 0,
              regularCount: 0,
              manualCount: 0,
              reopenedCount: 0,
              activityLogged: intervalActivity
            };
          }
          
          // Add metrics if they are in the list
          if (agentsInInterval[ldap]) {
            const caseType = r[RAW_COLS.CASE_TYPE];
            agentsInInterval[ldap].casesLogged += validCases;
            agentsInInterval[ldap].activityLogged = intervalActivity; // Track the activity they submitted

            if (caseType === 'Manual Assignment') {
              agentsInInterval[ldap].manualCount += validCases;
            } else if (caseType === 'Reopened Cases') {
              agentsInInterval[ldap].reopenedCount += validCases;
            } else {
              // Regular Email (Take Next), Telus Cases, and Cimba Cases are combined
              agentsInInterval[ldap].regularCount += validCases;
            }
          }
        }
      });
    }
    
    // 4. Compute Status
    const results = Object.values(agentsInInterval).sort((a, b) => a.ldap.localeCompare(b.ldap));
    
    results.forEach(agent => {
      let computedStatus = "";
      let sosHour = -1;
      let eosHour = -1;
      
      if (agent.sos !== "OT") {
        sosHour = parseInt(agent.sos.split(':')[0], 10);
        eosHour = parseInt(agent.eos.split(':')[0], 10);
      }
      
      if (agent.activityLogged === 'Coaching/Training') {
        computedStatus = "on Coaching/Training";
      } else if (agent.activityLogged === 'Break/Lunch' && agent.casesLogged >= 3) {
        computedStatus = "Break - 3";
      } else if (agent.reopenedCount > 0 && agent.reopenedCount === agent.casesLogged && agent.casesLogged >= 7) {
        computedStatus = "Closing Reopens";
      } else if (agent.casesLogged >= 7) {
        computedStatus = "Assigned - 7";
      } else if (!agent.isOT && targetHour === sosHour) {
        computedStatus = "SKIP SOS";
      } else if (!agent.isOT && targetHour === (eosHour - 1)) {
        computedStatus = "SKIP - EOS";
      }
      
      agent.computedStatus = computedStatus;
    });

    return results;
  } catch (e) {
    const user = Session.getActiveUser().getEmail() || 'Unknown';
    logError('getIntervalData', e.toString(), user);
    throw new Error("Unable to fetch interval schedule. Please try again.");
  }
}
