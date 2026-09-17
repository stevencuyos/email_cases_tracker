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

// ISO (yyyy-MM-dd) form, needed only where the client builds its own date keys
// (e.g. calendar heatmap) since toDateStringFastClient() on the frontend uses this format.
function toISODateStringFast(dateObj) {
  const y = dateObj.getFullYear();
  let m = dateObj.getMonth() + 1;
  let d = dateObj.getDate();
  if (m < 10) m = '0' + m;
  if (d < 10) d = '0' + d;
  return y + '-' + m + '-' + d;
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
      'Error_Logs': ['Timestamp', 'Function', 'User', 'Error Message'],
      'Interval_Status': ['Date', 'Interval', 'LDAP', 'Status', 'SetBy', 'Timestamp'],
      'Interval_CheckIns': ['Date', 'Interval', 'ScheduledPOC', 'CheckedInBy', 'Timestamp', 'Result', 'UnresolvedLDAPs', 'Notes'],
      'Escalation_Log': ['Date', 'Interval', 'EscalatedAt', 'ScheduledPOC', 'UnresolvedCount', 'TotalAgents'],
      'Escalation_Config': ['Role', 'Name', 'Email']
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
    
    let isManager = false;

    // 1. Check Roster Tab for Manager Auth (Column B)
    const rosterSheet = ss.getSheetByName('Roster');
    if (rosterSheet && rosterSheet.getLastRow() > 1) {
      const rosterData = rosterSheet.getRange(2, 2, rosterSheet.getLastRow() - 1, 1).getValues(); // Col B
      for (let i = 0; i < rosterData.length; i++) {
        if (rosterData[i][0] && rosterData[i][0].toString().toLowerCase() === currentLdap.toLowerCase()) {
          isManager = true;
          break;
        }
      }
    }

    // 2. Fetch demographic details from Masterlist Tab
    if (masterSheet && masterSheet.getLastRow() > 1) {
      // Fetch from Col A (1) all the way to Col AS (45)
      const masterData = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 45).getValues();
      
      for (let i = 0; i < masterData.length; i++) {
        const rowLdap = masterData[i][0]; // Col A (LDAP)

        // Grab their personal demographic details
        if (rowLdap && rowLdap.toString().toLowerCase() === currentLdap.toLowerCase()) {
          profile.name = masterData[i][1] || currentLdap; // Col B (Name)
          profile.lob = masterData[i][21] || '';          // Col V (LOB)
          profile.workflow = masterData[i][22] || '';     // Col W (Workflow)
          profile.site = masterData[i][44] || 'Unknown';  // Col AS (Site)
          break; // Found demographic details, no need to keep scanning
        }
      }
    }
    
    profile.isManager = isManager;
    if (isManager) profile.role = 'Leadership';

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
  CacheService.getScriptCache().removeAll(['allAgentsList', 'emailAgentsList', 'agentDemographicsList', 'emailAgentsMap_v2']);
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
        rawTs: (r[1] instanceof Date) ? r[1].getTime() : null,
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


// --- ESCALATION SWEEP LOGIC ---

function setupEscalationSweepTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'escalationSweep') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  const props = PropertiesService.getScriptProperties();
  const freqStr = props.getProperty('SWEEP_FREQUENCY_MINUTES') || '15';
  const freq = parseInt(freqStr, 10);

  ScriptApp.newTrigger('escalationSweep')
    .timeBased()
    .everyMinutes(freq)
    .create();
}

function escalationSweep() {
  try {
    const lock = LockService.getScriptLock();
    if (!lock.tryLock(5000)) return;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const escalationLogSheet = ss.getSheetByName('Escalation_Log');
    if (!escalationLogSheet) return;

    const props = PropertiesService.getScriptProperties();
    const graceMinutes = parseInt(props.getProperty('GRACE_PERIOD_MINUTES') || '30', 10);
    const now = new Date();

    for (let offset = 0; offset <= 6; offset++) {
      const targetTime = new Date(now.getTime() - (offset * 60 * 60 * 1000));
      const targetDateStr = toDateStringFast(targetTime);
      const h = targetTime.getHours();

      const intervalEnd = new Date(targetTime.getFullYear(), targetTime.getMonth(), targetTime.getDate(), h + 1, 0, 0, 0);
      const graceThreshold = new Date(intervalEnd.getTime() + (graceMinutes * 60000));

      if (now.getTime() > graceThreshold.getTime()) {
        const intervalHourObj = new Date(targetTime);
        intervalHourObj.setHours(h, 0, 0, 0);
        const intervalHourStr = Utilities.formatDate(intervalHourObj, Session.getScriptTimeZone(), "h:00 a");

        const logData = escalationLogSheet.getDataRange().getValues();
        let alreadyEscalated = false;
        for (let i = 1; i < logData.length; i++) {
          const rowDate = logData[i][0];
          const formattedRowDate = (rowDate instanceof Date) ? toDateStringFast(rowDate) : rowDate;
          if (formattedRowDate == targetDateStr && logData[i][1] == intervalHourStr) {
            alreadyEscalated = true;
            break;
          }
        }
        if (alreadyEscalated) continue;

        const checkIn = getCheckInStatus(targetDateStr, intervalHourStr);
        if (checkIn) continue;

        const agents = getIntervalData(targetDateStr, intervalHourStr);
        if (!agents || agents.length === 0) continue;

        const statusOverrides = getIntervalStatusOverrides(targetDateStr, intervalHourStr);
        let unresolvedAgents = [];

        agents.forEach(agent => {
          const hasOverride = (agent.ldap in statusOverrides) && statusOverrides[agent.ldap] !== "";
          const hasComputed = agent.computedStatus && agent.computedStatus !== "";
          const hasCases = agent.casesLogged > 0;

          if (!hasOverride && !hasComputed && !hasCases) {
            unresolvedAgents.push(agent.ldap);
          }
        });

        if (unresolvedAgents.length > 0) {
          let scheduledPOC = "Unknown";
          const pocResult = getPOCSchedule();
          if (pocResult && pocResult.success && pocResult.schedule) {
            const match = pocResult.schedule.find(s => s.time === intervalHourStr);
            if (match) scheduledPOC = match.poc;
          } else if (Array.isArray(pocResult)) {
            const match = pocResult.find(s => s.time === intervalHourStr);
            if (match) scheduledPOC = match.poc;
          }

          const escalatedAt = new Date();
          escalationLogSheet.appendRow([targetDateStr, intervalHourStr, escalatedAt, scheduledPOC, unresolvedAgents.length, agents.length]);

          sendEscalationEmail(targetDateStr, intervalHourStr, scheduledPOC, escalatedAt, unresolvedAgents, agents.length);
        }
      }
    }
  } catch (e) {
    logError('escalationSweep', e.toString(), 'SYSTEM');
  } finally {
    try { LockService.getScriptLock().releaseLock(); } catch(e){}
  }
}

function sendEscalationEmail(dateStr, intervalHourStr, scheduledPOC, escalatedAt, unresolvedAgents, totalAgents) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const rosterSheet = ss.getSheetByName('Roster');
  let toEmails = [];
  if (rosterSheet && rosterSheet.getLastRow() > 1) {
    const rosterData = rosterSheet.getRange(2, 2, rosterSheet.getLastRow() - 1, 1).getValues();
    rosterData.forEach(row => {
      const ldap = String(row[0]).trim();
      if (ldap) toEmails.push(ldap + '@google.com');
    });
  }

  if (toEmails.length === 0) {
     logError('sendEscalationEmail', 'No recipients configured in Roster tab.', 'SYSTEM');
     return;
  }

  let ccEmails = [];
  if (scheduledPOC !== 'Unknown' && scheduledPOC !== 'Unassigned') {
    ccEmails.push(scheduledPOC.toLowerCase() + '@google.com');
  }

  const timestampStr = Utilities.formatDate(escalatedAt, Session.getScriptTimeZone(), "MMM d, yyyy 'at' h:mm a");

  const maxVisible = 10;
  const visible = unresolvedAgents.slice(0, maxVisible);
  const hiddenCount = unresolvedAgents.length - maxVisible;

  let chipsHtml = visible.map(ldap =>
    '<span style="display:inline-block; background-color:#f8f9fa; border:1px solid #dadce0; color:#3c4043; border-radius:16px; padding:4px 12px; margin:0 6px 6px 0; font-size:13px; font-weight:500;">' + ldap + '</span>'
  ).join('');

  if (hiddenCount > 0) {
    chipsHtml += '<span style="display:inline-block; background-color:#e8eaed; color:#5f6368; border-radius:16px; padding:4px 12px; margin:0 6px 6px 0; font-size:13px; font-weight:500;">+' + hiddenCount + ' more</span>';
  }

  const appUrl = ScriptApp.getService().getUrl() || PropertiesService.getScriptProperties().getProperty('WEB_APP_URL') || '';

  const body = `
  <!DOCTYPE html>
  <html>
  <head>
  <style>
    body { font-family: 'Google Sans', Roboto, Arial, sans-serif; margin: 0; padding: 0; background-color: #f8f9fa; }
  </style>
  </head>
  <body style="font-family: 'Google Sans', Roboto, Arial, sans-serif; background-color: #f8f9fa; padding: 24px;">
    <table width="100%" cellpadding="0" cellspacing="0" border="0" style="max-width: 600px; margin: 0 auto; background-color: #ffffff; border-radius: 8px; border: 1px solid #dadce0; overflow: hidden;">

      <!-- Header -->
      <tr>
        <td style="background: linear-gradient(90deg, #fbbc04 0%, #ea4335 100%); padding: 24px 32px;">
          <div style="color: #ffffff; font-size: 11px; font-weight: 700; letter-spacing: 1px; text-transform: uppercase; margin-bottom: 8px;">Case Tracker · Automated Escalation</div>
          <div style="color: #ffffff; font-size: 24px; font-weight: 400; margin: 0;">Interval Not Checked In</div>
        </td>
      </tr>

      <!-- Warning Banner -->
      <tr>
        <td style="padding: 24px 32px 0 32px;">
          <table width="100%" cellpadding="0" cellspacing="0" border="0" style="background-color: #fff8e1; border: 1px solid #fbbc04; border-radius: 8px;">
            <tr>
              <td width="40" style="padding: 16px 0 16px 16px; font-size: 20px;">⚠️</td>
              <td style="padding: 16px; color: #b06000; font-size: 14px; font-weight: 500;">This interval was not checked in within the designated grace period.</td>
            </tr>
          </table>
        </td>
      </tr>

      <!-- Content -->
      <tr>
        <td style="padding: 32px;">
          <h2 style="margin: 0 0 4px 0; font-size: 20px; color: #202124; font-weight: 400;">${intervalHourStr} Interval — ${dateStr}</h2>

          <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-top: 24px; margin-bottom: 32px;">
            <tr>
              <td width="50%" valign="top">
                <div style="font-size: 11px; text-transform: uppercase; color: #5f6368; font-weight: 600; letter-spacing: 0.5px; margin-bottom: 4px;">Scheduled POC</div>
                <div style="font-size: 14px; color: #202124;">${scheduledPOC}</div>
              </td>
              <td width="50%" valign="top">
                <div style="font-size: 11px; text-transform: uppercase; color: #5f6368; font-weight: 600; letter-spacing: 0.5px; margin-bottom: 4px;">Escalated At</div>
                <div style="font-size: 14px; color: #202124;">${timestampStr}</div>
              </td>
            </tr>
          </table>

          <table width="100%" cellpadding="0" cellspacing="0" border="0" style="margin-bottom: 32px;">
            <tr>
              <td width="48%" valign="top" style="border: 1px solid #dadce0; border-radius: 8px; padding: 20px; text-align: center;">
                <div style="font-size: 32px; color: #202124; margin-bottom: 8px;">👥 ${totalAgents}</div>
                <div style="font-size: 12px; color: #5f6368; font-weight: 500; text-transform: uppercase; letter-spacing: 0.5px;">Agents Scheduled</div>
              </td>
              <td width="4%"></td>
              <td width="48%" valign="top" style="border: 1px solid #fad2cf; background-color: #fce8e6; border-radius: 8px; padding: 20px; text-align: center;">
                <div style="font-size: 32px; color: #c5221f; margin-bottom: 8px; font-weight: 500;">${unresolvedAgents.length}</div>
                <div style="font-size: 12px; color: #c5221f; font-weight: 600; text-transform: uppercase; letter-spacing: 0.5px;">No Status / Cases Logged</div>
              </td>
            </tr>
          </table>

          <div style="margin-bottom: 32px;">
            <div style="font-size: 13px; color: #202124; font-weight: 500; margin-bottom: 12px;">Unresolved LDAPs</div>
            <div>
              ${chipsHtml}
            </div>
          </div>

          ${appUrl ? `
          <table width="100%" cellpadding="0" cellspacing="0" border="0">
            <tr>
              <td align="center">
                <a href="${appUrl}" style="display: inline-block; background-color: #1a73e8; color: #ffffff; text-decoration: none; font-size: 14px; font-weight: 500; padding: 12px 32px; border-radius: 24px;">Open Case Tracker</a>
              </td>
            </tr>
          </table>` : ''}

        </td>
      </tr>

      <!-- Footer -->
      <tr>
        <td style="background-color: #f8f9fa; border-top: 1px solid #dadce0; padding: 20px 32px; text-align: center;">
          <div style="font-size: 11px; color: #5f6368; line-height: 1.5;">
            This is an automated notice generated by the Case Tracking Portal. <br>
            Note: The Scheduled POC listed above reflects the master schedule, but they may have been substituted without the schedule being updated.
          </div>
        </td>
      </tr>
    </table>
  </body>
  </html>
  `;

  MailApp.sendEmail({
    to: toEmails.join(','),
    cc: ccEmails.join(','),
    subject: "⚠️ Interval Not Checked In — " + dateStr + " " + intervalHourStr,
    htmlBody: body,
    name: "Case Tracking Portal",
    noReply: true
  });
}

// --- 10. Fetch Analytics Data ---
function getAnalyticsData(startDateStr, endDateStr) {
  try {
    requireManagerOrThrow();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rawSheet = ss.getSheetByName('Raw_Cases');
    const auditSheet = ss.getSheetByName('Audit Queue');

    const now = new Date();
    const todayStr = toDateStringFast(now);

    let start = new Date();
    start.setDate(start.getDate() - 6); // default to 7 days
    start.setHours(0, 0, 0, 0);

    let end = new Date();
    end.setHours(23, 59, 59, 999);

    if (startDateStr) {
       let sParts = startDateStr.split('-');
       if (sParts.length === 3) {
           start = new Date(sParts[0], sParts[1] - 1, sParts[2], 0, 0, 0, 0);
       }
    }

    if (endDateStr) {
       let eParts = endDateStr.split('-');
       if (eParts.length === 3) {
           end = new Date(eParts[0], eParts[1] - 1, eParts[2], 23, 59, 59, 999);
       }
    }

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
      leaderboard: {},
      otLeaderboard: {}, // { agent: validOTCases }
      otTimelineMap: {}  // { "Date|Hour": validOTCases }
    };

    // 1. Pending Audits KPI
    if (auditSheet && auditSheet.getLastRow() > 1) {
      const auditData = auditSheet.getRange(2, 1, auditSheet.getLastRow() - 1, 1).getValues();
      data.kpis.pendingAudits = auditData.filter(r => String(r[0]).includes('PENDING')).length;
    }

    if (!rawSheet || rawSheet.getLastRow() <= 1) return data;

    // Fetch up to Column 17 (OT Type is Col 16, Index 16. Valid is Col 11. Shift Type is Col 8)
    const rawData = rawSheet.getRange(2, 1, rawSheet.getLastRow() - 1, 17).getValues();

    // Format helpers
    const currentHourStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "h:00 a");

    let activeAgentsToday = new Set();
    let activeAgentsThisHour = new Set();

    // Heatmap structure prep
    const heatmapDataMap = {}; // { '9/24/2026': { '9:00 AM': 10, ... } }

    rawData.forEach(r => {
      const rowDateObj = (r[RAW_COLS.DATE] instanceof Date) ? r[RAW_COLS.DATE] : new Date(r[RAW_COLS.DATE]);

      // Only process data within the date range
      if (rowDateObj < start || rowDateObj > end) return;

      const rowDateStr = toDateStringFast(rowDateObj);
      const rowInterval = (r[RAW_COLS.INTERVAL] instanceof Date) ? Utilities.formatDate(r[RAW_COLS.INTERVAL], Session.getScriptTimeZone(), "h:00 a") : r[RAW_COLS.INTERVAL];
      const agent = r[RAW_COLS.AGENT] ? r[RAW_COLS.AGENT].toString().trim().toLowerCase() : 'unknown';
      const site = r[RAW_COLS.SITE] || 'Unknown';
      const workflow = r[RAW_COLS.CASE_TYPE] || 'Unknown';
      const isOT = r[RAW_COLS.SHIFT_TYPE] === 'Overtime';
      const validCases = Number(r[RAW_COLS.VALID]) || 0;

      // KPI: metrics for the selected range (out-of-range rows already returned above)
      if (validCases > 0) {
        data.kpis.totalValidToday += validCases;
        activeAgentsToday.add(agent);
        if (rowDateStr === todayStr && rowInterval === currentHourStr) {
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

      // Populate Leaderboards and OT
      if (validCases > 0) {
        if (!data.leaderboard[agent]) data.leaderboard[agent] = 0;
        data.leaderboard[agent] += validCases;

        if (isOT) {
          if (!data.otLeaderboard[agent]) data.otLeaderboard[agent] = 0;
          data.otLeaderboard[agent] += validCases;

          const otTimeKey = rowDateStr + "|" + rowInterval;
          if (!data.otTimelineMap[otTimeKey]) data.otTimelineMap[otTimeKey] = 0;
          data.otTimelineMap[otTimeKey] += validCases;
        }
      }
    });

    data.kpis.activeStaffToday = activeAgentsToday.size;
    data.kpis.activeStaffThisHour = activeAgentsThisHour.size;

    // Transform Heatmap data for ApexCharts
    // Sort dates ascending
    const sortedDates = Object.keys(heatmapDataMap).sort((a, b) => new Date(a) - new Date(b));
    const HOUR_LABELS = ["12:00 AM","1:00 AM","2:00 AM","3:00 AM","4:00 AM","5:00 AM","6:00 AM","7:00 AM","8:00 AM","9:00 AM","10:00 AM","11:00 AM","12:00 PM","1:00 PM","2:00 PM","3:00 PM","4:00 PM","5:00 PM","6:00 PM","7:00 PM","8:00 PM","9:00 PM","10:00 PM","11:00 PM"];
    // Only render the active hour window so the chart isn't 60% dead space
    let activeIdx = [];
    Object.keys(heatmapDataMap).forEach(function(d) {
      HOUR_LABELS.forEach(function(h, i) { if (heatmapDataMap[d][h]) activeIdx.push(i); });
    });
    const minH = activeIdx.length ? Math.min.apply(null, activeIdx) : 0;
    const maxH = activeIdx.length ? Math.max.apply(null, activeIdx) : 23;
    const allHours = HOUR_LABELS.slice(minH, maxH + 1);
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


// --- INTERVAL STATUS & CHECK-IN LOGIC ---

function getIntervalStatusOverrides(dateStr, intervalHourStr) {
  try {
    requireManagerOrThrow();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Interval_Status');
    if (!sheet) return {};

    const data = sheet.getDataRange().getValues();
    const overrides = {};

    for (let i = 1; i < data.length; i++) {
      const rowDate = data[i][0];
      const formattedDate = (rowDate instanceof Date) ? toDateStringFast(rowDate) : rowDate;
      const rowInterval = data[i][1];
      const formattedInterval = (rowInterval instanceof Date) ? Utilities.formatDate(rowInterval, Session.getScriptTimeZone(), "h:mm a") : String(rowInterval);

      if (formattedDate === dateStr && formattedInterval === intervalHourStr) {
        const ldap = String(data[i][2]).trim().toLowerCase();
        overrides[ldap] = data[i][3];
      }
    }
    return overrides;
  } catch (e) {
    const user = Session.getActiveUser().getEmail() || 'Unknown';
    logError('getIntervalStatusOverrides', e.toString(), user);
    return {};
  }
}

function setIntervalStatus(dateStr, intervalHourStr, ldap, status) {
  const profile = requireManagerOrThrow();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Interval_Status');
    if (!sheet) {
      initializeDatabase();
      sheet = ss.getSheetByName('Interval_Status');
    }

    const data = sheet.getDataRange().getValues();
    let foundRow = -1;
    const targetLdap = String(ldap).trim().toLowerCase();

    for (let i = 1; i < data.length; i++) {
      const rowDate = data[i][0];
      const formattedDate = (rowDate instanceof Date) ? toDateStringFast(rowDate) : rowDate;
      const rowInterval = data[i][1];
      const formattedInterval = (rowInterval instanceof Date) ? Utilities.formatDate(rowInterval, Session.getScriptTimeZone(), "h:mm a") : String(rowInterval);
      const rowLdap = String(data[i][2]).trim().toLowerCase();

      if (formattedDate === dateStr && formattedInterval === intervalHourStr && rowLdap === targetLdap) {
        foundRow = i + 1;
        break;
      }
    }

    const now = new Date();
    if (foundRow !== -1) {
      sheet.getRange(foundRow, 4).setValue(status);
      sheet.getRange(foundRow, 5).setValue(profile.ldap);
      sheet.getRange(foundRow, 6).setValue(now);
    } else {
      sheet.appendRow([dateStr, intervalHourStr, targetLdap, status, profile.ldap, now]);
    }
    return { success: true };
  } catch (e) {
    const user = Session.getActiveUser().getEmail() || 'Unknown';
    logError('setIntervalStatus', e.toString(), user);
    throw new Error("Failed to save status override.");
  } finally {
    lock.releaseLock();
  }
}

function getCheckInStatus(dateStr, intervalHourStr) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Interval_CheckIns');
    if (!sheet) return null;

    const data = sheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const rowDate = data[i][0];
      const formattedDate = (rowDate instanceof Date) ? toDateStringFast(rowDate) : rowDate;
      const rowInterval = data[i][1];
      const formattedInterval = (rowInterval instanceof Date) ? Utilities.formatDate(rowInterval, Session.getScriptTimeZone(), "h:mm a") : String(rowInterval);

      if (formattedDate === dateStr && formattedInterval === intervalHourStr) {
        return {
          date: formattedDate,
          interval: formattedInterval,
          scheduledPOC: data[i][2],
          checkedInBy: data[i][3],
          timestamp: (data[i][4] instanceof Date) ? data[i][4].getTime() : data[i][4],
          result: data[i][5],
          unresolvedLDAPs: data[i][6],
          notes: data[i][7]
        };
      }
    }
    return null;
  } catch (e) {
    const user = Session.getActiveUser().getEmail() || 'Unknown';
    logError('getCheckInStatus', e.toString(), user);
    return null;
  }
}

function checkInInterval(dateStr, intervalHourStr, overrideNote) {
  const profile = requireManagerOrThrow();
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = ss.getSheetByName('Interval_CheckIns');
    if (!sheet) {
      initializeDatabase();
      sheet = ss.getSheetByName('Interval_CheckIns');
    }

    const existingCheckIn = getCheckInStatus(dateStr, intervalHourStr);
    if (existingCheckIn) {
      return { alreadyCheckedIn: true, record: existingCheckIn };
    }

    const agents = getIntervalData(dateStr, intervalHourStr);
    if (!agents || agents.length === 0) {
      return { error: "No agents scheduled for this interval — check-in not applicable." };
    }

    const unresolved = [];
    const statusOverrides = getIntervalStatusOverrides(dateStr, intervalHourStr);

    agents.forEach(agent => {
      const hasOverride = (agent.ldap in statusOverrides) && statusOverrides[agent.ldap] !== "";
      const hasComputed = agent.computedStatus && agent.computedStatus !== "";
      const hasCases = agent.casesLogged > 0;

      if (!hasOverride && !hasComputed && !hasCases) {
        unresolved.push(agent.ldap);
      }
    });

    if (unresolved.length > 0 && (!overrideNote || overrideNote.trim() === "")) {
      return { needsNote: true, unresolvedLDAPs: unresolved };
    }

    let result = (unresolved.length > 0) ? "Complete w/ Exceptions" : "Complete";
    const unresolvedStr = unresolved.join(', ');
    const noteStr = overrideNote || "";

    let scheduledPOC = "Unknown";
    const pocResult = getPOCSchedule();
    if (pocResult && pocResult.success && pocResult.schedule) {
      const match = pocResult.schedule.find(s => s.time === intervalHourStr);
      if (match) scheduledPOC = match.poc;
    } else if (Array.isArray(pocResult)) {
      const match = pocResult.find(s => s.time === intervalHourStr);
      if (match) scheduledPOC = match.poc;
    }

    const now = new Date();
    sheet.appendRow([dateStr, intervalHourStr, scheduledPOC, profile.ldap, now, result, unresolvedStr, noteStr]);

    return {
      success: true,
      record: {
        date: dateStr,
        interval: intervalHourStr,
        scheduledPOC: scheduledPOC,
        checkedInBy: profile.ldap,
        timestamp: now.getTime(),
        result: result,
        unresolvedLDAPs: unresolvedStr,
        notes: noteStr
      }
    };
  } catch (e) {
    const user = Session.getActiveUser().getEmail() || 'Unknown';
    logError('checkInInterval', e.toString(), user);
    return { error: "System encountered an error during check-in." };
  } finally {
    lock.releaseLock();
  }
}

// 9b. Fetch Live POC Schedule
function getPOCSchedule() {
  try {
    const pocSheetUrl = 'https://docs.google.com/spreadsheets/d/1SwO6Wet3OWPQDkXC2jyQ3rbPTjCDMZbAc5fLbDK6HHU/edit';
    const ss = SpreadsheetApp.openByUrl(pocSheetUrl);
    const sheet = ss.getSheetByName('POC Schedule');

    if (!sheet) {
      return { success: false, error: 'POC Schedule tab not found in the source sheet.' };
    }

    const now = new Date();
    const todayStr = toDateStringFast(now); // "M/d/yyyy"

    // Header row is Row 8
    const headers = sheet.getRange(8, 1, 1, sheet.getLastColumn()).getValues()[0];

    // Find today's column (Starts from column A / Index 0)
    let targetColIdx = -1;
    for (let i = 0; i < headers.length; i++) {
      let cellDate = headers[i];
      let formattedCellDate = "";
      try {
        if (cellDate instanceof Date) {
          formattedCellDate = toDateStringFast(cellDate);
        } else if (cellDate) {
          formattedCellDate = toDateStringFast(new Date(cellDate));
        }
      } catch(e) {}

      if (formattedCellDate === todayStr) {
        targetColIdx = i;
        break;
      }
    }

    if (targetColIdx === -1) {
       return { success: false, error: 'Could not find a column for today (' + todayStr + ') in the POC sheet.' };
    }

    // Times are in Col A (Index 0), starting from Row 9 (Index 8 in array if we read from top)
    // We will fetch from Row 9 to Row 32 (24 hours)
    const scheduleData = sheet.getRange(9, 1, 24, sheet.getLastColumn()).getValues();

    // Fetch Site mapping from Masterlist
    const activeSs = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = activeSs.getSheetByName('Masterlist');
    const siteMap = {};
    if (masterSheet && masterSheet.getLastRow() > 1) {
      const mData = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 45).getValues();
      mData.forEach(r => {
        if (r[0]) siteMap[String(r[0]).trim().toLowerCase()] = String(r[44] || '').trim();
      });
    }

    let pocList = [];
    for (let r = 0; r < scheduleData.length; r++) {
       let timeVal = scheduleData[r][0];
       let pocVal = scheduleData[r][targetColIdx];
       let timeStr = (timeVal instanceof Date) ? Utilities.formatDate(timeVal, Session.getScriptTimeZone(), "h:00 a") : String(timeVal);

       let pocStr = pocVal ? String(pocVal).trim() : 'Unassigned';
       let siteStr = siteMap[pocStr.toLowerCase()] || '';

       pocList.push({ time: timeStr, poc: pocStr, site: siteStr });
    }

    return { success: true, schedule: pocList };

  } catch (e) {
    const user = Session.getActiveUser().getEmail() || 'Unknown';
    logError('getPOCSchedule', e.toString(), user);
    return { success: false, error: 'Failed to fetch POC Schedule.' };
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
    
    // 1. Build a whitelist of "Email" Channel Agents & cache Demographics from the Masterlist
    const cache = CacheService.getScriptCache();
    let emailAgentsList = cache.get('emailAgentsMap_v2');
    let emailAgents = {};
    
    if (emailAgentsList) {
      emailAgents = JSON.parse(emailAgentsList);
    } else if (masterSheet && masterSheet.getLastRow() > 1) {
      const masterData = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 45).getValues();
      for (let i = 0; i < masterData.length; i++) {
        const ldap = masterData[i][0] ? masterData[i][0].toString().toLowerCase() : '';
        if (!ldap) continue;
        
        const channel = masterData[i][23] ? masterData[i][23].toString().toLowerCase() : '';
        const supervisor = masterData[i][27] ? masterData[i][27].toString() : 'Unknown';
        if(ldap) emailAgents[ldap] = { isEmail: channel === 'email', supervisor: supervisor };
      }
      cache.put('emailAgentsMap_v2', JSON.stringify(emailAgents), 14400);
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
          if (!emailAgents[ldap] || !emailAgents[ldap].isEmail) {
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
                  supervisor: emailAgents[ldap] ? emailAgents[ldap].supervisor : 'Unknown',
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
              supervisor: emailAgents[ldap] ? emailAgents[ldap].supervisor : 'Unknown',
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
// 11. Fetch "My Profile" data — identity + personal stats, always scoped to the caller's own LDAP
function getMyProfileData() {
  try {
    const profile = getUserProfile();
    const ldap = profile.ldap;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = ss.getSheetByName('Masterlist');

    let details = {
      position: '', gradeLevel: '', employeeStatus: '', team: '',
      hireDateStr: '', tenureText: '',
      reportsTo: []
    };

    if (masterSheet && masterSheet.getLastRow() > 1) {
      // Through Col AC (Immediate Superior LDAP) — 29 columns
      const masterData = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 29).getValues();
      const byLdap = {};
      masterData.forEach(function(r) {
        const rowLdap = r[0] ? r[0].toString().trim().toLowerCase() : '';
        if (rowLdap) byLdap[rowLdap] = r;
      });

      const myRow = byLdap[ldap.toLowerCase()];
      if (myRow) {
        details.position = myRow[15] || '';        // Col P
        details.gradeLevel = myRow[16] || '';       // Col Q
        details.employeeStatus = myRow[18] || '';   // Col S
        details.team = myRow[26] || '';              // Col AA

        const hireDateVal = myRow[6]; // Col G — CNX Hire Date
        if (hireDateVal) {
          const hireDate = (hireDateVal instanceof Date) ? hireDateVal : new Date(hireDateVal);
          if (!isNaN(hireDate.getTime())) {
            details.hireDateStr = Utilities.formatDate(hireDate, Session.getScriptTimeZone(), 'MMM d, yyyy');
            details.tenureText = computeTenureText(hireDate);
          }
        }

        // Level 1: my Immediate Superior
        const sup1Name = myRow[27] || '';           // Col AB
        const sup1Ldap = myRow[28] ? myRow[28].toString().trim().toLowerCase() : ''; // Col AC
        if (sup1Ldap && sup1Ldap !== '-') {
          details.reportsTo.push({ name: sup1Name || sup1Ldap, ldap: sup1Ldap });

          // Level 2: that superior's own Immediate Superior
          const sup1Row = byLdap[sup1Ldap];
          if (sup1Row) {
            const sup2Name = sup1Row[27] || '';
            const sup2Ldap = sup1Row[28] ? sup1Row[28].toString().trim().toLowerCase() : '';
            if (sup2Ldap && sup2Ldap !== '-' && sup2Ldap !== sup1Ldap) {
              details.reportsTo.push({ name: sup2Name || sup2Ldap, ldap: sup2Ldap });
            }
          }
        }
      }
    }

    // Personal stats — scoped strictly to this caller's own LDAP
    const rawSheet = ss.getSheetByName('Raw_Cases');
    const auditSheet = ss.getSheetByName('Audit Queue');

    const stats = { today: 0, thisWeek: 0, thisMonth: 0, pendingAudits: 0 };
    const breakdown = { Regular: 0, Reopened: 0, Manual: 0 };
    const history = {}; // DateStr -> { All, Regular, Reopened, Manual, Telus, Cimba }

    const now = new Date();
    const todayStr = toDateStringFast(now);
    const weekStart = getWeekStartMonday(now);
    const monthStart = new Date(now.getFullYear(), now.getMonth(), 1);

    const trendMap = {};
    const trendMeta = [];
    for (let i = 6; i >= 0; i--) {
      const d = new Date(now);
      d.setDate(d.getDate() - i);
      const key = toDateStringFast(d);
      trendMap[key] = 0;
      trendMeta.push({ key: key, label: Utilities.formatDate(d, Session.getScriptTimeZone(), 'EEE') });
    }

    if (rawSheet && rawSheet.getLastRow() > 1) {
      const rawData = rawSheet.getRange(2, 1, rawSheet.getLastRow() - 1, 17).getValues();
      rawData.forEach(function(r) {
        const rowLdap = r[RAW_COLS.AGENT] ? r[RAW_COLS.AGENT].toString().trim().toLowerCase() : '';
        if (rowLdap !== ldap.toLowerCase()) return;

        const rowDateObj = (r[RAW_COLS.DATE] instanceof Date) ? r[RAW_COLS.DATE] : new Date(r[RAW_COLS.DATE]);
        const rowDateStr = toDateStringFast(rowDateObj);
        const validCases = Number(r[RAW_COLS.VALID]) || 0;
        if (validCases <= 0) return;

        if (rowDateStr === todayStr) stats.today += validCases;
        if (rowDateObj >= weekStart) stats.thisWeek += validCases;
        if (rowDateObj >= monthStart) stats.thisMonth += validCases;
        if (trendMap.hasOwnProperty(rowDateStr)) trendMap[rowDateStr] += validCases;

        const caseType = r[RAW_COLS.CASE_TYPE] || '';
        let typeCategory = 'Regular';
        if (caseType === 'Manual Assignment') typeCategory = 'Manual';
        else if (caseType === 'Reopened Cases') typeCategory = 'Reopened';
        else if (caseType === 'Telus') typeCategory = 'Telus';
        else if (caseType === 'Cimba') typeCategory = 'Cimba';

        if (rowDateObj >= monthStart) {
          if (typeCategory === 'Manual') breakdown.Manual += validCases;
          else if (typeCategory === 'Reopened') breakdown.Reopened += validCases;
          else breakdown.Regular += validCases;
        }

        const historyKey = toISODateStringFast(rowDateObj);
        if (!history[historyKey]) {
          history[historyKey] = { All: 0, Regular: 0, Reopened: 0, Manual: 0, Telus: 0, Cimba: 0 };
        }
        history[historyKey].All += validCases;
        if (history[historyKey][typeCategory] !== undefined) {
          history[historyKey][typeCategory] += validCases;
        } else {
          history[historyKey].Regular += validCases; // Fallback
        }
      });
    }

    if (auditSheet && auditSheet.getLastRow() > 1) {
      const auditData = auditSheet.getRange(2, 1, auditSheet.getLastRow() - 1, 3).getValues();
      auditData.forEach(function(r) {
        const rowLdap = r[2] ? r[2].toString().trim().toLowerCase() : '';
        if (rowLdap === ldap.toLowerCase() && String(r[0]).includes('PENDING')) stats.pendingAudits++;
      });
    }

    return {
      ldap: ldap,
      name: profile.name,
      site: profile.site,
      workflow: profile.workflow,
      position: details.position,
      gradeLevel: details.gradeLevel,
      employeeStatus: details.employeeStatus,
      team: details.team,
      hireDateStr: details.hireDateStr,
      tenureText: details.tenureText,
      reportsTo: details.reportsTo,
      stats: stats,
      trend: trendMeta.map(function(t) { return { label: t.label, count: trendMap[t.key] }; }),
      breakdown: breakdown,
      history: history
    };
  } catch (e) {
    const user = Session.getActiveUser().getEmail() || 'Unknown';
    logError('getMyProfileData', e.toString(), user);
    throw new Error("Unable to load profile. Please try again.");
  }
}

function computeTenureText(hireDate) {
  const now = new Date();
  let years = now.getFullYear() - hireDate.getFullYear();
  let months = now.getMonth() - hireDate.getMonth();
  if (now.getDate() < hireDate.getDate()) months--;
  if (months < 0) { years--; months += 12; }
  if (years <= 0 && months <= 0) return 'Less than a month';
  const yStr = years > 0 ? years + (years === 1 ? ' yr' : ' yrs') : '';
  const mStr = months > 0 ? months + (months === 1 ? ' mo' : ' mos') : '';
  return [yStr, mStr].filter(Boolean).join(' ');
}

function getWeekStartMonday(date) {
  const d = new Date(date);
  const day = d.getDay();
  const diff = (day === 0 ? -6 : 1 - day);
  d.setDate(d.getDate() + diff);
  d.setHours(0, 0, 0, 0);
  return d;
}
