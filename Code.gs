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
    console.error("Failed to write to Error_Logs tab: " + (e && e.message ? e.message : e));
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

// Normalizes any interval-hour representation — a Date object, "16:00" (24-hour
// text), or "4:00 PM" (12-hour text) — to a 0-23 integer. Comparisons should
// always go through this instead of comparing formatted strings directly,
// since different call sites (UI vs. escalationSweep vs. Sheets auto-typing)
// don't agree on which text format they're using.
function parseHourToInt(val) {
  if (val instanceof Date) return val.getHours();
  const str = String(val).trim();

  let m = str.match(/^(\d{1,2}):\d{2}\s*(AM|PM)$/i);
  if (m) {
    let h = parseInt(m[1], 10) % 12;
    if (m[2].toUpperCase() === 'PM') h += 12;
    return h;
  }

  m = str.match(/^(\d{1,2}):\d{2}$/);
  if (m) return parseInt(m[1], 10);

  return null;
}

function startOfDay_(d) {
  return new Date(d.getFullYear(), d.getMonth(), d.getDate());
}

// Sheets here are append-only (chronological). Reads from the bottom up in chunks and stops
// once it passes sinceDate, instead of loading every row.
// Returns { startRow, values }; startRow is the sheet row number of values[0].
function readSheetTail_(sheet, numCols, dateColIdx, sinceDate) {
  const last = sheet.getLastRow();
  if (last <= 1) return { startRow: 2, values: [] };
  const CHUNK = 3000;
  const chunks = [];
  let endRow = last;
  let startRow = last + 1;
  while (endRow >= 2) {
    const chunkStart = Math.max(2, endRow - CHUNK + 1);
    const vals = sheet.getRange(chunkStart, 1, endRow - chunkStart + 1, numCols).getValues();
    chunks.unshift(vals);
    startRow = chunkStart;
    const v = vals[0][dateColIdx];
    const d = (v instanceof Date) ? v : new Date(v);
    if (!isNaN(d.getTime()) && d < sinceDate) break;
    endRow = chunkStart - 1;
  }
  return { startRow: startRow, values: [].concat.apply([], chunks) };
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
  var template = HtmlService.createTemplateFromFile('Index');
  
  // Fetching here forces the Apps Script scanner to detect the required permission
  try {
    template.tailwindCss = getTailwindJs_();
  } catch (e) {
    logError('doGet', 'Tailwind unavailable: ' + e, 'SYSTEM');
    template.tailwindCss = '';
  }
  
  return template.evaluate()
      .setTitle('Play Case Tracker')
      .setFaviconUrl('https://www.gstatic.com/images/branding/product/2x/play_prism_64dp.png')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// Third-party JS is fetched server-side (some networks block the CDNs in the browser), gzipped into
// Script Properties, and served from there. URLs are only hit when no stored copy exists.
const CDN_ASSETS_ = {
  tailwind: { url: 'https://cdn.tailwindcss.com/3.4.17', minLen: 100000 },
  apex:     { url: 'https://cdn.jsdelivr.net/npm/apexcharts', minLen: 100000 }
};

function getCdnJs_(name) {
  try {
    const all = PropertiesService.getScriptProperties().getProperties();
    const n = parseInt(all['CDN_' + name + '_N'] || '0', 10);
    if (n > 0) {
      let b64 = '';
      for (let i = 0; i < n; i++) b64 += all['CDN_' + name + '_' + i] || '';
      return Utilities.ungzip(Utilities.newBlob(Utilities.base64Decode(b64), 'application/x-gzip')).getDataAsString();
    }
  } catch (e) {
    logError('getCdnJs_', 'Stored ' + name + ' copy unreadable: ' + e, 'SYSTEM');
  }
  return refreshCdnJs_(name);
}

function refreshCdnJs_(name) {
  const asset = CDN_ASSETS_[name];
  const res = UrlFetchApp.fetch(asset.url, { muteHttpExceptions: true, followRedirects: true });
  const js = res.getContentText();
  if (res.getResponseCode() !== 200 || js.length < asset.minLen) {
    throw new Error(name + ' fetch failed: HTTP ' + res.getResponseCode());
  }
  try {
    const b64 = Utilities.base64Encode(Utilities.gzip(Utilities.newBlob(js, 'text/plain', name + '.js')).getBytes());
    const SIZE = 8000; // Script Properties cap is 9 KB per value
    const n = Math.ceil(b64.length / SIZE);
    const obj = {};
    obj['CDN_' + name + '_N'] = String(n);
    for (let i = 0; i < n; i++) obj['CDN_' + name + '_' + i] = b64.substr(i * SIZE, SIZE);
    PropertiesService.getScriptProperties().setProperties(obj);
  } catch (e) {
    logError('refreshCdnJs_', 'Could not store ' + name + ' copy: ' + e, 'SYSTEM');
  }
  return js;
}

function getTailwindJs_() { return getCdnJs_('tailwind'); }

// Called by the client the first time a manager opens Analytics.
function getApexJs() {
  requireManagerOrThrow();
  return getCdnJs_('apex');
}

// Run once from the editor (as a Roster account) after deploying. Re-run to refresh the stored copies.
function warmCdnCache() {
  requireManagerOrThrow();
  Object.keys(CDN_ASSETS_).forEach(function(k) { refreshCdnJs_(k); });
}

// 2. Auto-Initialize the Database Structure
function initializeDatabase() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    
    const sheetsConfig = {
      'Raw_Cases': ['Timestamp', 'Date', 'Interval', 'Agent', 'Name', 'Site', 'LOB', 'Workflow', 'Shift Type', 'Case Type', 'Total', 'Valid', 'Flagged', 'Case IDs', 'Audit Notes', 'Interval Activity', 'OT Type'],
      'Index_CaseIDs': ['Case ID', 'Type Logged', 'Date Logged', 'Agent', 'Timestamp', 'Interval', 'Flagged'],
      'Audit Queue': ['Status', 'Timestamp', 'Agent', 'Site', 'Case Type', 'Total Logged', 'Flagged IDs', 'Audit Reason', 'Resolution', 'RawRowRef'],
      'Error_Logs': ['Timestamp', 'Function', 'User', 'Error Message'],
      'Interval_Status': ['Date', 'Interval', 'LDAP', 'Status', 'SetBy', 'Timestamp'],
      'Interval_CheckIns': ['Date', 'Interval', 'ScheduledPOC', 'CheckedInBy', 'Timestamp', 'Result', 'UnresolvedLDAPs', 'Notes', 'IsSubstitute'],
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
    logError('initializeDatabase', 'Initialization error: ' + e, 'SYSTEM');
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
  const nowStr = toDateStringFast(new Date());
  CacheService.getUserCache().remove('userProfile_' + currentLdap);
  CacheService.getScriptCache().removeAll(['allAgentsList', 'emailAgentsList', 'agentDemographicsList', 'emailAgentsMap_v2', 'poc_schedule_' + nowStr]);
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

    const ALLOWED_SHIFT = ['Regular Shift', 'Overtime'];
    const ALLOWED_OT = ['', 'Pre-Shift OT', 'Post-Shift OT', 'RD OT'];
    const ALLOWED_ACTIVITY = ['Normal Production', 'Break/Lunch', 'Coaching/Training'];
    const ALLOWED_CASE_TYPE = ['Regular Email (Take Next)', 'Reopened Cases', 'Telus Cases', 'Manual Assignment', 'Cimba Cases', 'N/A'];
    if (ALLOWED_SHIFT.indexOf(shiftType) === -1 || ALLOWED_OT.indexOf(otType) === -1 ||
        ALLOWED_ACTIVITY.indexOf(intervalActivity) === -1 || ALLOWED_CASE_TYPE.indexOf(caseType) === -1 ||
        (intervalActivity === 'Normal Production' && caseType === 'N/A')) {
      return { success: false, error: "Invalid submission data. Please refresh the page and try again." };
    }

    const CASE_ID_PATTERN = /^\d-\d{7,14}$/;
    let rawIds = rawText.split(/[\n,;\s]+/).map(id => id.trim()).filter(id => id !== '');
    let uniqueIds = [...new Set(rawIds)];
    let malformedIds = uniqueIds.filter(id => !CASE_ID_PATTERN.test(id));
    uniqueIds = uniqueIds.filter(id => CASE_ID_PATTERN.test(id));
    let validIds = [];
    let flaggedIds = [];
    let reasonCounts = {};
    let flaggedDetails = [];

    const idSet = new Set(uniqueIds);
    const historyById = {};
    const indexLast = indexSheet.getLastRow();
    if (indexLast > 1 && idSet.size > 0) {
      const idCol = indexSheet.getRange(2, 1, indexLast - 1, 1).getValues();
      let firstHit = -1, lastHit = -1;
      idCol.forEach((r, i) => {
        if (idSet.has(String(r[0]).trim())) { if (firstHit === -1) firstHit = i; lastHit = i; }
      });
      if (firstHit !== -1) {
        indexSheet.getRange(firstHit + 2, 1, lastHit - firstHit + 1, 4).getValues().forEach(row => {
          const k = String(row[0]).trim();
          if (idSet.has(k)) (historyById[k] = historyById[k] || []).push(row);
        });
      }
    }

    uniqueIds.forEach(id => {
      let isFlagged = false;
      let reason = "";
      let history = historyById[id] || [];
      let firstEntry = null;

      if (caseType === 'Regular Email (Take Next)' && history.length > 0) {
        isFlagged = true;
        reason = "Previously logged in system";
        firstEntry = history[0];
      } else if (caseType === 'Reopened Cases') {
        let todayHistory = history.filter(row => toDateStringFast(row[2]) === dateStr && row[1] === 'Reopened Cases' && row[3] === ldap);
        if (todayHistory.length > 0) {
          isFlagged = true;
          reason = "Already reopened today";
          firstEntry = todayHistory[0];
        }
      }

      if (isFlagged) {
        flaggedIds.push(id);
        reasonCounts[reason] = (reasonCounts[reason] || 0) + 1;
        flaggedDetails.push({
          id: id,
          reason: reason,
          firstBy: firstEntry ? String(firstEntry[3]).trim().toLowerCase() : '',
          firstOn: firstEntry ? toDateStringFast(firstEntry[2]) : ''
        });
      } else {
        validIds.push(id);
      }
    });

    let totalCount = uniqueIds.length;
    let validCount = validIds.length;
    let flaggedCount = flaggedIds.length;
    let breakLunchShort = intervalActivity === 'Break/Lunch' && validCount < 3;
    let auditNotes = flaggedCount > 0
      ? (breakLunchShort ? "🔴 Break/Lunch minimum unmet (" + validCount + "/3 valid) — " : "⚠️ ") +
        Object.entries(reasonCounts).map(([reason, count]) => `${reason} (${count})`).join(', ')
      : "Clean";

    let rawRowNumber = rawSheet.getLastRow() + 1;
    // Writes LOB, Workflow, and Name into the raw sheet, appended Interval Activity at the end
    rawSheet.appendRow([timestamp, dateStr, intervalStr, ldap, userProfile.name, site, userProfile.lob, userProfile.workflow, shiftType, caseType, totalCount, validCount, flaggedCount, uniqueIds.join(', '), auditNotes, intervalActivity, otType]);
    SpreadsheetApp.flush();
    rawRowNumber = rawSheet.getLastRow();
    rawSheet.getRange(rawRowNumber, 3).setNumberFormat('@'); // Prevent Sheets from auto-converting this to a Date

    const flaggedIdSet = new Set(flaggedIds);
    let indexDataToAppend = uniqueIds.map(id => [id, caseType, dateStr, ldap, timestamp, intervalStr, flaggedIdSet.has(id) ? 'Yes' : 'No']);
    if (indexDataToAppend.length > 0) {
      indexSheet.getRange(indexSheet.getLastRow() + 1, 1, indexDataToAppend.length, 7).setValues(indexDataToAppend);
    }

    if (flaggedCount > 0) {
      auditSheet.appendRow(["🔴 PENDING", timestamp, ldap, site, caseType, totalCount, flaggedIds.join(', '), auditNotes, "", rawRowNumber]);
    }
    
    // Real-time cross-check against the Expired/Expiring queue (never blocks a submission)
    try { markExpiringWorked_(uniqueIds, ldap, timestamp); } catch (qe) { logError('markExpiringWorked_', qe.toString(), ldap); }

    // If this agent was tagged "Absent" for any interval earlier in today's
    // shift, reclassify those specific intervals as "Late" now that they've
    // logged cases. Regular Shift only — OT has no scheduled start to be
    // "late" against.
    if (shiftType === 'Regular Shift') {
      convertAbsentToLate_(ldap, dateStr, timestamp.getHours());
    }

    return { success: true, valid: validCount, flagged: flaggedCount, rejected: malformedIds.length, flaggedDetails: flaggedDetails.slice(0, 50) };
    
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
      const rawData = readSheetTail_(rawSheet, 15, 1, startOfDay_(new Date())).values; 

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
  if (resolution !== 'Approve' && resolution !== 'Reject') {
    return { success: false, error: "Invalid resolution." };
  }
  if (!auditsToProcess || auditsToProcess.length === 0) {
    return { success: true };
  }

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const auditSheet = ss.getSheetByName('Audit Queue');
    const rawSheet = ss.getSheetByName('Raw_Cases');
    const indexSheet = ss.getSheetByName('Index_CaseIDs');
    if (!auditSheet || auditSheet.getLastRow() <= 1) return { success: true };

    // Only the audit ROW numbers come from the client; status, RawRowRef and IDs are read from the sheet.
    const isApprove = resolution === 'Approve';
    const wanted = new Set(auditsToProcess.map(a => parseInt(a.auditRow, 10)));
    const auditData = auditSheet.getRange(2, 1, auditSheet.getLastRow() - 1, 10).getValues();

    const targets = [];
    auditData.forEach((r, i) => {
      const rowNum = i + 2;
      if (!wanted.has(rowNum)) return;
      if (!String(r[0]).includes('PENDING')) return; // already resolved
      targets.push({
        auditRow: rowNum,
        ts: r[1],
        agent: String(r[2]).trim().toLowerCase(),
        caseType: r[4],
        flaggedIds: String(r[6] || '').split(',').map(s => s.trim()).filter(Boolean),
        rawRow: parseInt(r[9], 10)
      });
    });
    if (targets.length === 0) return { success: true };

    // 1. Raw_Cases: one read + two writes for the whole batch
    const rawRows = targets.map(t => t.rawRow).filter(n => n >= 2);
    if (rawSheet && rawRows.length > 0) {
      const minRaw = Math.min.apply(null, rawRows);
      const maxRaw = Math.max.apply(null, rawRows);
      const span = maxRaw - minRaw + 1;
      const block = rawSheet.getRange(minRaw, 12, span, 4).getValues(); // L Valid, M Flagged, N Case IDs (untouched), O Audit Notes
      targets.forEach(t => {
        const b = block[t.rawRow - minRaw];
        if (!b) return;
        if (isApprove) {
          b[0] = (Number(b[0]) || 0) + (Number(b[1]) || 0);
          b[3] = "✅ Resolved by Manager";
        } else {
          b[3] = "❌ Rejected by Manager (Duplicate/Fraud)";
        }
        b[1] = 0;
      });
      rawSheet.getRange(minRaw, 12, span, 2).setValues(block.map(b => [b[0], b[1]]));
      rawSheet.getRange(minRaw, 15, span, 1).setValues(block.map(b => [b[3]]));
    }

    // 2. Index_CaseIDs: an approved ID is no longer a flagged duplicate (rejected ones stay flagged)
    if (isApprove && indexSheet && indexSheet.getLastRow() > 1) {
      let earliest = null;
      targets.forEach(t => { if (t.ts instanceof Date && (!earliest || t.ts < earliest)) earliest = t.ts; });
      const tail = readSheetTail_(indexSheet, 7, 2, earliest ? startOfDay_(earliest) : new Date(2000, 0, 1));

      const wantedIds = {}; // "id|agent|type" -> [{ tsMs, day }]
      targets.forEach(t => {
        const tsMs = (t.ts instanceof Date) ? t.ts.getTime() : null;
        const day = (t.ts instanceof Date) ? toDateStringFast(t.ts) : null;
        t.flaggedIds.forEach(id => {
          const k = id + '|' + t.agent + '|' + t.caseType;
          (wantedIds[k] = wantedIds[k] || []).push({ tsMs: tsMs, day: day });
        });
      });

      const flagCol = tail.values.map(r => [r[6]]);
      let changed = false;
      tail.values.forEach((r, i) => {
        if (String(r[6]).trim().toLowerCase() !== 'yes') return;
        const cands = wantedIds[String(r[0]).trim() + '|' + String(r[3]).trim().toLowerCase() + '|' + r[1]];
        if (!cands) return;
        const rowMs = (r[4] instanceof Date) ? r[4].getTime() : null;
        const rowDay = toDateStringFast(r[2]);
        const hit = cands.some(c => (rowMs !== null && c.tsMs !== null) ? Math.abs(rowMs - c.tsMs) < 2000 : rowDay === c.day);
        if (hit) { flagCol[i][0] = 'No'; changed = true; }
      });
      if (changed) indexSheet.getRange(tail.startRow, 7, flagCol.length, 1).setValues(flagCol);
    }

    // 3. Audit Queue last, so a failure above leaves the audits pending
    const statusCol = auditData.map(r => [r[0]]);
    const resolutionCol = auditData.map(r => [r[8]]);
    targets.forEach(t => {
      statusCol[t.auditRow - 2][0] = isApprove ? "🟢 APPROVED" : "⚫ REJECTED";
      resolutionCol[t.auditRow - 2][0] = isApprove ? "Approved" : "Rejected";
    });
    auditSheet.getRange(2, 1, statusCol.length, 1).setValues(statusCol);
    auditSheet.getRange(2, 9, resolutionCol.length, 1).setValues(resolutionCol);

    return { success: true, resolved: targets.length };
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
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rawSheet = ss.getSheetByName('Raw_Cases');
    const userProfile = getUserProfile();
    
    // If targetLdap is provided and the user is a manager, use it. Otherwise default to their own ldap.
    const queryLdap = (userProfile.isManager && targetLdap) ? targetLdap.toLowerCase() : userProfile.ldap.toLowerCase();
    
    let submissions = [];

    if (rawSheet && rawSheet.getLastRow() > 1) {
      const dp = dateStr.split('/');
      const sinceDate = new Date(parseInt(dp[2], 10), parseInt(dp[0], 10) - 1, parseInt(dp[1], 10));
      const tail = readSheetTail_(rawSheet, 16, 1, sinceDate);
      const rawData = tail.values;
      const rowOffset = tail.startRow; // sheet row of rawData[0]
      
      const dedupeMap = {};

      rawData.forEach((r, ri) => {
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
            rawRow: rowOffset + ri,
            flaggedCount: r[12],          // Col M
            notes: r[14],                 // Col O
            caseIds: caseIdsVal           // Col N
          };
        }
      });

      submissions = Object.values(dedupeMap);

      // Attach review status for anything that was flagged
      const auditSheet = ss.getSheetByName('Audit Queue');
      if (auditSheet && auditSheet.getLastRow() > 1) {
        const statusByRawRow = {};
        auditSheet.getRange(2, 1, auditSheet.getLastRow() - 1, 10).getValues().forEach(a => {
          if (String(a[2]).trim().toLowerCase() !== queryLdap.trim()) return;
          const s = String(a[0]);
          statusByRawRow[a[9]] = s.includes('PENDING') ? 'Pending' : (s.includes('APPROVED') ? 'Approved' : (s.includes('REJECTED') ? 'Rejected' : ''));
        });
        submissions.forEach(s => { s.auditStatus = statusByRawRow[s.rawRow] || ''; });
      }
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


// --- CASE SUBMISSION LOG (global, read-only view for all users) ---

// Returns every Case ID logged for dateStr (M/d/yyyy), optionally narrowed to
// one interval hour, newest first. No manager gate — this is the one view
// every agent can see regardless of role, by design.
function getGlobalCaseLog(dateStr, intervalHourStr) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Index_CaseIDs');
    if (!sheet || sheet.getLastRow() <= 1) return [];

    const dp = dateStr.split('/');
    const sinceDate = new Date(parseInt(dp[2], 10), parseInt(dp[0], 10) - 1, parseInt(dp[1], 10));
    const data = readSheetTail_(sheet, 7, 2, sinceDate).values;
    const wantHour = intervalHourStr ? parseHourToInt(intervalHourStr) : null;

    const rows = [];
    data.forEach(r => {
      const rowDateStr = toDateStringFast(r[2]);
      if (rowDateStr !== dateStr) return;
      if (wantHour !== null && parseHourToInt(r[5]) !== wantHour) return;

      const ts = r[4];
      rows.push({
        caseId: r[0],
        caseType: r[1],
        agent: r[3] ? r[3].toString().trim().toLowerCase() : '',
        interval: (r[5] instanceof Date)
          ? Utilities.formatDate(r[5], Session.getScriptTimeZone(), "h:00 a")
          : (r[5] || ''),
        timestamp: (ts instanceof Date) ? Utilities.formatDate(ts, Session.getScriptTimeZone(), "h:mm a") : '',
        rawTs: (ts instanceof Date) ? ts.getTime() : 0,
        flagged: String(r[6]).trim().toLowerCase() === 'yes'
      });
    });

    rows.sort((a, b) => b.rawTs - a.rawTs);
    return rows;
  } catch (e) {
    const user = Session.getActiveUser().getEmail() || 'Unknown';
    logError('getGlobalCaseLog', e.toString(), user);
    throw new Error("Unable to load the case submission log. Please try again.");
  }
}

// Every time a Case ID was logged, across all dates. Read-only and open to all users, like the Case Log.
function lookupCaseId(caseId) {
  try {
    const id = String(caseId || '').trim();
    if (!/^\d-\d{7,14}$/.test(id)) return { success: false, error: 'Enter a Case ID like 1-12345678.' };

    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Index_CaseIDs');
    if (!sheet || sheet.getLastRow() <= 1) return { caseId: id, entries: [], total: 0 };

    const hits = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1)
      .createTextFinder(id).matchEntireCell(true).findAll();

    const tz = Session.getScriptTimeZone();
    const entries = hits.slice(0, 50).map(function(cell) {
      const r = sheet.getRange(cell.getRow(), 1, 1, 7).getValues()[0];
      const ts = r[4];
      return {
        caseType: r[1],
        date: (r[2] instanceof Date) ? Utilities.formatDate(r[2], tz, 'M/d/yyyy') : String(r[2]),
        agent: r[3] ? String(r[3]).trim().toLowerCase() : '',
        interval: (r[5] instanceof Date) ? Utilities.formatDate(r[5], tz, 'h:00 a') : (r[5] || ''),
        timestamp: (ts instanceof Date) ? Utilities.formatDate(ts, tz, 'M/d/yyyy h:mm a') : '',
        rawTs: (ts instanceof Date) ? ts.getTime() : 0,
        flagged: String(r[6]).trim().toLowerCase() === 'yes'
      };
    });
    entries.sort(function(a, b) { return a.rawTs - b.rawTs; });
    return { caseId: id, entries: entries, total: hits.length };
  } catch (e) {
    const user = Session.getActiveUser().getEmail() || 'Unknown';
    logError('lookupCaseId', e.toString(), user);
    throw new Error("Unable to look up that Case ID. Please try again.");
  }
}

// One-time backfill for rows logged before Timestamp/Interval/Flagged existed
// on Index_CaseIDs. Cross-references Raw_Cases (which does have real
// timestamps) to reconstruct them. Safe to re-run — it only fills rows that
// are still blank, and consumes matching Raw_Cases entries in sheet order so
// repeat Case IDs (including flagged duplicates) line up correctly.
function backfillCaseLogMetadata() {
  requireManagerOrThrow();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const indexSheet = ss.getSheetByName('Index_CaseIDs');
  const rawSheet = ss.getSheetByName('Raw_Cases');
  const auditSheet = ss.getSheetByName('Audit Queue');
  if (!indexSheet || !rawSheet) return { success: false, error: 'Required sheets not found.' };

  // Ensure headers exist on sheets created before this update.
  const headerRow = indexSheet.getRange(1, 1, 1, 7).getValues()[0];
  if (!headerRow[4]) indexSheet.getRange(1, 5, 1, 3).setValues([['Timestamp', 'Interval', 'Flagged']]);

  // Map RawRowRef -> Set of flagged Case IDs, from the Audit Queue.
  const flaggedByRow = {};
  if (auditSheet && auditSheet.getLastRow() > 1) {
    const auditData = auditSheet.getRange(2, 1, auditSheet.getLastRow() - 1, 10).getValues();
    auditData.forEach(r => {
      const rawRowRef = r[9];
      const ids = String(r[6] || '').split(',').map(s => s.trim()).filter(Boolean);
      flaggedByRow[rawRowRef] = new Set(ids);
    });
  }

  // Build an ordered queue per Case ID + Date + Agent + Type key from Raw_Cases,
  // in sheet order (top to bottom = chronological, since rows are only appended).
  const queues = {};
  if (rawSheet.getLastRow() > 1) {
    const rawData = rawSheet.getRange(2, 1, rawSheet.getLastRow() - 1, 14).getValues();
    rawData.forEach((r, i) => {
      const rawRowNumber = i + 2;
      const rowDateStr = toDateStringFast(r[RAW_COLS.DATE]);
      const agent = r[RAW_COLS.AGENT] ? r[RAW_COLS.AGENT].toString().trim().toLowerCase() : '';
      const caseType = r[RAW_COLS.CASE_TYPE];
      const ts = r[RAW_COLS.TIMESTAMP];
      const intervalStr = r[RAW_COLS.INTERVAL];
      const ids = String(r[RAW_COLS.CASE_IDS] || '').split(',').map(s => s.trim()).filter(Boolean);
      const flaggedSet = flaggedByRow[rawRowNumber] || new Set();

      ids.forEach(id => {
        const key = id + '|' + rowDateStr + '|' + agent + '|' + caseType;
        if (!queues[key]) queues[key] = [];
        queues[key].push({ ts: ts, interval: intervalStr, flagged: flaggedSet.has(id) });
      });
    });
  }

  // Walk Index_CaseIDs in sheet order, consuming one queue entry per row.
  let filled = 0;
  if (indexSheet.getLastRow() > 1) {
    const indexData = indexSheet.getRange(2, 1, indexSheet.getLastRow() - 1, 7).getValues();
    const updates = []; // { row, values: [ts, interval, flagged] }

    indexData.forEach((r, i) => {
      if (r[4]) return; // already has a timestamp, skip
      const key = r[0] + '|' + toDateStringFast(r[2]) + '|' + String(r[3]).trim().toLowerCase() + '|' + r[1];
      const q = queues[key];
      if (!q || q.length === 0) return; // no matching Raw_Cases entry found

      const entry = q.shift();
      updates.push({ row: i + 2, values: [entry.ts, entry.interval, entry.flagged ? 'Yes' : 'No'] });
    });

    updates.forEach(u => {
      indexSheet.getRange(u.row, 5, 1, 3).setValues([u.values]);
    });
    filled = updates.length;
  }

  return { success: true, filled: filled };
}

// --- ESCALATION SWEEP LOGIC ---

function setupEscalationSweepTrigger() {
  requireManagerOrThrow();
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

function escalationSweep(e) {
  try {
    if (!e || !e.triggerUid) return; // trigger-only; blocks casual client-side calls
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty('ESCALATION_SWEEP_ENABLED') !== 'true') return;

    const lock = LockService.getScriptLock();
    if (!lock.tryLock(5000)) return;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const escalationLogSheet = ss.getSheetByName('Escalation_Log');
    if (!escalationLogSheet) return;

    const graceMinutes = parseInt(props.getProperty('GRACE_PERIOD_MINUTES') || '30', 10);
    const now = new Date();

    // Read the log once per sweep run, not once per hour-offset checked below.
    const logData = escalationLogSheet.getDataRange().getValues();

    for (let offset = 0; offset <= 6; offset++) {
      const targetTime = new Date(now.getTime() - (offset * 60 * 60 * 1000));
      const targetDateStr = toDateStringFast(targetTime);
      const h = targetTime.getHours();

      const intervalEnd = new Date(targetTime.getFullYear(), targetTime.getMonth(), targetTime.getDate(), h + 1, 0, 0, 0);
      const graceThreshold = new Date(intervalEnd.getTime() + (graceMinutes * 60000));

      if (now.getTime() > graceThreshold.getTime()) {
        // Canonical 24-hour form ("16:00") — the same shape the UI's <select>
        // sends, and the only form getIntervalData / getCheckInStatus /
        // getIntervalStatusOverrides expect. A separate 12-hour label is built
        // just for the human-readable email below.
        const intervalHourObj = new Date(targetTime);
        intervalHourObj.setHours(h, 0, 0, 0);
        const intervalHourStr = Utilities.formatDate(intervalHourObj, Session.getScriptTimeZone(), "H:00");
        const intervalHourLabel = Utilities.formatDate(intervalHourObj, Session.getScriptTimeZone(), "h:00 a");

        let alreadyEscalated = false;
        for (let i = 1; i < logData.length; i++) {
          const rowDate = logData[i][0];
          const formattedRowDate = (rowDate instanceof Date) ? toDateStringFast(rowDate) : rowDate;
          if (formattedRowDate == targetDateStr && parseHourToInt(logData[i][1]) === parseHourToInt(intervalHourStr)) {
            alreadyEscalated = true;
            break;
          }
        }
        if (alreadyEscalated) continue;

        const checkIn = getCheckInStatus(targetDateStr, intervalHourStr);
        if (checkIn) continue;

        const agents = getIntervalData_(targetDateStr, intervalHourStr);
        if (!agents || agents.length === 0) continue;

        const statusOverrides = getIntervalStatusOverrides_(targetDateStr, intervalHourStr);
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
            const match = pocResult.schedule.find(s => parseHourToInt(s.time) === parseHourToInt(intervalHourStr));
            if (match) scheduledPOC = match.poc;
          } else if (Array.isArray(pocResult)) {
            const match = pocResult.find(s => parseHourToInt(s.time) === parseHourToInt(intervalHourStr));
            if (match) scheduledPOC = match.poc;
          }

          const escalatedAt = new Date();
          escalationLogSheet.appendRow([targetDateStr, intervalHourStr, escalatedAt, scheduledPOC, unresolvedAgents.length, agents.length]);

          sendEscalationEmail_(targetDateStr, intervalHourLabel, scheduledPOC, escalatedAt, unresolvedAgents, agents.length);
        }
      }
    }
  } catch (e) {
    logError('escalationSweep', e.toString(), 'SYSTEM');
  } finally {
    try { LockService.getScriptLock().releaseLock(); } catch(e){}
  }
}

// Returns the scheduled shift start hour (0-23) for ldap on dateStr from the
// 'Agent Shifts' sheet, or null if off/VL/LOA/AWOL or not found. Used to
// bound the Absent->Late conversion window below. OT agents have no entry
// here — intentional, since "late" only makes sense against a scheduled SOS.
function getAgentShiftStartHour_(ldap, dateStr) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const shiftSheet = ss.getSheetByName('Agent Shifts');
    if (!shiftSheet || shiftSheet.getLastRow() <= 2) return null;

    const shiftData = shiftSheet.getDataRange().getValues();
    const headers = shiftData[1];

    let dateColIdx = -1;
    for (let c = 8; c < headers.length; c++) {
      let cellDate = headers[c];
      let formattedCellDate = "";
      try {
        if (cellDate instanceof Date) formattedCellDate = toDateStringFast(cellDate);
        else if (cellDate) formattedCellDate = toDateStringFast(new Date(cellDate));
      } catch (e) {}
      if (formattedCellDate === dateStr) { dateColIdx = c; break; }
    }
    if (dateColIdx === -1) return null;

    for (let r = 2; r < shiftData.length; r++) {
      const rowLdap = shiftData[r][0] ? shiftData[r][0].toString().trim().toLowerCase() : '';
      if (rowLdap !== ldap.toLowerCase()) continue;

      const shiftVal = shiftData[r][dateColIdx];
      if (!shiftVal || shiftVal === "OFF" || shiftVal === "VL" || shiftVal === "LOA" || shiftVal === "AWOL") return null;

      if (shiftVal instanceof Date) return shiftVal.getHours();
      if (typeof shiftVal === 'string' && shiftVal.includes(':')) return parseInt(shiftVal.split(':')[0], 10);
      if (typeof shiftVal === 'number') return Math.round(shiftVal * 24);
      return null;
    }
    return null;
  } catch (e) {
    logError('getAgentShiftStartHour', e.toString(), ldap);
    return null;
  }
}

// When an agent who was tagged "Absent" for one or more intervals earlier in
// their own shift today finally submits cases, those specific intervals are
// reclassified as "Late" — they did show up, just later than scheduled.
// Only intervals strictly BEFORE arrivalHour are touched; anything tagged
// Absent from that point forward is left alone (a real mid-shift absence,
// not tardiness). Statuses other than exactly "Absent" (VL/SL, on Live
// Channel, etc.) were a deliberate POC call and are never overwritten here.
function convertAbsentToLate_(ldap, dateStr, arrivalHour) {
  try {
    const shiftStartHour = getAgentShiftStartHour_(ldap, dateStr);
    if (shiftStartHour === null) return; // OT or no scheduled shift — "Late" doesn't apply

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Interval_Status');
    if (!sheet || sheet.getLastRow() <= 1) return;

    const data = sheet.getDataRange().getValues();
    const targetLdap = ldap.toLowerCase();

    for (let i = 1; i < data.length; i++) {
      const rowDate = data[i][0];
      const formattedDate = (rowDate instanceof Date) ? toDateStringFast(rowDate) : rowDate;
      if (formattedDate !== dateStr) continue;

      const rowLdap = String(data[i][2]).trim().toLowerCase();
      if (rowLdap !== targetLdap) continue;

      const rowHour = parseHourToInt(data[i][1]);
      if (rowHour === null || rowHour < shiftStartHour || rowHour >= arrivalHour) continue;

      if (String(data[i][3]).trim() === 'Absent') {
        sheet.getRange(i + 1, 4).setValue('Late');
        sheet.getRange(i + 1, 5).setValue('System (auto)');
        sheet.getRange(i + 1, 6).setValue(new Date());
      }
    }
  } catch (e) {
    logError('convertAbsentToLate', e.toString(), ldap);
  }
}

function sendEscalationEmail_(dateStr, intervalHourStr, scheduledPOC, escalatedAt, unresolvedAgents, totalAgents) {
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

    // OT timeline granularity: hourly for a short range (readable), daily once it
    // would otherwise cram too many points onto one axis. Scaffold every bucket
    // to 0 up front so a gap in the chart means "no OT that hour/day", not
    // "no submission happened to land in our map".
    const rangeDaySpan = Math.round((end.getTime() - start.getTime()) / 86400000) + 1;
    const otTimelineHourly = rangeDaySpan <= 2;
    const otTimelineBuckets = {}; // epoch ms of bucket start -> valid OT cases
    if (otTimelineHourly) {
      for (let t = startOfDay_(start).getTime(); t <= end.getTime(); t += 3600000) {
        otTimelineBuckets[t] = 0;
      }
    } else {
      for (let t = startOfDay_(start).getTime(); t <= end.getTime(); t += 86400000) {
        otTimelineBuckets[t] = 0;
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
      otLeaderboard: {} // { agent: validOTCases }
    };

    // 1. Pending Audits KPI
    if (auditSheet && auditSheet.getLastRow() > 1) {
      const auditData = auditSheet.getRange(2, 1, auditSheet.getLastRow() - 1, 1).getValues();
      data.kpis.pendingAudits = auditData.filter(r => String(r[0]).includes('PENDING')).length;
    }

    if (!rawSheet || rawSheet.getLastRow() <= 1) return data;

    // Fetch up to Column 17 (OT Type is Col 16, Index 16. Valid is Col 11. Shift Type is Col 8), from the range's first day
    const rawData = readSheetTail_(rawSheet, 17, 1, startOfDay_(start)).values;

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

          const ts = (r[RAW_COLS.TIMESTAMP] instanceof Date) ? r[RAW_COLS.TIMESTAMP] : new Date(r[RAW_COLS.TIMESTAMP]);
          if (!isNaN(ts.getTime())) {
            const bucketKey = otTimelineHourly
              ? new Date(ts.getFullYear(), ts.getMonth(), ts.getDate(), ts.getHours()).getTime()
              : startOfDay_(ts).getTime();
            if (bucketKey in otTimelineBuckets) {
              otTimelineBuckets[bucketKey] += validCases;
            }
          }
        }
      }
    });

    data.kpis.activeStaffToday = activeAgentsToday.size;
    data.kpis.activeStaffThisHour = activeAgentsThisHour.size;

    // Transform Heatmap data for ApexCharts
    // Sort dates ascending
    const sortedDates = Object.keys(heatmapDataMap).sort((a, b) => new Date(a) - new Date(b));
    const HOUR_LABELS = ["12:00 AM","1:00 AM","2:00 AM","3:00 AM","4:00 AM","5:00 AM","6:00 AM","7:00 AM","8:00 AM","9:00 AM","10:00 AM","11:00 AM","12:00 PM","1:00 PM","2:00 PM","3:00 PM","4:00 PM","5:00 PM","6:00 PM","7:00 PM","8:00 PM","9:00 PM","10:00 PM","11:00 PM"];
    // Show only active hours, in shift order: the window starts right after the longest
    // stretch of empty hours, so a shift crossing midnight stays one block (e.g. 2p ... 1a).
    const activeHr = new Array(24).fill(false);
    Object.keys(heatmapDataMap).forEach(function(d) {
      HOUR_LABELS.forEach(function(h, i) { if (heatmapDataMap[d][h]) activeHr[i] = true; });
    });
    let startH = 0, span = 24, bestGap = 0;
    for (let s = 0; s < 24; s++) {
      if (activeHr[s] || !activeHr[(s + 23) % 24]) continue; // first empty hour of each gap only
      let len = 0;
      while (len < 24 && !activeHr[(s + len) % 24]) len++;
      if (len < 24 && len > bestGap) { bestGap = len; startH = (s + len) % 24; span = 24 - len; }
    }
    const allHours = [];
    for (let k = 0; k < span; k++) allHours.push(HOUR_LABELS[(startH + k) % 24]);
    // Reverse the dates so newest is on top of the Y-axis for standard heatmap look
    sortedDates.reverse().forEach(dateStr => {
      // Get a short friendly name (e.g. "Mon, Sep 24")
      const d = new Date(dateStr);
      const friendlyName = Utilities.formatDate(d, Session.getScriptTimeZone(), "EEE, MMM d");

      const daySeries = { name: friendlyName, data: [] };
      allHours.forEach(hour => {
        daySeries.data.push({
          x: hour.replace(':00 AM', 'a').replace(':00 PM', 'p'),
          y: heatmapDataMap[dateStr][hour] || 0
        });
      });
      data.heatmap.push(daySeries);
    });

    // OT timeline: chronological array of {x: epoch ms, y: valid OT cases},
    // one point per hour or per day depending on otTimelineHourly above.
    data.otTimeline = Object.keys(otTimelineBuckets)
      .map(Number)
      .sort((a, b) => a - b)
      .map(t => ({ x: t, y: otTimelineBuckets[t] }));
    data.otTimelineGranularity = otTimelineHourly ? 'hourly' : 'daily';

    return data;
  } catch (e) {
    const user = Session.getActiveUser().getEmail() || 'Unknown';
    logError('getAnalyticsData', e.toString(), user);
    throw new Error("Unable to fetch analytics data. Check network and retry.");
  }
}


// --- INTERVAL STATUS & CHECK-IN LOGIC ---

function getIntervalStatusOverrides(dateStr, intervalHourStr) {
  try { requireManagerOrThrow(); } catch (e) { return {}; }
  return getIntervalStatusOverrides_(dateStr, intervalHourStr);
}

function getIntervalStatusOverrides_(dateStr, intervalHourStr) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Interval_Status');
    if (!sheet) return {};

    const data = sheet.getDataRange().getValues();
    const overrides = {};

    for (let i = 1; i < data.length; i++) {
      const rowDate = data[i][0];
      const formattedDate = (rowDate instanceof Date) ? toDateStringFast(rowDate) : rowDate;
      const rowInterval = data[i][1];

      if (formattedDate === dateStr && parseHourToInt(rowInterval) === parseHourToInt(intervalHourStr)) {
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

function setIntervalStatusBulk(dateStr, intervalHourStr, updates) {
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
    const now = new Date();

    // Process updates
    updates.forEach(update => {
      const targetLdap = String(update.ldap).trim().toLowerCase();
      let foundRow = -1;

      for (let i = 1; i < data.length; i++) {
        const rowDate = data[i][0];
        const formattedDate = (rowDate instanceof Date) ? toDateStringFast(rowDate) : rowDate;
        const rowInterval = data[i][1];
        const rowLdap = String(data[i][2]).trim().toLowerCase();

        if (formattedDate === dateStr && parseHourToInt(rowInterval) === parseHourToInt(intervalHourStr) && rowLdap === targetLdap) {
          foundRow = i + 1;
          break;
        }
      }

      if (foundRow !== -1) {
        sheet.getRange(foundRow, 4).setValue(update.status);
        sheet.getRange(foundRow, 5).setValue(profile.ldap);
        sheet.getRange(foundRow, 6).setValue(now);
        data[foundRow - 1][3] = update.status; // Update local array to prevent duplicate searches acting incorrectly
      } else {
        sheet.appendRow([dateStr, intervalHourStr, targetLdap, update.status, profile.ldap, now]);
        data.push([dateStr, intervalHourStr, targetLdap, update.status, profile.ldap, now]);
      }
    });

    return { success: true };
  } catch (e) {
    const user = Session.getActiveUser().getEmail() || 'Unknown';
    logError('setIntervalStatusBulk', e.toString(), user);
    return { success: false, error: "Failed to process bulk status updates." };
  } finally {
    lock.releaseLock();
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
      const rowLdap = String(data[i][2]).trim().toLowerCase();

      if (formattedDate === dateStr && parseHourToInt(rowInterval) === parseHourToInt(intervalHourStr) && rowLdap === targetLdap) {
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

      if (formattedDate === dateStr && parseHourToInt(rowInterval) === parseHourToInt(intervalHourStr)) {
        return {
          date: formattedDate,
          interval: intervalHourStr,
          scheduledPOC: data[i][2],
          checkedInBy: data[i][3],
          timestamp: (data[i][4] instanceof Date) ? data[i][4].getTime() : data[i][4],
          result: data[i][5],
          unresolvedLDAPs: data[i][6],
          notes: data[i][7],
          isSubstitute: data[i][8] === true || data[i][8] === 'true'
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
      return { success: false, error: "No agents scheduled for this interval — check-in not applicable." };
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
      const match = pocResult.schedule.find(s => parseHourToInt(s.time) === parseHourToInt(intervalHourStr));
      if (match) scheduledPOC = match.poc;
    } else if (Array.isArray(pocResult)) {
      const match = pocResult.find(s => parseHourToInt(s.time) === parseHourToInt(intervalHourStr));
      if (match) scheduledPOC = match.poc;
    }

    let isSubstitute = false;
    if (scheduledPOC !== "Unknown" && scheduledPOC.toLowerCase() !== profile.ldap.toLowerCase()) {
       isSubstitute = true;
    }

    const now = new Date();
    sheet.appendRow([dateStr, intervalHourStr, scheduledPOC, profile.ldap, now, result, unresolvedStr, noteStr, isSubstitute]);

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
        notes: noteStr,
        isSubstitute: isSubstitute
      }
    };
  } catch (e) {
    const user = Session.getActiveUser().getEmail() || 'Unknown';
    logError('checkInInterval', e.toString(), user);
    return { success: false, error: "System encountered an error during check-in." };
  } finally {
    lock.releaseLock();
  }
}

// 9b. Fetch Live POC Schedule
function getPOCSchedule() {
  try {
    const cache = CacheService.getScriptCache();
    const now = new Date();
    const todayStr = toDateStringFast(now); // "M/d/yyyy"
    const cacheKey = 'poc_schedule_' + todayStr;
    const cachedData = cache.get(cacheKey);

    if (cachedData) {
      try {
        const parsed = JSON.parse(cachedData);
        return { success: true, schedule: parsed };
      } catch (e) {
        // Failed to parse, ignore cache and fetch
      }
    }

    const pocSheetUrl = 'https://docs.google.com/spreadsheets/d/1SwO6Wet3OWPQDkXC2jyQ3rbPTjCDMZbAc5fLbDK6HHU/edit';
    const ss = SpreadsheetApp.openByUrl(pocSheetUrl);
    const sheet = ss.getSheetByName('POC Schedule');

    if (!sheet) {
      return { success: false, error: 'POC Schedule tab not found in the source sheet.' };
    }

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

    // Cache the successful fetch for 10 minutes to significantly speed up Live POC loading
    try {
      cache.put(cacheKey, JSON.stringify(pocList), 600); // 10 minutes
    } catch(err) {
      // Ignore cache put errors
    }

    return { success: true, schedule: pocList };

  } catch (e) {
    const user = Session.getActiveUser().getEmail() || 'Unknown';
    logError('getPOCSchedule', e.toString(), user);
    return { success: false, error: 'Failed to fetch POC Schedule.' };
  }
}

// Lightweight check used by client-side polling to detect new activity for
// an interval without recomputing the full schedule/status table. Compares
// row count + summed valid cases against what the client last loaded.
function getIntervalActivitySignature(dateStr, intervalHourStr) {
  requireManagerOrThrow();
  return getIntervalActivitySignature_(dateStr, intervalHourStr);
}

function getIntervalActivitySignature_(dateStr, intervalHourStr) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const rawSheet = ss.getSheetByName('Raw_Cases');
    if (!rawSheet || rawSheet.getLastRow() <= 1) return { rowCount: 0, totalValid: 0 };

    const targetHour = parseInt(intervalHourStr.split(':')[0], 10);
    let targetDateObj = new Date();
    targetDateObj.setHours(targetHour, 0, 0, 0);
    const intervalLabel = Utilities.formatDate(targetDateObj, Session.getScriptTimeZone(), "h:00 a");

    const sp = dateStr.split('/');
    const rawData = readSheetTail_(rawSheet, 12, 1, new Date(parseInt(sp[2], 10), parseInt(sp[0], 10) - 1, parseInt(sp[1], 10))).values;
    let rowCount = 0, totalValid = 0;
    rawData.forEach(r => {
      const rowDateStr = toDateStringFast(r[RAW_COLS.DATE]);
      const rowInterval = (r[RAW_COLS.INTERVAL] instanceof Date)
        ? Utilities.formatDate(r[RAW_COLS.INTERVAL], Session.getScriptTimeZone(), "h:00 a")
        : r[RAW_COLS.INTERVAL];
      if (rowDateStr === dateStr && rowInterval === intervalLabel) {
        rowCount++;
        totalValid += Number(r[RAW_COLS.VALID]) || 0;
      }
    });
    return { rowCount: rowCount, totalValid: totalValid };
  } catch (e) {
    logError('getIntervalActivitySignature', e.toString(), Session.getActiveUser().getEmail());
    return null;
  }
}

// 9. Fetch Data for Interval View
function getIntervalData(dateStr, intervalHourStr) {
  requireManagerOrThrow();
  return getIntervalData_(dateStr, intervalHourStr);
}

function getIntervalData_(dateStr, intervalHourStr) {
  try {
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
        const ldap = masterData[i][0] ? masterData[i][0].toString().trim().toLowerCase() : '';
        if (!ldap) continue;
        
        const channel = masterData[i][23] ? masterData[i][23].toString().toLowerCase() : '';
        const supervisor = masterData[i][27] ? masterData[i][27].toString() : 'Unknown';
        if(ldap) emailAgents[ldap] = { isEmail: channel === 'email', supervisor: supervisor };
      }
      try { cache.put('emailAgentsMap_v2', JSON.stringify(emailAgents), 14400); } catch (e) {}
    }

        // 2. Process Regular Shifts from 'Agent Shifts'
    if (shiftSheet && shiftSheet.getLastRow() > 2) {
      const shiftData = shiftSheet.getDataRange().getValues();
      const headers = shiftData[1]; // Row 2 holds the actual dates

      const findDateCol = function(targetStr) {
        for (let c = 8; c < headers.length; c++) {
          const cellDate = headers[c];
          let f = "";
          try {
            if (cellDate instanceof Date) f = toDateStringFast(cellDate);
            else if (cellDate) f = toDateStringFast(new Date(cellDate));
          } catch (e) {}
          if (f === targetStr) return c;
        }
        return -1;
      };

      // Previous calendar day: shifts that started yesterday and run past midnight
      const dp = dateStr.split('/');
      const prevDateStr = toDateStringFast(new Date(parseInt(dp[2], 10), parseInt(dp[0], 10) - 1, parseInt(dp[1], 10) - 1));

      const passes = [
        { col: findDateCol(dateStr), isPrevDay: false },
        { col: findDateCol(prevDateStr), isPrevDay: true }
      ];

      passes.forEach(function(pass) {
        if (pass.col === -1) return;

        for (let r = 2; r < shiftData.length; r++) {
          const ldap = shiftData[r][0] ? shiftData[r][0].toString().trim().toLowerCase() : '';
          if (!emailAgents[ldap] || !emailAgents[ldap].isEmail) continue;
          if (agentsInInterval[ldap]) continue;

          const site = shiftData[r][4]; // Col E
          const shiftVal = shiftData[r][pass.col];
          if (!shiftVal || shiftVal === "OFF" || shiftVal === "VL" || shiftVal === "LOA" || shiftVal === "AWOL") continue;

          let startHour = -1;
          if (shiftVal instanceof Date) {
            startHour = shiftVal.getHours();
          } else if (typeof shiftVal === 'string' && shiftVal.includes(':')) {
            startHour = parseInt(shiftVal.split(':')[0], 10);
          } else if (typeof shiftVal === 'number') {
            startHour = Math.round(shiftVal * 24);
          }
          if (startHour === -1) continue;

          const endHour = startHour + 9;
          const isOnShift = pass.isPrevDay
            ? (endHour > 24 && targetHour < (endHour - 24))
            : (targetHour >= startHour && targetHour < Math.min(endHour, 24));
          if (!isOnShift) continue;

          const eosActual = endHour > 24 ? endHour - 24 : endHour;
          agentsInInterval[ldap] = {
            ldap: ldap,
            sos: startHour + ":00",
            eos: eosActual + ":00",
            site: site,
            supervisor: emailAgents[ldap] ? emailAgents[ldap].supervisor : 'Unknown',
            isOT: false,
            casesLogged: 0,
            regularCount: 0,
            manualCount: 0,
            reopenedCount: 0
          };
        }
      });
    }
    
    // 3. Process Overtime & Live Metrics from 'Raw_Cases'
    if (rawSheet && rawSheet.getLastRow() > 1) {
      const rp = dateStr.split('/');
      const rawData = readSheetTail_(rawSheet, 17, 1, new Date(parseInt(rp[2], 10), parseInt(rp[0], 10) - 1, parseInt(rp[1], 10))).values; 
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
function getMyProfileData(targetLdap) {
  try {
    const callerProfile = getUserProfile();
    const isOwnProfile = !targetLdap || targetLdap.trim().toLowerCase() === callerProfile.ldap.toLowerCase();

    if (!isOwnProfile && !callerProfile.isManager) {
      throw new Error("Access denied: manager permissions required to view other agents' profiles.");
    }

    const ldap = isOwnProfile ? callerProfile.ldap.toLowerCase() : targetLdap.trim().toLowerCase();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const masterSheet = ss.getSheetByName('Masterlist');

    let name = ldap, site = 'Unknown', lob = '', workflow = '';
    let details = {
      position: '', gradeLevel: '', employeeStatus: '', team: '',
      hireDateStr: '', tenureText: '',
      reportsTo: []
    };

    if (masterSheet && masterSheet.getLastRow() > 1) {
      // Through Col AS (Site) — 45 columns — covers both identity fields
      // (name/lob/workflow/site) and the org-detail fields (position, grade,
      // tenure, reports-to) in a single read, for either the caller's own
      // LDAP or a manager-requested target LDAP.
      const masterData = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 45).getValues();
      const byLdap = {};
      masterData.forEach(function(r) {
        const rowLdap = r[0] ? r[0].toString().trim().toLowerCase() : '';
        if (rowLdap) byLdap[rowLdap] = r;
      });

      const myRow = byLdap[ldap];
      if (myRow) {
        name = myRow[1] || ldap;         // Col B
        lob = myRow[21] || '';           // Col V
        workflow = myRow[22] || '';      // Col W
        site = myRow[44] || 'Unknown';   // Col AS

        details.position = myRow[15] || '';        // Col P
        details.gradeLevel = myRow[16] || '';       // Col Q
        details.employeeStatus = myRow[18] || '';   // Col S
        details.team = myRow[26] || '';              // Col AA

        const hireDateVal = myRow[6]; // Col G
        if (hireDateVal) {
          const hireDate = (hireDateVal instanceof Date) ? hireDateVal : new Date(hireDateVal);
          if (!isNaN(hireDate.getTime())) {
            details.hireDateStr = Utilities.formatDate(hireDate, Session.getScriptTimeZone(), 'MMM d, yyyy');
            details.tenureText = computeTenureText(hireDate);
          }
        }

        const sup1Name = myRow[27] || '';           // Col AB
        const sup1Ldap = myRow[28] ? myRow[28].toString().trim().toLowerCase() : ''; // Col AC
        if (sup1Ldap && sup1Ldap !== '-') {
          details.reportsTo.push({ name: sup1Name || sup1Ldap, ldap: sup1Ldap });

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

    const rawSheet = ss.getSheetByName('Raw_Cases');
    const auditSheet = ss.getSheetByName('Audit Queue');

    const stats = { today: 0, thisWeek: 0, lastWeek: 0, thisMonth: 0, pendingAudits: 0 };
    const breakdown = { Regular: 0, Reopened: 0, Manual: 0, Telus: 0, Cimba: 0 };
    const history = {}; // dateKey -> { All, Regular, Reopened, Manual, Telus, Cimba, hasAbsent, hasLate }

    const now = new Date();
    const todayStr = toDateStringFast(now);
    const weekStart = getWeekStartMonday(now);
    const lastWeekStart = new Date(weekStart); lastWeekStart.setDate(lastWeekStart.getDate() - 7);
    const lastWeekEnd = new Date(weekStart); lastWeekEnd.setMilliseconds(-1);
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

    function ensureHistoryEntry(key) {
      if (!history[key]) {
        history[key] = { All: 0, Regular: 0, Reopened: 0, Manual: 0, Telus: 0, Cimba: 0, hasAbsent: false, hasLate: false };
      }
      return history[key];
    }

    if (rawSheet && rawSheet.getLastRow() > 1) {
      const rawData = rawSheet.getRange(2, 1, rawSheet.getLastRow() - 1, 17).getValues();
      rawData.forEach(function(r) {
        const rowLdap = r[RAW_COLS.AGENT] ? r[RAW_COLS.AGENT].toString().trim().toLowerCase() : '';
        if (rowLdap !== ldap) return;

        const rowDateObj = (r[RAW_COLS.DATE] instanceof Date) ? r[RAW_COLS.DATE] : new Date(r[RAW_COLS.DATE]);
        const rowDateStr = toDateStringFast(rowDateObj);
        const validCases = Number(r[RAW_COLS.VALID]) || 0;
        if (validCases <= 0) return;

        if (rowDateStr === todayStr) stats.today += validCases;
        if (rowDateObj >= weekStart) stats.thisWeek += validCases;
        if (rowDateObj >= lastWeekStart && rowDateObj <= lastWeekEnd) stats.lastWeek += validCases;
        if (rowDateObj >= monthStart) stats.thisMonth += validCases;
        if (trendMap.hasOwnProperty(rowDateStr)) trendMap[rowDateStr] += validCases;

        const caseType = r[RAW_COLS.CASE_TYPE] || '';
        let typeCategory = 'Regular';
        if (caseType === 'Manual Assignment') typeCategory = 'Manual';
        else if (caseType === 'Reopened Cases') typeCategory = 'Reopened';
        else if (caseType === 'Telus Cases') typeCategory = 'Telus';
        else if (caseType === 'Cimba Cases') typeCategory = 'Cimba';

        if (rowDateObj >= monthStart && breakdown.hasOwnProperty(typeCategory)) {
          breakdown[typeCategory] += validCases;
        }

        const entry = ensureHistoryEntry(toISODateStringFast(rowDateObj));
        entry.All += validCases;
        if (entry[typeCategory] !== undefined) entry[typeCategory] += validCases;
        else entry.Regular += validCases;
      });
    }

    // Merge Absent/Late tags so the calendar can flag a true no-show day
    // distinctly from an ordinary zero-case day (rest day, VL, etc. with no
    // tag at all), and mark days the agent started late.
    const statusSheet = ss.getSheetByName('Interval_Status');
    if (statusSheet && statusSheet.getLastRow() > 1) {
      const statusData = statusSheet.getRange(2, 1, statusSheet.getLastRow() - 1, 4).getValues();
      statusData.forEach(function(r) {
        const rowLdap = r[2] ? r[2].toString().trim().toLowerCase() : '';
        if (rowLdap !== ldap) return;
        const val = String(r[3]).trim();
        if (val !== 'Absent' && val !== 'Late') return;

        const rowDate = r[0];
        const rowDateObj = (rowDate instanceof Date) ? rowDate : new Date(rowDate);
        if (isNaN(rowDateObj.getTime())) return;

        const entry = ensureHistoryEntry(toISODateStringFast(rowDateObj));
        if (val === 'Absent') entry.hasAbsent = true;
        if (val === 'Late') entry.hasLate = true;
      });
    }

    if (auditSheet && auditSheet.getLastRow() > 1) {
      const auditData = auditSheet.getRange(2, 1, auditSheet.getLastRow() - 1, 3).getValues();
      auditData.forEach(function(r) {
        const rowLdap = r[2] ? r[2].toString().trim().toLowerCase() : '';
        if (rowLdap === ldap && String(r[0]).includes('PENDING')) stats.pendingAudits++;
      });
    }

    return {
      ldap: ldap,
      isOwnProfile: isOwnProfile,
      viewerIsManager: callerProfile.isManager,
      name: name,
      site: site,
      workflow: workflow,
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

function testEscalationEmail() {
  requireManagerOrThrow();
  sendEscalationEmail_(
    '9/17/2026',
    '4:00 PM',
    'stevenjosephc',
    new Date(),
    ['testagent1', 'testagent2'],
    5
  );
}

// --- EXPIRED / EXPIRING CASES ---
// Source: "CNX PLAY Case_Mon" (updated by hand roughly every 2 hours). Every pool tab after
// CASE_MON_BOUNDARY_TAB is scanned (Case ID = col E, Case Status = col F) and mirrored into the
// 'Expiring Case IDs' tab, one row per Case ID: Open -> Worked (an agent logged it) or Cleared (left the source list).
// Finished rows move to 'Expiring_History' after EXP_KEEP_RESOLVED_DAYS; analytics read queue + history.
const CASE_MON_SS_ID = '1msc87JAze65oA7PGnTAKDLamov4HNP8w-2l2c5cqU4s';
const CASE_MON_BOUNDARY_TAB = 'ND Pool Monitoring';
const EXPIRING_SHEET = 'Expiring Case IDs';
const EXPIRING_HISTORY_SHEET = 'Expiring_History';
const EXPIRING_HEADERS = ['Case ID', 'Case Status', 'Pool', 'Type', 'First Seen', 'Last Seen', 'State', 'Resolved At', 'Resolved By', 'Hours Left', 'Missed Scans', 'First Type', 'First Hours', 'Expired At', 'Worked At', 'Worked By', 'Left At'];
const EXPIRING_LEGACY_COLS = 11;       // the first 11 headers are the previous layout (migrated in place)
const EXP_COLS = { ID: 0, STATUS: 1, POOL: 2, TYPE: 3, FIRST: 4, LAST: 5, STATE: 6, RESOLVED_AT: 7, RESOLVED_BY: 8, HOURS: 9, MISSED: 10, FIRST_TYPE: 11, FIRST_HOURS: 12, EXPIRED_AT: 13, WORKED_AT: 14, WORKED_BY: 15, LEFT_AT: 16 };
const EXP_TYPE_EXPIRED = 'Expired';
const EXP_TYPE_15 = 'Expiring (<15h)';
const EXP_TYPE_24 = 'Expiring (<24h)';
const EXP_MISS_LIMIT = 2;                  // consecutive scans a case must be missing before it counts as gone
const EXP_WORKED_REOPEN_HOURS = 24;        // a Worked case still/again listed after this long starts a new lifecycle
const EXP_WORKED_LOOKBACK_HOURS = 2;       // agent logs up to this long before First Seen still count (source list can be stale)
const EXP_WORKED_AFTER_CLEAR_HOURS = 24;   // an agent log up to this long after a case is Cleared is still credited
const EXP_KEEP_RESOLVED_DAYS = 3;          // Worked/Cleared rows older than this move to Expiring_History
const EXP_DATE_FMT = 'M/d/yyyy h:mm AM/PM';
const EXP_CASE_ID_RE = /^\d-\d{7,14}$/;

function applyExpiringFormats_(sheet) {
  const rows = Math.max(sheet.getMaxRows() - 1, 1);
  [EXP_COLS.FIRST, EXP_COLS.LAST, EXP_COLS.RESOLVED_AT, EXP_COLS.EXPIRED_AT, EXP_COLS.WORKED_AT, EXP_COLS.LEFT_AT].forEach(function(c) {
    sheet.getRange(2, c + 1, rows, 1).setNumberFormat(EXP_DATE_FMT);
  });
}

function ensureExpiringSheet_(ss) {
  let sheet = ss.getSheetByName(EXPIRING_SHEET);
  if (!sheet) sheet = ss.insertSheet(EXPIRING_SHEET);
  const n = EXPIRING_HEADERS.length;
  if (sheet.getMaxColumns() < n) sheet.insertColumnsAfter(sheet.getMaxColumns(), n - sheet.getMaxColumns());
  const current = sheet.getRange(1, 1, 1, n).getValues()[0];
  if (current.join('|') !== EXPIRING_HEADERS.join('|')) {
    const legacy = EXPIRING_HEADERS.slice(0, EXPIRING_LEGACY_COLS).join('|');
    if (current.slice(0, EXPIRING_LEGACY_COLS).join('|') !== legacy) sheet.clear(); // unknown/old layout: start fresh
    // previous 11-column layout: rows are kept, only the new columns are added
    sheet.getRange(1, 1, 1, n).setValues([EXPIRING_HEADERS]).setFontWeight('bold').setBackground('#f3f3f3');
    sheet.setFrozenRows(1);
    applyExpiringFormats_(sheet);
  }
  return sheet;
}

function ensureExpiringHistorySheet_(ss) {
  const headers = EXPIRING_HEADERS.concat(['Archived At']);
  let sheet = ss.getSheetByName(EXPIRING_HISTORY_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(EXPIRING_HISTORY_SHEET);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold').setBackground('#f3f3f3');
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function appendExpiringHistory_(ss, rows) {
  if (!rows || rows.length === 0) return;
  const sheet = ensureExpiringHistorySheet_(ss);
  const n = EXPIRING_HEADERS.length + 1;
  const start = sheet.getLastRow() + 1;
  const missing = start + rows.length - 1 - sheet.getMaxRows();
  if (missing > 0) sheet.insertRowsAfter(sheet.getMaxRows(), missing);
  sheet.getRange(start, 1, rows.length, n).setValues(rows);
}

// Returns { type, hours } for a Case Status string, or null if it is not expired/expiring (e.g. "SLA met").
function classifyExpiringStatus_(statusRaw) {
  const s = String(statusRaw || '').trim().toLowerCase();
  if (!s || s === 'sla met') return null;
  if (s.indexOf('ago') !== -1) return { type: EXP_TYPE_EXPIRED, hours: null };
  const m = s.match(/(\d+)/);
  if (!m) return null;
  let hours = parseInt(m[1], 10);
  if (s.indexOf('day') !== -1) hours *= 24;
  else if (s.indexOf('min') !== -1) hours = 0;
  if (hours > 24) return null;
  return { type: hours < 15 ? EXP_TYPE_15 : EXP_TYPE_24, hours: hours };
}

// Reads every pool tab in Case_Mon. Throws if the source cannot be read, before anything is written.
function scanCaseMon_() {
  const src = SpreadsheetApp.openById(CASE_MON_SS_ID);
  const sheets = src.getSheets();
  let boundary = -1;
  for (let b = 0; b < sheets.length; b++) {
    if (sheets[b].getName() === CASE_MON_BOUNDARY_TAB) { boundary = b; break; }
  }
  if (boundary === -1) throw new Error('Boundary tab "' + CASE_MON_BOUNDARY_TAB + '" not found in Case_Mon.');

  const found = {}; // caseId -> { caseId, status, pool, type, hours }
  const rankOf = function(hours) { return hours === null ? -1 : hours; };
  for (let i = boundary + 1; i < sheets.length; i++) {
    const sh = sheets[i];
    if (sh.isSheetHidden()) continue;
    const last = sh.getLastRow();
    if (last < 2) continue;
    sh.getRange(2, 5, last - 1, 2).getValues().forEach(function(row) {
      const id = String(row[0]).trim();
      if (!EXP_CASE_ID_RE.test(id)) return; // also skips header rows
      const cls = classifyExpiringStatus_(row[1]);
      if (!cls) return;
      const prev = found[id];
      if (prev && rankOf(prev.hours) <= rankOf(cls.hours)) return; // same ID on two tabs: keep the most urgent
      found[id] = { caseId: id, status: String(row[1]).trim(), pool: sh.getName(), type: cls.type, hours: cls.hours };
    });
  }
  return found;
}

// Credits agent logs from Index_CaseIDs. Open rows become Worked; Cleared rows (recently cleared) keep their
// state but get Worked At/By, so we can measure both "left the queue" and "an agent actually worked it".
function markWorkedFromIndex_(rows, nowMs) {
  const C = EXP_COLS;
  const cand = {};
  let earliest = null;
  rows.forEach(function(r) {
    if (r[C.WORKED_AT]) return;
    const st = r[C.STATE];
    if (st === 'Cleared') {
      const t = (r[C.RESOLVED_AT] instanceof Date) ? r[C.RESOLVED_AT].getTime() : 0;
      if (!t || nowMs - t > EXP_WORKED_AFTER_CLEAR_HOURS * 3600000) return;
    } else if (st !== 'Open' && st !== '') {
      return;
    }
    cand[String(r[C.ID]).trim()] = r;
    const fs = (r[C.FIRST] instanceof Date) ? r[C.FIRST] : null;
    if (fs && (!earliest || fs < earliest)) earliest = fs;
  });
  if (!earliest) return 0;
  const idxSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Index_CaseIDs');
  if (!idxSheet || idxSheet.getLastRow() <= 1) return 0;

  const lookbackMs = EXP_WORKED_LOOKBACK_HOURS * 3600000;
  let since = new Date(earliest.getTime() - lookbackMs);
  const floor = new Date(nowMs - 3 * 86400000);
  if (since < floor) since = floor;

  let marked = 0;
  readSheetTail_(idxSheet, 7, 2, startOfDay_(since)).values.forEach(function(r) {
    const ts = r[4];
    if (!(ts instanceof Date)) return;
    const row = cand[String(r[0]).trim()];
    if (!row || row[C.WORKED_AT]) return;
    const fs = (row[C.FIRST] instanceof Date) ? row[C.FIRST] : null;
    if (!fs || ts.getTime() < fs.getTime() - lookbackMs) return;
    const agent = String(r[3]).trim().toLowerCase();
    row[C.WORKED_AT] = ts;
    row[C.WORKED_BY] = agent;
    if (row[C.STATE] === 'Open' || row[C.STATE] === '') {
      row[C.STATE] = 'Worked'; row[C.RESOLVED_AT] = ts; row[C.RESOLVED_BY] = agent;
    }
    marked++;
  });
  return marked;
}

// Called from submitCases (which already holds the script lock, so this must not lock).
function markExpiringWorked_(ids, ldap, ts) {
  if (!ids || ids.length === 0) return;
  const C = EXP_COLS;
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EXPIRING_SHEET);
  if (!sheet || sheet.getLastRow() <= 1) return;
  if (sheet.getRange(1, C.LEFT_AT + 1).getValue() !== EXPIRING_HEADERS[C.LEFT_AT]) return; // not migrated yet
  const want = new Set(ids);
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, C.WORKED_BY + 1).getValues();
  data.forEach(function(r, i) {
    if (!want.has(String(r[C.ID]).trim()) || r[C.WORKED_AT]) return;
    const row = i + 2;
    if (r[C.STATE] === 'Open' || r[C.STATE] === '') {
      sheet.getRange(row, C.STATE + 1, 1, 3).setValues([['Worked', ts, ldap]]);
      sheet.getRange(row, C.WORKED_AT + 1, 1, 2).setValues([[ts, ldap]]);
    } else if (r[C.STATE] === 'Cleared') {
      const t = (r[C.RESOLVED_AT] instanceof Date) ? r[C.RESOLVED_AT].getTime() : 0;
      if (t && ts.getTime() - t <= EXP_WORKED_AFTER_CLEAR_HOURS * 3600000) {
        sheet.getRange(row, C.WORKED_AT + 1, 1, 2).setValues([[ts, ldap]]);
      }
    }
  });
}

function runExpiringSync_(by) {
  // Read the slow external source BEFORE taking the script lock so agent submissions are never blocked by it
  const found = scanCaseMon_();
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ensureExpiringSheet_(ss);
    // (source was already read above, outside the lock)
    const now = new Date();
    const nowMs = now.getTime();
    const n = EXPIRING_HEADERS.length;
    const C = EXP_COLS;

    const prevLast = sheet.getLastRow();
    const existing = prevLast > 1 ? sheet.getRange(2, 1, prevLast - 1, n).getValues() : [];
    const byId = {};
    existing.forEach(function(r, i) { byId[String(r[C.ID]).trim()] = i; });

    // Safety: a source that suddenly lists nothing is more likely a half-refreshed sheet than "every case got assigned"
    const openBefore = existing.filter(function(r) { return r[C.STATE] === 'Open' || r[C.STATE] === ''; }).length;
    if (Object.keys(found).length === 0 && openBefore >= 5) {
      throw new Error('Case_Mon returned no expired/expiring cases, so the queue was left unchanged.');
    }

    const stats = { total: 0, added: 0, reopened: 0, cleared: 0, worked: 0, pruned: 0 };
    const seen = {};
    const archiveRows = [];

    Object.keys(found).forEach(function(id) {
      const f = found[id];
      const hoursVal = f.hours === null ? '' : f.hours;
      stats.total++;
      seen[id] = true;
      const idx = byId[id];
      if (idx === undefined) {
        existing.push([id, f.status, f.pool, f.type, now, now, 'Open', '', '', hoursVal, 0, f.type, hoursVal, f.type === EXP_TYPE_EXPIRED ? now : '', '', '', '']);
        byId[id] = existing.length - 1;
        stats.added++;
        return;
      }
      const r = existing[idx];
      const resolvedMs = (r[C.RESOLVED_AT] instanceof Date) ? r[C.RESOLVED_AT].getTime() : 0;
      if (r[C.STATE] === 'Worked' && resolvedMs && nowMs - resolvedMs > EXP_WORKED_REOPEN_HOURS * 3600000) {
        archiveRows.push(r.concat([now])); // close out the previous lifecycle in history
        r[C.STATE] = 'Open'; r[C.FIRST] = now; r[C.RESOLVED_AT] = ''; r[C.RESOLVED_BY] = '';
        r[C.FIRST_TYPE] = f.type; r[C.FIRST_HOURS] = hoursVal; r[C.EXPIRED_AT] = ''; r[C.WORKED_AT] = ''; r[C.WORKED_BY] = '';
        stats.reopened++;
      } else if (r[C.STATE] === 'Cleared') {
        r[C.STATE] = 'Open'; r[C.RESOLVED_AT] = ''; r[C.RESOLVED_BY] = '';
        stats.reopened++;
      }
      if (!r[C.FIRST_TYPE]) { r[C.FIRST_TYPE] = r[C.TYPE] || f.type; r[C.FIRST_HOURS] = r[C.HOURS]; } // rows from the old layout
      r[C.STATUS] = f.status;
      r[C.POOL] = f.pool;
      r[C.TYPE] = f.type;
      r[C.LAST] = now;
      r[C.HOURS] = hoursVal;
      r[C.MISSED] = 0;
      r[C.LEFT_AT] = '';
      if (f.type === EXP_TYPE_EXPIRED && !r[C.EXPIRED_AT]) r[C.EXPIRED_AT] = now;
    });

    // Cases missing from the source: only counted as gone after EXP_MISS_LIMIT scans in a row (guards against a half-refreshed source)
    existing.forEach(function(r) {
      const id = String(r[C.ID]).trim();
      const st = r[C.STATE];
      if (seen[id] || st === 'Cleared' || (r[C.LEFT_AT] instanceof Date)) return;
      r[C.MISSED] = (Number(r[C.MISSED]) || 0) + 1;
      if (r[C.MISSED] >= EXP_MISS_LIMIT) {
        r[C.LEFT_AT] = (r[C.LAST] instanceof Date) ? r[C.LAST] : now;
        if (st === 'Open' || st === '') {
          r[C.STATE] = 'Cleared'; r[C.RESOLVED_AT] = now; r[C.RESOLVED_BY] = 'Left source list';
          stats.cleared++;
        }
      }
    });

    stats.worked = markWorkedFromIndex_(existing, nowMs);

    const keepCutoff = nowMs - EXP_KEEP_RESOLVED_DAYS * 86400000;
    const kept = existing.filter(function(r) {
      if (r[C.STATE] === 'Open' || r[C.STATE] === '') return true;
      const t = (r[C.RESOLVED_AT] instanceof Date) ? r[C.RESOLVED_AT].getTime() : 0;
      if (t && t < keepCutoff) { archiveRows.push(r.concat([now])); stats.pruned++; return false; }
      return true;
    });

    appendExpiringHistory_(ss, archiveRows); // history first: a failure later can duplicate a row, never lose one

    if (kept.length > 0) {
      if (sheet.getMaxRows() < kept.length + 1) sheet.insertRowsAfter(sheet.getMaxRows(), kept.length + 1 - sheet.getMaxRows());
      sheet.getRange(2, 1, kept.length, n).setValues(kept);
    }
    if (prevLast > kept.length + 1) sheet.getRange(kept.length + 2, 1, prevLast - kept.length - 1, n).clearContent();

    // Remember when the source list itself last changed (drives the "source may be stale" hint and time-left estimate)
    const props = PropertiesService.getScriptProperties();
    const sigSrc = Object.keys(found).sort().map(function(id) { return id + '|' + found[id].status; }).join(';') || '-';
    const sig = Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, sigSrc));
    if (props.getProperty('EXP_SOURCE_SIG') !== sig || !props.getProperty('EXP_SOURCE_CHANGED_AT')) {
      props.setProperties({ EXP_SOURCE_SIG: sig, EXP_SOURCE_CHANGED_AT: String(nowMs) });
    }
    props.setProperties({ EXP_LAST_SCAN_AT: String(nowMs), EXP_LAST_SCAN_BY: String(by || '') });

    return stats;
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

// Manager button ("Grab now")
function grabExpiringCasesNow() {
  const profile = getUserProfile();
  if (!profile.isManager) return { success: false, error: 'Access denied: manager permissions required.' };
  try {
    return Object.assign({ success: true }, runExpiringSync_(profile.ldap));
  } catch (e) {
    logError('grabExpiringCasesNow', e.toString(), profile.ldap);
    return { success: false, error: 'Extraction failed: ' + e.message };
  }
}

// Time-based trigger target
function expiringCasesSweep(e) {
  if (!e || !e.triggerUid) return;
  try {
    runExpiringSync_('System (auto)');
  } catch (err) {
    logError('expiringCasesSweep', err.toString(), 'SYSTEM');
  }
}

// Run once from the editor to (re)create the every-30-minutes trigger.
function setupExpiringCasesTrigger() {
  requireManagerOrThrow();
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'expiringCasesSweep') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('expiringCasesSweep').timeBased().everyMinutes(30).create();
}

// Data for the Expired/Expiring Cases tab (manager only). Dates go to the client as epoch ms.
function getExpiringCases() {
  try {
    requireManagerOrThrow();
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(EXPIRING_SHEET);
    const props = PropertiesService.getScriptProperties();
    const meta = {
      lastScanAt: Number(props.getProperty('EXP_LAST_SCAN_AT')) || 0,
      sourceChangedAt: Number(props.getProperty('EXP_SOURCE_CHANGED_AT')) || 0
    };
    const rows = [];
    const ms = function(v) { return (v instanceof Date) ? v.getTime() : 0; };
    if (sheet && sheet.getLastRow() > 1 && sheet.getRange(1, EXP_COLS.FIRST + 1).getValue() === EXPIRING_HEADERS[EXP_COLS.FIRST]) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, EXPIRING_HEADERS.length).getValues().forEach(function(r) {
        const id = String(r[EXP_COLS.ID]).trim();
        if (!id) return;
        rows.push({
          caseId: id,
          status: String(r[EXP_COLS.STATUS]),
          pool: String(r[EXP_COLS.POOL]),
          type: r[EXP_COLS.TYPE],
          firstSeen: ms(r[EXP_COLS.FIRST]),
          lastSeen: ms(r[EXP_COLS.LAST]),
          state: r[EXP_COLS.STATE] || 'Open',
          resolvedAt: ms(r[EXP_COLS.RESOLVED_AT]),
          resolvedBy: String(r[EXP_COLS.RESOLVED_BY] || ''),
          hoursLeft: (r[EXP_COLS.HOURS] === '' || r[EXP_COLS.HOURS] === null) ? null : Number(r[EXP_COLS.HOURS])
        });
      });
    }
    const typeRank = function(t) { return t === EXP_TYPE_EXPIRED ? 0 : (t === EXP_TYPE_15 ? 1 : 2); };
    rows.sort(function(a, b) {
      const sa = a.state === 'Open' ? 0 : 1, sb = b.state === 'Open' ? 0 : 1;
      if (sa !== sb) return sa - sb;
      if (typeRank(a.type) !== typeRank(b.type)) return typeRank(a.type) - typeRank(b.type);
      const ha = a.hoursLeft === null ? -1 : a.hoursLeft, hb = b.hoursLeft === null ? -1 : b.hoursLeft;
      if (ha !== hb) return ha - hb;
      return a.firstSeen - b.firstSeen;
    });
    return { rows: rows, meta: meta };
  } catch (e) {
    const user = Session.getActiveUser().getEmail() || 'Unknown';
    logError('getExpiringCases', e.toString(), user);
    throw new Error('Unable to load expired/expiring cases. Please try again.');
  }
}

// --- EXPIRED / EXPIRING ANALYTICS ---
// Outcome per case (cases are grouped by the day/hour they were First Seen):
//   arrived    = already Expired when first seen (inherited backlog, reported separately)
//   saved      = left the queue (assigned/cleared or worked by an agent) before we ever saw it Expired
//   missedLate = we saw it go from Expiring to Expired while open, and it left the queue afterwards
//   missedOpen = same, and it is still open
//   open       = still open and not expired yet
// Saved rate = saved / (saved + missedLate + missedOpen).
function classifyExpiringCase_(r) {
  const C = EXP_COLS;
  const ms = function(v) { return (v instanceof Date) ? v.getTime() : 0; };
  const state = r[C.STATE] || 'Open';
  const type = r[C.TYPE];
  const firstSeen = ms(r[C.FIRST]);
  const firstType = r[C.FIRST_TYPE] || type;
  let expiredAt = ms(r[C.EXPIRED_AT]);
  if (!expiredAt && type === EXP_TYPE_EXPIRED) expiredAt = firstSeen;
  const workedAt = ms(r[C.WORKED_AT]) || (state === 'Worked' ? ms(r[C.RESOLVED_AT]) : 0);
  const workedBy = String((r[C.WORKED_BY] || (state === 'Worked' ? r[C.RESOLVED_BY] : '')) || '').trim().toLowerCase();
  const leftAt = ms(r[C.LEFT_AT]) || (state === 'Cleared' ? ms(r[C.RESOLVED_AT]) : 0);
  const exits = [workedAt, leftAt].filter(function(t) { return t > 0; });
  const exitAt = exits.length ? Math.max(firstSeen, Math.min.apply(null, exits)) : 0;

  let outcome;
  if (firstType === EXP_TYPE_EXPIRED) outcome = 'arrived';
  else if (exitAt && (!expiredAt || exitAt < expiredAt)) outcome = 'saved';
  else if (expiredAt && exitAt) outcome = 'missedLate';
  else if (expiredAt) outcome = 'missedOpen';
  else outcome = 'open';

  return {
    pool: String(r[C.POOL] || 'Unknown'), state: state, type: type, firstSeen: firstSeen, expiredAt: expiredAt,
    workedAt: workedAt, workedBy: workedBy, leftAt: leftAt, exitAt: exitAt, outcome: outcome,
    firstHours: (r[C.FIRST_HOURS] === '' || r[C.FIRST_HOURS] === null || r[C.FIRST_HOURS] === undefined) ? null : Number(r[C.FIRST_HOURS])
  };
}

function getExpiringAnalytics(startDateStr, endDateStr) {
  try {
    requireManagerOrThrow();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const tz = Session.getScriptTimeZone();
    const C = EXP_COLS;
    const now = new Date();
    const nowMs = now.getTime();

    let start = new Date(); start.setDate(start.getDate() - 6); start.setHours(0, 0, 0, 0);
    let end = new Date(); end.setHours(23, 59, 59, 999);
    if (startDateStr) { const p = startDateStr.split('-'); if (p.length === 3) start = new Date(p[0], p[1] - 1, p[2], 0, 0, 0, 0); }
    if (endDateStr) { const p = endDateStr.split('-'); if (p.length === 3) end = new Date(p[0], p[1] - 1, p[2], 23, 59, 59, 999); }
    const startMs = start.getTime(), endMs = end.getTime();

    // Queue (current) + history (finished), de-duplicated by Case ID + First Seen
    const all = {};
    const add = function(r) {
      const id = String(r[C.ID]).trim();
      if (!id) return;
      all[id + '|' + ((r[C.FIRST] instanceof Date) ? r[C.FIRST].getTime() : 0)] = r;
    };
    const q = ss.getSheetByName(EXPIRING_SHEET);
    if (q && q.getLastRow() > 1 && q.getRange(1, C.FIRST + 1).getValue() === EXPIRING_HEADERS[C.FIRST]) {
      q.getRange(2, 1, q.getLastRow() - 1, EXPIRING_HEADERS.length).getValues().forEach(add);
    }
    const h = ss.getSheetByName(EXPIRING_HISTORY_SHEET);
    if (h && h.getLastRow() > 1) {
      readSheetTail_(h, EXPIRING_HEADERS.length + 1, EXPIRING_HEADERS.length, startOfDay_(start)).values.forEach(add);
    }
    const cases = Object.keys(all).map(function(k) { return classifyExpiringCase_(all[k]); });

    const median = function(arr) {
      if (!arr.length) return null;
      const s = arr.slice().sort(function(a, b) { return a - b; });
      const m = Math.floor(s.length / 2);
      return s.length % 2 ? s[m] : (s[m - 1] + s[m]) / 2;
    };
    const r1 = function(x) { return x === null ? null : Math.round(x * 10) / 10; };
    const hrs = function(a, b) { return Math.max(0, (b - a) / 3600000); };

    const poolMap = {};
    const poolRow = function(p) {
      return poolMap[p] || (poolMap[p] = { pool: p, newCases: 0, saved: 0, missedLate: 0, missedOpen: 0, arrived: 0, open: 0, openNow: 0, openExpired: 0, clear: [] });
    };

    // Trend buckets (by First Seen): hourly for a 1-2 day range, otherwise daily
    const rangeDays = Math.round((endMs - startMs) / 86400000) + 1;
    const hourly = rangeDays <= 2;
    const sameDay = toDateStringFast(start) === toDateStringFast(end);
    const labels = [], bucketIdx = {};
    for (let t = startOfDay_(start).getTime(); t <= endMs; t += hourly ? 3600000 : 86400000) {
      bucketIdx[t] = labels.length;
      labels.push(Utilities.formatDate(new Date(t), tz, hourly ? (sameDay ? 'h a' : 'M/d h a') : 'MMM d'));
    }
    const zeros = function() { return labels.map(function() { return 0; }); };
    const trend = { granularity: hourly ? 'hourly' : 'daily', labels: labels, saved: zeros(), missed: zeros(), arrived: zeros(), open: zeros() };

    const out = { saved: 0, missedLate: 0, missedOpen: 0, arrived: 0, open: 0 };
    const clearHrs = [], workedHrs = [], assignHrs = [], runway = [];
    let resolved = 0, resolvedWithAgent = 0, newCases = 0;
    const missByHour = new Array(24).fill(0);

    cases.forEach(function(c) {
      if (c.firstSeen < startMs || c.firstSeen > endMs) return;
      newCases++;
      const p = poolRow(c.pool);
      p.newCases++;
      out[c.outcome]++;
      p[c.outcome]++;
      if (c.firstHours !== null && c.outcome !== 'arrived') runway.push(c.firstHours);

      if (c.exitAt) {
        resolved++;
        if (c.workedAt) resolvedWithAgent++;
        const ch = hrs(c.firstSeen, c.exitAt);
        clearHrs.push(ch);
        p.clear.push(ch);
      }
      if (c.workedAt) workedHrs.push(hrs(c.firstSeen, c.workedAt));
      if (c.leftAt) assignHrs.push(hrs(c.firstSeen, c.leftAt));

      const d = new Date(c.firstSeen);
      const key = hourly ? new Date(d.getFullYear(), d.getMonth(), d.getDate(), d.getHours()).getTime() : startOfDay_(d).getTime();
      const bi = bucketIdx[key];
      if (bi !== undefined) {
        if (c.outcome === 'saved') trend.saved[bi]++;
        else if (c.outcome === 'missedLate' || c.outcome === 'missedOpen') trend.missed[bi]++;
        else if (c.outcome === 'arrived') trend.arrived[bi]++;
        else trend.open[bi]++;
      }
      if ((c.outcome === 'missedLate' || c.outcome === 'missedOpen') && c.expiredAt) {
        missByHour[new Date(c.expiredAt).getHours()]++;
      }
    });

    // Live snapshot (independent of the date range)
    const live = { expired: 0, lt15: 0, lt24: 0, oldestExpiredHrs: null };
    cases.forEach(function(c) {
      if (c.state !== 'Open') return;
      const p = poolRow(c.pool);
      p.openNow++;
      if (c.type === EXP_TYPE_EXPIRED) {
        live.expired++; p.openExpired++;
        if (c.expiredAt) {
          const age = (nowMs - c.expiredAt) / 3600000;
          if (live.oldestExpiredHrs === null || age > live.oldestExpiredHrs) live.oldestExpiredHrs = age;
        }
      } else if (c.type === EXP_TYPE_15) live.lt15++;
      else live.lt24++;
    });
    live.oldestExpiredHrs = r1(live.oldestExpiredHrs);
    const props = PropertiesService.getScriptProperties();
    live.lastScanAt = Number(props.getProperty('EXP_LAST_SCAN_AT')) || 0;
    live.sourceChangedAt = Number(props.getProperty('EXP_SOURCE_CHANGED_AT')) || 0;

    // Agents credited with working these cases (by the time they logged them)
    const agentMap = {};
    cases.forEach(function(c) {
      if (!c.workedBy || !c.workedAt || c.workedAt < startMs || c.workedAt > endMs) return;
      const a = agentMap[c.workedBy] || (agentMap[c.workedBy] = { ldap: c.workedBy, count: 0, beforeExpiry: 0 });
      a.count++;
      if (c.outcome === 'saved') a.beforeExpiry++;
    });
    const agents = Object.keys(agentMap).map(function(k) { return agentMap[k]; })
      .sort(function(a, b) { return b.count - a.count; }).slice(0, 15);

    const decided = out.saved + out.missedLate + out.missedOpen;
    const pools = Object.keys(poolMap).map(function(k) {
      const p = poolMap[k];
      const dec = p.saved + p.missedLate + p.missedOpen;
      return {
        pool: p.pool, newCases: p.newCases, saved: p.saved, missed: p.missedLate + p.missedOpen, arrived: p.arrived,
        open: p.open, openNow: p.openNow, openExpired: p.openExpired,
        savedRate: dec ? Math.round(p.saved * 1000 / dec) / 10 : null,
        medianClearHrs: r1(median(p.clear))
      };
    }).sort(function(a, b) { return (b.missed - a.missed) || (b.newCases - a.newCases); });

    return {
      kpis: {
        newCases: newCases, saved: out.saved, missed: out.missedLate + out.missedOpen, missedOpen: out.missedOpen,
        arrived: out.arrived, openAtRisk: out.open,
        savedRate: decided ? Math.round(out.saved * 1000 / decided) / 10 : null,
        medianClearHrs: r1(median(clearHrs)), medianWorkedHrs: r1(median(workedHrs)), medianAssignHrs: r1(median(assignHrs)),
        medianRunwayHrs: r1(median(runway)),
        workedCoverage: resolved ? Math.round(resolvedWithAgent * 100 / resolved) : null
      },
      outcomes: out,
      pools: pools,
      trend: trend,
      missByHour: missByHour,
      agents: agents,
      live: live
    };
  } catch (e) {
    const user = Session.getActiveUser().getEmail() || 'Unknown';
    logError('getExpiringAnalytics', e.toString(), user);
    throw new Error('Unable to load expired/expiring analytics. Please try again.');
  }
}

// --- ARCHIVING ---
// Run setupArchiveTrigger() once from the editor to schedule archiveOldData() monthly.
// Rows older than ARCHIVE_KEEP_DAYS (Script Property, default 90) move from Raw_Cases and Index_CaseIDs
// into a separate archive spreadsheet (ID stored in the ARCHIVE_SPREADSHEET_ID Script Property; created on first run).
// Error_Logs older than 30 days are deleted.
// NOTE: Case IDs older than the keep window no longer count as "previously logged" in duplicate checks.

function setupArchiveTrigger() {
  requireManagerOrThrow();
  ScriptApp.getProjectTriggers().forEach(function(t) {
    if (t.getHandlerFunction() === 'archiveOldData') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('archiveOldData').timeBased().onMonthDay(1).atHour(3).create();
}

function archiveOldData(e) {
  if (!(e && e.triggerUid)) requireManagerOrThrow(); // the trigger runs unattended; manual runs must be a manager
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(30000);
    const keepDays = parseInt(PropertiesService.getScriptProperties().getProperty('ARCHIVE_KEEP_DAYS') || '90', 10);
    const cutoff = startOfDay_(new Date());
    cutoff.setDate(cutoff.getDate() - keepDays);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const archive = getArchiveSpreadsheet_(ss);
    const result = {
      rawCases: archiveRawCases_(ss, archive, cutoff),
      indexCaseIds: archiveLeadingRows_(ss.getSheetByName('Index_CaseIDs'), archive, 'Index_CaseIDs', 7, 2, cutoff),
      expiringHistory: archiveLeadingRows_(ss.getSheetByName('Expiring_History'), archive, 'Expiring_History', 18, 17, cutoff),
      errorLogs: pruneErrorLogs_(ss, 30)
    };

    const blank = archive.getSheetByName('Sheet1');
    if (blank && archive.getSheets().length > 1) archive.deleteSheet(blank);
    return result;
  } catch (err) {
    logError('archiveOldData', err.toString(), 'SYSTEM');
    throw err;
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function getArchiveSpreadsheet_(ss) {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('ARCHIVE_SPREADSHEET_ID');
  if (id) return SpreadsheetApp.openById(id); // throws if inaccessible; never silently creates a second archive
  const archive = SpreadsheetApp.create(ss.getName() + ' - Archive');
  props.setProperty('ARCHIVE_SPREADSHEET_ID', archive.getId());
  return archive;
}

// Raw_Cases rows are referenced by row number in Audit Queue col J, so: never archive past the oldest
// pending audit, and re-point every RawRowRef after the rows shift up.
function archiveRawCases_(ss, archive, cutoff) {
  const raw = ss.getSheetByName('Raw_Cases');
  const audit = ss.getSheetByName('Audit Queue');
  if (!raw || raw.getLastRow() <= 1) return 0;

  let maxRows;
  let auditData = [];
  if (audit && audit.getLastRow() > 1) {
    auditData = audit.getRange(2, 1, audit.getLastRow() - 1, 10).getValues();
    auditData.forEach(function(r) {
      const ref = parseInt(r[9], 10);
      if (String(r[0]).includes('PENDING') && ref >= 2) {
        const allowed = ref - 2; // rows 2..ref-1
        maxRows = (maxRows === undefined) ? allowed : Math.min(maxRows, allowed);
      }
    });
  }

  const moved = archiveLeadingRows_(raw, archive, 'Raw_Cases', 17, 1, cutoff, maxRows);
  if (moved > 0 && auditData.length > 0) {
    const refs = auditData.map(function(r) {
      const ref = parseInt(r[9], 10);
      if (!(ref >= 2)) return [r[9]];
      return [(ref - moved) >= 2 ? (ref - moved) : 'archived'];
    });
    audit.getRange(2, 10, refs.length, 1).setValues(refs);
  }
  return moved;
}

// Copies the leading rows older than cutoff into the archive, then deletes them. Returns rows moved.
function archiveLeadingRows_(sheet, archive, tabName, numCols, dateColIdx, cutoff, maxRows) {
  if (!sheet || sheet.getLastRow() <= 2) return 0;
  const last = sheet.getLastRow();
  const dates = sheet.getRange(2, dateColIdx + 1, last - 1, 1).getValues();

  let count = 0;
  while (count < dates.length) {
    const v = dates[count][0];
    const d = (v instanceof Date) ? v : new Date(v);
    if (isNaN(d.getTime()) || d >= cutoff) break;
    count++;
  }
  if (typeof maxRows === 'number') count = Math.min(count, maxRows);
  count = Math.min(count, last - 2); // never delete every data row
  if (count <= 0) return 0;

  let target = archive.getSheetByName(tabName);
  if (!target) {
    target = archive.insertSheet(tabName);
    target.getRange(1, 1, 1, numCols).setValues(sheet.getRange(1, 1, 1, numCols).getValues()).setFontWeight('bold');
    target.setFrozenRows(1);
  }

  const CHUNK = 5000;
  for (let off = 0; off < count; off += CHUNK) {
    const n = Math.min(CHUNK, count - off);
    const vals = sheet.getRange(2 + off, 1, n, numCols).getValues();
    const startRow = target.getLastRow() + 1;
    const missing = startRow + n - 1 - target.getMaxRows();
    if (missing > 0) target.insertRowsAfter(target.getMaxRows(), missing);
    target.getRange(startRow, 1, n, numCols).setValues(vals);
  }
  SpreadsheetApp.flush();
  sheet.deleteRows(2, count);
  return count;
}

function pruneErrorLogs_(ss, keepDays) {
  const sheet = ss.getSheetByName('Error_Logs');
  if (!sheet || sheet.getLastRow() <= 2) return 0;
  const cutoff = new Date();
  cutoff.setDate(cutoff.getDate() - keepDays);
  const last = sheet.getLastRow();
  const stamps = sheet.getRange(2, 1, last - 1, 1).getValues();
  let count = 0;
  while (count < stamps.length) {
    const d = stamps[count][0];
    if (!(d instanceof Date) || d >= cutoff) break;
    count++;
  }
  count = Math.min(count, last - 2);
  if (count > 0) sheet.deleteRows(2, count);
  return count;
}
