// 1. Serve the Web App Interface
function doGet() {
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Case Tracking Portal')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 2. Auto-Initialize the Database Structure
function initializeDatabase() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const sheetsConfig = {
    // Added Name, LOB, and Workflow to the database
    'Raw_Cases': ['Timestamp', 'Date', 'Interval', 'Agent', 'Name', 'Site', 'LOB', 'Workflow', 'Shift Type', 'Case Type', 'Total', 'Valid', 'Flagged', 'Case IDs', 'Audit Notes'],
    'Index_CaseIDs': ['Case ID', 'Type Logged', 'Date Logged', 'Agent'],
    'Audit Queue': ['Status', 'Timestamp', 'Agent', 'Site', 'Case Type', 'Total Logged', 'Flagged IDs', 'Audit Reason', 'Resolution', 'RawRowRef']
  };

  for (const [sheetName, headers] of Object.entries(sheetsConfig)) {
    let sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight("bold").setBackground("#f3f3f3");
      sheet.setFrozenRows(1);
      console.log(`Created missing tab: ${sheetName}`);
    }
  }
}

// 3. Fetch User Profile from Masterlist Tab
function getUserProfile() {
  const email = Session.getActiveUser().getEmail();
  const currentLdap = email ? email.split('@')[0] : 'unknown_agent';
  
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
  } catch (e) {
    console.error("Error fetching user profile: " + e);
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
    const caseType = formObject.caseType;
    const rawText = formObject.caseIdsText;

    let rawIds = rawText.split(/[\n,;\s]+/).map(id => id.trim()).filter(id => id !== '');
    let uniqueIds = [...new Set(rawIds)];
    let validIds = [];
    let flaggedIds = [];
    let auditReasons = [];

    const indexData = indexSheet.getLastRow() > 1 ? indexSheet.getRange(2, 1, indexSheet.getLastRow() - 1, 4).getValues() : [];

    uniqueIds.forEach(id => {
      let isFlagged = false;
      let reason = "";
      let history = indexData.filter(row => row[0] == id);

      // RULE 1: Reopened Cases have their own specific criteria
      if (caseType === 'Reopened Cases') {
        let todayHistory = history.filter(row => row[2] == dateStr && row[1] === 'Reopened Cases' && row[3] === ldap);
        if (todayHistory.length > 0) {
          isFlagged = true;
          reason = "Already reopened today by this agent";
        }
      } 
      // RULE 2: Regular Email, Manual Assignment, Telus, and Cimba are bound to Global Uniqueness
      else {
        if (history.length > 0) {
          isFlagged = true;
          reason = `Case ID already logged previously under type: [${history[0][1]}]`;
        }
      }

      if (isFlagged) {
        flaggedIds.push(id);
        auditReasons.push(`${id}: ${reason}`);
      } else {
        validIds.push(id);
      }
    });

    let totalCount = uniqueIds.length;
    let validCount = validIds.length;
    let flaggedCount = flaggedIds.length;
    let auditNotes = flaggedCount > 0 ? "⚠️ " + auditReasons.join(' | ') : "Clean";

    let rawRowNumber = rawSheet.getLastRow() + 1;
    // Writes LOB, Workflow, and Name into the raw sheet
    rawSheet.appendRow([timestamp, dateStr, intervalStr, ldap, userProfile.name, site, userProfile.lob, userProfile.workflow, shiftType, caseType, totalCount, validCount, flaggedCount, uniqueIds.join(', '), auditNotes]);

    let indexDataToAppend = uniqueIds.map(id => [id, caseType, dateStr, ldap]);
    if (indexDataToAppend.length > 0) {
      indexSheet.getRange(indexSheet.getLastRow() + 1, 1, indexDataToAppend.length, 4).setValues(indexDataToAppend);
    }

    if (flaggedCount > 0) {
      auditSheet.appendRow(["🔴 PENDING", timestamp, ldap, site, caseType, totalCount, flaggedIds.join(', '), auditNotes, "", rawRowNumber]);
    }
    
    return { success: true, valid: validCount, flagged: flaggedCount };
    
  } catch (error) {
    return { success: false, error: error.toString() };
  } finally {
    lock.releaseLock();
  }
}

// 5. Fetch Dashboard Data
function getDashboardData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const auditSheet = ss.getSheetByName('Audit Queue');
  const rawSheet = ss.getSheetByName('Raw_Cases');

  // A. Get Pending Audits
  let audits = [];
  if (auditSheet.getLastRow() > 1) {
    // FIXED: Now fetching 10 columns instead of 9 to grab the shifted RawRowRef
    const auditData = auditSheet.getRange(2, 1, auditSheet.getLastRow() - 1, 10).getValues();
    audits = auditData.map((r, i) => ({
      row: i + 2, 
      status: r[0],
      timestamp: Utilities.formatDate(new Date(r[1]), Session.getScriptTimeZone(), "h:mm a"),
      agent: r[2],
      site: r[3],
      caseType: r[4],
      totalLogged: r[5],
      flaggedIds: r[6],
      reason: r[7],
      rawRowRef: r[9] // FIXED: Shifted from index 8 to index 9 (Column J)
    })).filter(a => a.status.includes('PENDING'));
  }

  // B. Get Today's Metrics (Grouped by Agent)
  let metrics = {};
  if (rawSheet.getLastRow() > 1) {
    const todayStr = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "M/d/yyyy");
    const rawData = rawSheet.getRange(2, 1, rawSheet.getLastRow() - 1, 15).getValues(); 

    rawData.forEach(r => {
      // FIXED: Safely convert Google Sheets Date objects back to text for matching
      let rowDateStr = "";
      if (r[1] instanceof Date) {
        rowDateStr = Utilities.formatDate(r[1], Session.getScriptTimeZone(), "M/d/yyyy");
      } else {
        rowDateStr = String(r[1]);
      }

      if (rowDateStr === todayStr) { 
        const agent = r[3];
        const site = r[5]; 
        const valid = Number(r[11]) || 0; 
        const flagged = Number(r[12]) || 0; 
        
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
}

// 6. Resolve Soft Audits (Approve/Reject)
function resolveAudit(auditRow, rawRowRef, resolution) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const auditSheet = ss.getSheetByName('Audit Queue');
    const rawSheet = ss.getSheetByName('Raw_Cases');

    if (resolution === 'Approve') {
      auditSheet.getRange(auditRow, 1).setValue("🟢 APPROVED");
      auditSheet.getRange(auditRow, 9).setValue("Approved");

      // Fetch Valid (Col 12) and Flagged (Col 13)
      const validCount = rawSheet.getRange(rawRowRef, 12).getValue();
      const flaggedCount = rawSheet.getRange(rawRowRef, 13).getValue();

      rawSheet.getRange(rawRowRef, 12).setValue(validCount + flaggedCount); // Update Valid
      rawSheet.getRange(rawRowRef, 13).setValue(0); // Zero out Flagged
      rawSheet.getRange(rawRowRef, 15).setValue("✅ Resolved by Manager"); // Audit Notes

    } else {
      auditSheet.getRange(auditRow, 1).setValue("⚫ REJECTED");
      auditSheet.getRange(auditRow, 9).setValue("Rejected");
      
      // FIXED: We must zero out the flagged count on Reject so it drops off the pending metrics!
      rawSheet.getRange(rawRowRef, 13).setValue(0); 
      
      rawSheet.getRange(rawRowRef, 15).setValue("❌ Rejected by Manager (Duplicate/Fraud)");
    }
    return { success: true };
  } catch (e) {
    return { success: false, error: e.toString() };
  } finally {
    lock.releaseLock();
  }
}
// 7. Fetch Data for Interval View
function getIntervalData(dateStr, intervalHourStr) {
  // dateStr format expected: "9/15/2026"
  // intervalHourStr format expected: "16:00" (24-hour format string)
  
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const shiftSheet = ss.getSheetByName('Agent Shifts');
  const rawSheet = ss.getSheetByName('Raw_Cases');
  const masterSheet = ss.getSheetByName('Masterlist'); 
  
  let targetHour = parseInt(intervalHourStr.split(':')[0], 10);
  let agentsInInterval = {};
  
  // NEW: 1. Build a whitelist of "Email" Channel Agents from the Masterlist
  let emailAgents = new Set();
  if (masterSheet && masterSheet.getLastRow() > 1) {
    // Fetch columns A through X (index 0 to 23). Col X is Channel.
    const masterData = masterSheet.getRange(2, 1, masterSheet.getLastRow() - 1, 24).getValues();
    for (let i = 0; i < masterData.length; i++) {
      const ldap = masterData[i][0] ? masterData[i][0].toString().toLowerCase() : '';
      const channel = masterData[i][23] ? masterData[i][23].toString().toLowerCase() : ''; 
      
      if (channel === 'email') {
        emailAgents.add(ldap);
      }
    }
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
          formattedCellDate = Utilities.formatDate(cellDate, Session.getScriptTimeZone(), "M/d/yyyy");
        } else {
           formattedCellDate = Utilities.formatDate(new Date(cellDate), Session.getScriptTimeZone(), "M/d/yyyy");
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
        const ldap = shiftData[r][0] ? shiftData[r][0].toString() : '';
        
        // NEW FILTER: Skip agent if they are not in the Email channel (unless they do OT later)
        if (!emailAgents.has(ldap.toLowerCase())) {
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
                casesLogged: 0
              };
            }
          }
        }
      }
    }
  }
  
  // 3. Process Overtime & Live Metrics from 'Raw_Cases'
  if (rawSheet && rawSheet.getLastRow() > 1) {
    const rawData = rawSheet.getRange(2, 1, rawSheet.getLastRow() - 1, 15).getValues(); 
    // Format target hour to match Raw_Cases "h:00 a" format (e.g. "4:00 PM")
    let targetDateObj = new Date();
    targetDateObj.setHours(targetHour, 0, 0, 0);
    const intervalLabel = Utilities.formatDate(targetDateObj, Session.getScriptTimeZone(), "h:00 a");
    
    rawData.forEach(r => {
      let rowDateStr = (r[1] instanceof Date) ? Utilities.formatDate(r[1], Session.getScriptTimeZone(), "M/d/yyyy") : String(r[1]);
      
      if (rowDateStr === dateStr && r[2] === intervalLabel) {
        const ldap = r[3];
        const site = r[5];
        const isOvertime = r[8] === "Overtime";
        const validCases = Number(r[11]) || 0;
        
        // If OT agent isn't on the shift list, add them dynamically regardless of their Channel
        if (!agentsInInterval[ldap] && isOvertime) {
          agentsInInterval[ldap] = {
            ldap: ldap,
            sos: "OT",
            eos: "OT",
            site: site,
            isOT: true,
            casesLogged: 0
          };
        }
        
        // Add metrics if they are in the list
        if (agentsInInterval[ldap]) {
          agentsInInterval[ldap].casesLogged += validCases;
        }
      }
    });
  }
  
  return Object.values(agentsInInterval).sort((a, b) => a.ldap.localeCompare(b.ldap));
}
