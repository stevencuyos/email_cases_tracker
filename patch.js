const fs = require('fs');
let code = fs.readFileSync('/app/Code.gs', 'utf8');

const search = `    // 4. Compute Status
    const results = Object.values(agentsInInterval).sort((a, b) => a.ldap.localeCompare(b.ldap));

    results.forEach(agent => {
      let computedStatus = "";`;

const replace = `    // 3.5 Process Pending Audits for this Interval
    const auditSheet = ss.getSheetByName('Audit Queue');
    let pendingAuditsByLdap = {};
    if (auditSheet && auditSheet.getLastRow() > 1) {
      const auditData = auditSheet.getRange(2, 1, auditSheet.getLastRow() - 1, 10).getValues();
      let targetDateObj = new Date();
      targetDateObj.setHours(targetHour, 0, 0, 0);
      const targetIntervalLabel = Utilities.formatDate(targetDateObj, Session.getScriptTimeZone(), "h:00 a");

      auditData.forEach((r, i) => {
        const status = r[0];
        if (String(status).includes('PENDING')) {
          const timestamp = (r[1] instanceof Date) ? Utilities.formatDate(r[1], Session.getScriptTimeZone(), "h:mm a") : String(r[1]);
          const auditDateObj = (r[1] instanceof Date) ? r[1] : new Date(r[1]);
          const auditDateStr = toDateStringFast(auditDateObj);

          const auditInterval = Utilities.formatDate(auditDateObj, Session.getScriptTimeZone(), "h:00 a");

          if (auditDateStr === dateStr && auditInterval === targetIntervalLabel) {
            const agent = String(r[2]).trim().toLowerCase();
            if (!pendingAuditsByLdap[agent]) {
              pendingAuditsByLdap[agent] = [];
            }
            pendingAuditsByLdap[agent].push({
              row: i + 2,
              status: status,
              timestamp: timestamp,
              rawTs: auditDateObj.getTime(),
              agent: r[2],
              site: r[3],
              caseType: r[4],
              totalLogged: r[5],
              flaggedIds: r[6],
              reason: r[7],
              rawRowRef: r[9]
            });
          }
        }
      });
    }

    // 4. Compute Status
    const results = Object.values(agentsInInterval).sort((a, b) => a.ldap.localeCompare(b.ldap));

    results.forEach(agent => {
      agent.pendingAudits = pendingAuditsByLdap[agent.ldap] || [];
      let computedStatus = "";`;

if (code.includes(search)) {
  code = code.replace(search, replace);
  fs.writeFileSync('/app/Code.gs', code);
  console.log('Success');
} else {
  console.log('Search block not found!');
}
