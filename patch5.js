const fs = require('fs');
let code = fs.readFileSync('/app/Index.html', 'utf8');

const jsCode = `
    // --- AGENT FLAGGED MODAL LOGIC (Interval View) ---
    function openAgentFlaggedModal(ldap, auditsJsonStr) {
      document.getElementById('agentFlaggedName').textContent = ldap;
      const tbody = document.getElementById('agentFlaggedTableBody');
      const selectAll = document.getElementById('selectAllAgentFlagged');
      selectAll.checked = false;

      let audits = [];
      try {
        audits = JSON.parse(auditsJsonStr);
      } catch (e) {
        console.error("Failed to parse audits JSON", e);
      }

      if (audits.length === 0) {
        tbody.innerHTML = '<tr><td colspan="6" class="p-8 text-center text-gray-400">No pending audits found.</td></tr>';
        updateAgentFlaggedBulkButtons();
        document.getElementById('agentFlaggedModal').classList.remove('hidden');
        return;
      }

      tbody.innerHTML = audits.map(a => \`
        <tr class="row-hover transition" id="agent-flagged-row-\${a.row}">
          <td class="p-4">
            <input type="checkbox" class="agent-flagged-checkbox w-4 h-4 text-theme-primary rounded border-gray-300 dark:border-gray-600 dark:bg-gray-700" onchange="updateAgentFlaggedBulkButtons()" data-audit-row="\${a.row}" data-raw-row="\${a.rawRowRef}">
          </td>
          <td class="p-4">
            <span class="cell-strong">\${a.timestamp}</span>
          </td>
          <td class="p-4 text-xs cell-muted">\${a.caseType}</td>
          <td class="p-4 align-top">
            <div class="font-mono text-xs chip-box-alert rounded p-2 flex flex-wrap gap-2 leading-relaxed inline-flex items-center">
              \${formatCaseIdsAsLinks(a.flaggedIds, true)}
            </div>
          </td>
          <td class="p-4 text-xs \${String(a.reason).includes('minimum unmet') ? 'text-red-600 dark:text-red-400 font-medium' : 'cell-muted'}">\${a.reason}</td>
          <td class="p-4 text-right">
            <div class="flex items-center justify-end gap-2">
              <button onclick="handleAgentFlaggedAudit(\${a.row}, \${a.rawRowRef}, 'Approve')" class="text-green-600 bg-green-50 hover:bg-green-100 p-1.5 rounded transition" title="Approve">
                <span class="material-symbols-outlined text-[20px]">check</span>
              </button>
              <button onclick="handleAgentFlaggedAudit(\${a.row}, \${a.rawRowRef}, 'Reject')" class="text-red-600 bg-red-50 hover:bg-red-100 p-1.5 rounded transition" title="Reject">
                <span class="material-symbols-outlined text-[20px]">close</span>
              </button>
            </div>
          </td>
        </tr>
      \`).join('');

      updateAgentFlaggedBulkButtons();
      document.getElementById('agentFlaggedModal').classList.remove('hidden');
    }

    function closeAgentFlaggedModal() {
      document.getElementById('agentFlaggedModal').classList.add('hidden');
    }

    function toggleAllAgentFlaggedCheckboxes() {
      const selectAll = document.getElementById('selectAllAgentFlagged');
      const checkboxes = document.querySelectorAll('.agent-flagged-checkbox');
      checkboxes.forEach(cb => cb.checked = selectAll.checked);
      updateAgentFlaggedBulkButtons();
    }

    function updateAgentFlaggedBulkButtons() {
      const checkedBoxes = document.querySelectorAll('.agent-flagged-checkbox:checked');
      const btnApprove = document.getElementById('agentFlaggedBulkApproveBtn');
      const btnReject = document.getElementById('agentFlaggedBulkRejectBtn');

      const hasChecked = checkedBoxes.length > 0;
      btnApprove.disabled = !hasChecked;
      btnReject.disabled = !hasChecked;

      const allCheckboxes = document.querySelectorAll('.agent-flagged-checkbox');
      const selectAllCb = document.getElementById('selectAllAgentFlagged');
      if (selectAllCb && allCheckboxes.length > 0) {
        selectAllCb.checked = allCheckboxes.length === checkedBoxes.length;
      }
    }

    function handleAgentFlaggedAudit(auditRow, rawRowRef, resolution) {
      const executeAudit = () => {
        const rowEl = document.getElementById(\`agent-flagged-row-\${auditRow}\`);
        if (rowEl) rowEl.style.opacity = '0.5';

        google.script.run
          .withSuccessHandler(res => {
            if (res.success) {
              showToast(\`Successfully \${resolution.toLowerCase()}d cases.\`, 'success');
              // Reload interval data to refresh the view and close the modal implicitly when re-rendered, or just reload data.
              loadIntervalData();
              closeAgentFlaggedModal();
            } else {
              showToast(res.error, 'error');
              if (rowEl) rowEl.style.opacity = '1';
            }
          })
          .withFailureHandler(err => {
            showToast("Network error.", 'error');
            if (rowEl) rowEl.style.opacity = '1';
          })
          .resolveAudit(auditRow, rawRowRef, resolution);
      };

      if (resolution === 'Reject') {
        showConfirmModal(
          'Confirm Rejection',
          'Reject this flagged submission as duplicate/fraud? <br><br><b>This can\\'t be undone from this screen.</b>',
          'Reject Case',
          'bg-red-600 hover:bg-red-700',
          'warning',
          'text-red-500',
          executeAudit
        );
      } else {
        executeAudit();
      }
    }

    function handleAgentFlaggedBulkAudit(resolution) {
      const checkedBoxes = document.querySelectorAll('.agent-flagged-checkbox:checked');
      if (checkedBoxes.length === 0) return;

      const executeBulkAudit = () => {
        const auditsToProcess = Array.from(checkedBoxes).map(cb => ({
          auditRow: parseInt(cb.getAttribute('data-audit-row'), 10),
          rawRowRef: parseInt(cb.getAttribute('data-raw-row'), 10)
        }));

        auditsToProcess.forEach(audit => {
          const rowEl = document.getElementById(\`agent-flagged-row-\${audit.auditRow}\`);
          if (rowEl) rowEl.style.opacity = '0.5';
        });

        document.getElementById('agentFlaggedBulkApproveBtn').disabled = true;
        document.getElementById('agentFlaggedBulkRejectBtn').disabled = true;

        google.script.run
          .withSuccessHandler(res => {
            if (res.success) {
              showToast(\`Successfully \${resolution.toLowerCase()}d \${auditsToProcess.length} cases.\`, 'success');
              loadIntervalData();
              closeAgentFlaggedModal();
            } else {
              showToast(res.error, 'error');
              updateAgentFlaggedBulkButtons(); // re-enable buttons
              auditsToProcess.forEach(audit => {
                const rowEl = document.getElementById(\`agent-flagged-row-\${audit.auditRow}\`);
                if (rowEl) rowEl.style.opacity = '1';
              });
            }
          })
          .withFailureHandler(err => {
            showToast("Network error.", 'error');
            updateAgentFlaggedBulkButtons();
            auditsToProcess.forEach(audit => {
              const rowEl = document.getElementById(\`agent-flagged-row-\${audit.auditRow}\`);
              if (rowEl) rowEl.style.opacity = '1';
            });
          })
          .resolveAuditsBulk(auditsToProcess, resolution);
      };

      if (resolution === 'Reject') {
        showConfirmModal(
          'Confirm Bulk Rejection',
          \`Reject <b>\${checkedBoxes.length}</b> flagged submission(s) as duplicate/fraud? <br><br><b>This can\\'t be undone from this screen.</b>\`,
          'Reject Cases',
          'bg-red-600 hover:bg-red-700',
          'warning',
          'text-red-500',
          executeBulkAudit
        );
      } else if (resolution === 'Approve' && checkedBoxes.length > 1) {
        showConfirmModal(
          'Confirm Bulk Approval',
          \`Approve <b>\${checkedBoxes.length}</b> flagged submissions? <br><br>This will count them as valid across \${checkedBoxes.length} case(s).\`,
          'Approve Cases',
          'bg-green-600 hover:bg-green-700',
          'check_circle',
          'text-green-500',
          executeBulkAudit
        );
      } else {
        executeBulkAudit();
      }
    }

`;

if (!code.includes('openAgentFlaggedModal')) {
  code = code.replace('// --- DASHBOARD LOGIC ---', jsCode + '\n    // --- DASHBOARD LOGIC ---');
  fs.writeFileSync('/app/Index.html', code);
  console.log('Success');
} else {
  console.log('Functions already exist');
}
