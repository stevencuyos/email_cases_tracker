const fs = require('fs');
let code = fs.readFileSync('/app/Index.html', 'utf8');

const modalHtml = `
  <!-- Agent Flagged Cases Modal -->
  <div id="agentFlaggedModal" class="hidden fixed inset-0 z-[60] flex items-center justify-center p-4">
    <div class="absolute inset-0 bg-black/40 backdrop-blur-sm" onclick="closeAgentFlaggedModal()"></div>
    <div class="relative bg-theme-sidebar rounded-theme shadow-2xl w-full max-w-4xl max-h-[85vh] flex flex-col overflow-hidden animate-fade-in">
      <div class="p-5 border-b border-gray-100 dark:border-gray-700 flex items-center justify-between">
        <h2 class="text-lg font-medium flex items-center gap-2 cell-strong">
          <span class="material-symbols-outlined text-red-500">error</span>
          Flagged Submissions: <span id="agentFlaggedName" class="font-bold"></span>
        </h2>
        <button onclick="closeAgentFlaggedModal()" class="text-gray-400 hover:text-gray-600 dark:hover:text-gray-300 transition focus:outline-none">
          <span class="material-symbols-outlined">close</span>
        </button>
      </div>
      <div class="p-5 overflow-y-auto flex-1 custom-scrollbar">
        <div class="flex items-center gap-2 mb-4">
          <button onclick="handleAgentFlaggedBulkAudit('Approve')" id="agentFlaggedBulkApproveBtn" disabled class="px-3 py-1.5 bg-green-50 text-green-700 dark:bg-green-900/30 dark:text-green-400 border border-green-200 dark:border-green-800/50 rounded text-sm font-medium hover:bg-green-100 dark:hover:bg-green-900/50 transition disabled:opacity-50 disabled:cursor-not-allowed flex items-center gap-1">
            <span class="material-symbols-outlined text-[16px]">check</span> Approve Selected
          </button>
          <button onclick="handleAgentFlaggedBulkAudit('Reject')" id="agentFlaggedBulkRejectBtn" disabled class="px-3 py-1.5 bg-red-50 text-red-700 dark:bg-red-900/30 dark:text-red-400 border border-red-200 dark:border-red-800/50 rounded text-sm font-medium hover:bg-red-100 dark:hover:bg-red-900/50 transition disabled:opacity-50 disabled:cursor-not-allowed flex items-center gap-1">
            <span class="material-symbols-outlined text-[16px]">close</span> Reject Selected
          </button>
        </div>
        <div class="overflow-x-auto">
          <table class="w-full text-left text-sm border-collapse">
            <thead>
              <tr class="bg-gray-50 dark:bg-gray-800/50 border-b border-gray-100 dark:border-gray-700">
                <th class="p-4 font-semibold text-gray-500 dark:text-gray-400 w-12"><input type="checkbox" id="selectAllAgentFlagged" class="w-4 h-4 text-theme-primary rounded border-gray-300 dark:border-gray-600 dark:bg-gray-700" onchange="toggleAllAgentFlaggedCheckboxes()"></th>
                <th class="p-4 font-semibold text-gray-500 dark:text-gray-400 w-24">Time</th>
                <th class="p-4 font-semibold text-gray-500 dark:text-gray-400 w-32">Type</th>
                <th class="p-4 font-semibold text-gray-500 dark:text-gray-400">Flagged IDs</th>
                <th class="p-4 font-semibold text-gray-500 dark:text-gray-400 w-1/4">Reason</th>
                <th class="p-4 font-semibold text-gray-500 dark:text-gray-400 text-right w-24">Action</th>
              </tr>
            </thead>
            <tbody id="agentFlaggedTableBody" class="divide-y divide-gray-50 dark:divide-gray-800/50">
              <!-- Rendered via JS -->
            </tbody>
          </table>
        </div>
      </div>
    </div>
  </div>
`;

if (!code.includes('id="agentFlaggedModal"')) {
  code = code.replace('<!-- Check-In Modal -->', modalHtml + '\n  <!-- Check-In Modal -->');
  fs.writeFileSync('/app/Index.html', code);
  console.log('Success');
} else {
  console.log('Modal already exists');
}
