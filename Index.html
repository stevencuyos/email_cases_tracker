<!DOCTYPE html>
<html lang="en">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Case Tracking Portal</title>
  
  <!-- Google Fonts & Icons -->
  <link href="https://fonts.googleapis.com/css2?family=Roboto:wght@300;400;500;700&display=swap" rel="stylesheet">
  <link href="https://fonts.googleapis.com/css2?family=Material+Symbols+Outlined:opsz,wght,FILL,GRAD@20..48,100..700,0..1,-50..200" rel="stylesheet" />
  
  <!-- Tailwind CSS -->
  <script src="https://cdn.tailwindcss.com"></script>
  
  <style>
    body { font-family: 'Roboto', sans-serif; background-color: #f8f9fa; }
    .google-blue { background-color: #1a73e8; }
    .google-blue-text { color: #1a73e8; }
    .nav-active { border-bottom: 3px solid #1a73e8; color: #1a73e8; font-weight: 500; }
    .nav-inactive { color: #5f6368; font-weight: 400; }
    .nav-inactive:hover { background-color: #f1f3f4; }
    textarea::-webkit-scrollbar { width: 8px; }
    textarea::-webkit-scrollbar-thumb { background-color: #dadce0; border-radius: 4px; }
    
    /* NEW: Smooth reveal animation for the Manager Nav */
    @keyframes fadeSlideDown {
      0% { opacity: 0; transform: translateY(-10px); }
      100% { opacity: 1; transform: translateY(0); }
    }
    .manager-reveal {
      animation: fadeSlideDown 0.5s ease-out forwards;
    }
  </style>
</head>
<body class="text-gray-800 antialiased min-h-screen flex flex-col items-center">

  <!-- Global Header -->
  <header class="w-full bg-white shadow-sm border-b border-gray-200 sticky top-0 z-40">
    <div class="h-1 w-full bg-[#fbbc04]"></div> <!-- Google Yellow Accent -->
    <div class="max-w-5xl mx-auto px-6 h-16 flex items-center justify-between">
      
      <div class="flex items-center gap-2">
        <span class="material-symbols-outlined text-blue-600 text-3xl">track_changes</span>
        <h1 class="text-xl font-normal text-gray-700">Case Tracker</h1>
      </div>

      <!-- Identity Chip -->
      <div id="identityChip" class="flex items-center gap-3 bg-gray-50 px-3 py-1.5 rounded-full border border-gray-200 animate-pulse">
        <span class="text-sm text-gray-500">Loading profile...</span>
      </div>

    </div>

    <!-- Shared/Manager Navigation -->
    <div id="managerNav" class="max-w-5xl mx-auto px-6 flex gap-2">
      <button onclick="switchTab('submissionView')" id="tab-submissionView" class="px-4 py-3 text-sm nav-active transition-colors">
        Submit Cases
      </button>
      <button onclick="switchTab('mySubmissionsView')" id="tab-mySubmissionsView" class="px-4 py-3 text-sm nav-inactive rounded-t-md transition-colors flex items-center gap-1">
        My Submissions
      </button>
      <div id="managerOnlyTabs" class="flex gap-2 hidden">
        <button onclick="switchTab('dashboardView')" id="tab-dashboardView" class="px-4 py-3 text-sm nav-inactive rounded-t-md transition-colors flex items-center gap-1">
          Manager Dashboard
        </button>
        <button onclick="switchTab('intervalView')" id="tab-intervalView" class="px-4 py-3 text-sm nav-inactive rounded-t-md transition-colors flex items-center gap-1">
          Interval View
        </button>
      </div>
    </div>
  </header>

  <!-- Main Content Area -->
  <main class="w-full px-4 py-8 relative flex flex-col items-center">

    <!-- VIEW 1: Submission Form -->
    <div id="submissionView" class="w-full max-w-2xl bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden transition-all duration-300">
      <div class="p-8 pb-6 border-b border-gray-100">
        <h2 class="text-2xl font-normal mb-1">Log Your Interval</h2>
        <p class="text-sm text-gray-500">Paste your Case IDs separated by commas or new lines. Duplicates will be automatically audited.</p>
      </div>

      <form id="trackerForm" onsubmit="handleFormSubmit(event)" class="p-8 pt-6 space-y-8">
        
        <!-- Shift Type -->
        <div>
          <h3 class="text-sm font-medium mb-3 text-gray-700">Is this Regular Shift or OT? <span class="text-red-500">*</span></h3>
          <div class="space-y-3">
            <label class="flex items-center gap-3 cursor-pointer">
              <input type="radio" name="shiftType" value="Regular Shift" required class="w-4 h-4 text-blue-600 focus:ring-blue-500 border-gray-300">
              <span class="text-sm text-gray-800">Regular Shift</span>
            </label>
            <label class="flex items-center gap-3 cursor-pointer">
              <input type="radio" name="shiftType" value="Overtime" class="w-4 h-4 text-blue-600 focus:ring-blue-500 border-gray-300">
              <span class="text-sm text-gray-800">Overtime</span>
            </label>
          </div>
        </div>

        <!-- Interval Activity -->
        <div>
          <h3 class="text-sm font-medium mb-3 text-gray-700">Activity for this interval? <span class="text-red-500">*</span></h3>
          <div class="space-y-3">
            <label class="flex items-center gap-3 cursor-pointer">
              <input type="radio" name="intervalActivity" value="Normal Production" required checked onchange="handleActivityChange()" class="w-4 h-4 text-blue-600 focus:ring-blue-500 border-gray-300">
              <span class="text-sm text-gray-800">Normal Production</span>
            </label>
            <label class="flex items-center gap-3 cursor-pointer">
              <input type="radio" name="intervalActivity" value="Break/Lunch" onchange="handleActivityChange()" class="w-4 h-4 text-blue-600 focus:ring-blue-500 border-gray-300">
              <span class="text-sm text-gray-800">Break/Lunch</span>
            </label>
            <label class="flex items-center gap-3 cursor-pointer">
              <input type="radio" name="intervalActivity" value="Coaching/Training" onchange="handleActivityChange()" class="w-4 h-4 text-blue-600 focus:ring-blue-500 border-gray-300">
              <span class="text-sm text-gray-800">Coaching/Training</span>
            </label>
          </div>
        </div>

        <!-- Case Type -->
        <div>
          <h3 class="text-sm font-medium mb-3 text-gray-700">Email Cases type <span class="text-red-500">*</span></h3>
          <div class="grid grid-cols-1 sm:grid-cols-2 gap-3">
            <label class="flex items-center gap-3 cursor-pointer">
              <input type="radio" name="caseType" value="Regular Email (Take Next)" required class="w-4 h-4 text-blue-600">
              <span class="text-sm text-gray-800">Regular (Take Next)</span>
            </label>
            <label class="flex items-center gap-3 cursor-pointer">
              <input type="radio" name="caseType" value="Reopened Cases" class="w-4 h-4 text-blue-600">
              <span class="text-sm text-gray-800">Reopened Cases</span>
            </label>
            <label class="flex items-center gap-3 cursor-pointer">
              <input type="radio" name="caseType" value="Telus Cases" class="w-4 h-4 text-blue-600">
              <span class="text-sm text-gray-800">Telus Cases</span>
            </label>
            <label class="flex items-center gap-3 cursor-pointer">
              <input type="radio" name="caseType" value="Manual Assignment" class="w-4 h-4 text-blue-600">
              <span class="text-sm text-gray-800">Manual Assignment</span>
            </label>
            <label class="flex items-center gap-3 cursor-pointer">
              <input type="radio" name="caseType" value="Cimba Cases" class="w-4 h-4 text-blue-600">
              <span class="text-sm text-gray-800">Cimba Cases</span>
            </label>
          </div>
        </div>

        <!-- Bulk Case IDs -->
        <div>
          <h3 class="text-sm font-medium mb-2 text-gray-700">Tracked Case IDs <span class="text-red-500">*</span></h3>
          <textarea id="caseIds" rows="5" required placeholder="Example:&#10;1-12345678&#10;1-87654321" class="w-full p-3 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent text-sm resize-y shadow-inner"></textarea>
        </div>

        <!-- Action Footer -->
        <div class="flex items-center justify-between pt-4 border-t border-gray-100">
          <button type="button" onclick="document.getElementById('trackerForm').reset()" class="text-sm font-medium text-gray-500 hover:text-gray-800 px-4 py-2 rounded transition">
            Clear form
          </button>
          <button type="submit" id="submitBtn" class="google-blue text-white font-medium text-sm px-6 py-2.5 rounded shadow hover:bg-[#1557b0] transition flex items-center gap-2">
            <span>Submit Cases</span>
            <span id="spinner" class="material-symbols-outlined animate-spin hidden text-[20px]">progress_activity</span>
          </button>
        </div>
      </form>
    </div>

    <!-- VIEW: My Submissions / Agent Submissions -->
    <div id="mySubmissionsView" class="hidden w-full max-w-5xl space-y-6 transition-all duration-300 pb-12">
      <div class="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden">
        <div class="p-6 border-b border-gray-100 flex justify-between items-center bg-gray-50/50">
          <div>
            <h2 id="mySubmissionsTitle" class="text-xl font-normal text-gray-800 flex items-center gap-2">
              <span class="material-symbols-outlined text-blue-500">list_alt</span>
              My Submissions
            </h2>
            <p id="mySubmissionsDesc" class="text-sm text-gray-500 mt-1">View your tracked cases per interval for the selected date.</p>
          </div>
          <div class="flex items-center gap-4">
            <div id="managerLdapSearch" class="hidden relative">
              <input type="text" id="targetAgentLdap" list="agentList" onchange="loadMySubmissions()" placeholder="Agent LDAP" class="p-2 border border-gray-300 rounded text-sm focus:ring-blue-500 w-36">
              <datalist id="agentList"></datalist>
            </div>
            <input type="date" id="mySubmissionsDate" onchange="loadMySubmissions()" class="p-2 border border-gray-300 rounded text-sm focus:ring-blue-500">
            <button onclick="loadMySubmissions()" class="text-gray-500 hover:text-blue-600 transition" title="Refresh Data">
              <span class="material-symbols-outlined">refresh</span>
            </button>
          </div>
        </div>
        <div class="overflow-x-auto">
          <table class="w-full text-left border-collapse">
            <thead>
              <tr class="bg-gray-50 text-gray-500 text-xs uppercase tracking-wider">
                <th class="p-4 font-medium border-b border-gray-200">Interval</th>
                <th class="p-4 font-medium border-b border-gray-200">Activity</th>
                <th class="p-4 font-medium border-b border-gray-200">Case Type</th>
                <th class="p-4 font-medium border-b border-gray-200 text-center">Valid Count</th>
                <th class="p-4 font-medium border-b border-gray-200 min-w-[250px]">Case IDs</th>
              </tr>
            </thead>
            <tbody id="mySubmissionsTableBody" class="text-sm divide-y divide-gray-100">
              <tr><td colspan="5" class="p-8 text-center text-gray-400">Loading submissions...</td></tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>

    <!-- VIEW 2: Manager Dashboard -->
    <div id="dashboardView" class="hidden w-full max-w-5xl space-y-6 transition-all duration-300 pb-12">
      
      <!-- Section 1: Audit Queue -->
      <div class="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden">
        <div class="p-6 border-b border-gray-100 flex justify-between items-center bg-red-50/30">
          <div>
            <h2 class="text-xl font-normal text-gray-800 flex items-center gap-2">
              <span class="material-symbols-outlined text-red-500">gavel</span>
              Action Required: Audit Queue
            </h2>
            <p class="text-sm text-gray-500 mt-1">Review flagged duplicate submissions.</p>
          </div>
          <button onclick="loadDashboard()" class="text-gray-500 hover:text-blue-600 transition" title="Refresh Data">
            <span class="material-symbols-outlined">refresh</span>
          </button>
        </div>
        <div class="overflow-x-auto">
          <table class="w-full text-left border-collapse">
            <thead>
              <tr class="bg-gray-50 text-gray-500 text-xs uppercase tracking-wider">
                <th class="p-4 font-medium border-b border-gray-200">Agent</th>
                <th class="p-4 font-medium border-b border-gray-200">Time / Type</th>
                <th class="p-4 font-medium border-b border-gray-200 min-w-[180px]">Flagged IDs</th>
                <th class="p-4 font-medium border-b border-gray-200">Reason</th>
                <th class="p-4 font-medium border-b border-gray-200 text-right">Action</th>
              </tr>
            </thead>
            <tbody id="auditTableBody" class="text-sm divide-y divide-gray-100">
              <tr><td colspan="5" class="p-8 text-center text-gray-400">Loading audits...</td></tr>
            </tbody>
          </table>
        </div>
      </div>

      <!-- Section 2: Today's Metrics -->
      <div class="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden">
        <div class="p-6 border-b border-gray-100">
          <h2 class="text-xl font-normal text-gray-800 flex items-center gap-2">
            <span class="material-symbols-outlined text-green-600">bar_chart</span>
            Today's Live Metrics
          </h2>
          <p class="text-sm text-gray-500 mt-1">Valid tracked cases per agent for the current day.</p>
        </div>
        <div class="overflow-x-auto">
          <table class="w-full text-left border-collapse">
            <thead>
              <tr class="bg-gray-50 text-gray-500 text-xs uppercase tracking-wider">
                <th class="p-4 font-medium border-b border-gray-200">Agent</th>
                <th class="p-4 font-medium border-b border-gray-200 text-center">Valid Cases</th>
                <th class="p-4 font-medium border-b border-gray-200 text-center">Pending Flags</th>
              </tr>
            </thead>
            <tbody id="metricsTableBody" class="text-sm divide-y divide-gray-100">
            </tbody>
          </table>
        </div>
      </div>
    </div>

    <!-- VIEW 3: Interval Management -->
    <div id="intervalView" class="hidden w-full max-w-5xl space-y-6 transition-all duration-300 pb-12">
      <div class="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden">
        
        <!-- Controls Header -->
        <div class="p-6 border-b border-gray-100 flex flex-wrap gap-4 items-end bg-blue-50/30">
          <div>
            <label class="block text-xs font-medium text-gray-500 mb-1">Select Date</label>
            <input type="date" id="intervalDate" class="p-2 border border-gray-300 rounded text-sm focus:ring-blue-500">
          </div>
          <div>
            <label class="block text-xs font-medium text-gray-500 mb-1">Select Interval Hour</label>
            <select id="intervalHour" class="p-2 border border-gray-300 rounded text-sm focus:ring-blue-500 min-w-[120px]">
              <option value="0:00">12:00 AM</option>
              <option value="1:00">1:00 AM</option>
              <option value="2:00">2:00 AM</option>
              <option value="3:00">3:00 AM</option>
              <option value="4:00">4:00 AM</option>
              <option value="5:00">5:00 AM</option>
              <option value="6:00">6:00 AM</option>
              <option value="7:00">7:00 AM</option>
              <option value="8:00">8:00 AM</option>
              <option value="9:00">9:00 AM</option>
              <option value="10:00">10:00 AM</option>
              <option value="11:00">11:00 AM</option>
              <option value="12:00">12:00 PM</option>
              <option value="13:00">1:00 PM</option>
              <option value="14:00">2:00 PM</option>
              <option value="15:00">3:00 PM</option>
              <option value="16:00" selected>4:00 PM</option>
              <option value="17:00">5:00 PM</option>
              <option value="18:00">6:00 PM</option>
              <option value="19:00">7:00 PM</option>
              <option value="20:00">8:00 PM</option>
              <option value="21:00">9:00 PM</option>
              <option value="22:00">10:00 PM</option>
              <option value="23:00">11:00 PM</option>
            </select>
          </div>
          <div class="flex items-center gap-2">
            <button onclick="loadIntervalData()" class="bg-white border border-gray-300 text-gray-700 font-medium text-sm px-4 py-2 rounded shadow-sm hover:bg-gray-50 transition flex items-center gap-2">
              Load Interval
            </button>
            <button onclick="copyIntervalReport()" id="copyReportBtn" class="bg-[#1a73e8] border border-[#1a73e8] text-white font-medium text-sm px-4 py-2 rounded shadow-sm hover:bg-blue-700 transition flex items-center gap-2 hidden">
              <span class="material-symbols-outlined text-[18px]">content_copy</span>
              Copy Report
            </button>
          </div>
        </div>

        <!-- Interval Table -->
        <div class="overflow-x-auto">
          <table class="w-full text-left border-collapse">
            <thead>
              <tr class="bg-gray-800 text-white text-xs tracking-wider">
                <th class="p-3 font-medium">LDAP</th>
                <th class="p-3 font-medium text-center">SOS</th>
                <th class="p-3 font-medium text-center">EOS</th>
                <th class="p-3 font-medium text-center">Site</th>
                <th class="p-3 font-medium text-center">Status</th>
                <th class="p-3 font-medium text-center">Cases Logged</th>
              </tr>
            </thead>
            <tbody id="intervalTableBody" class="text-sm divide-y divide-gray-200">
              <tr><td colspan="6" class="p-8 text-center text-gray-400">Select a date and time to load the interval.</td></tr>
            </tbody>
          </table>
        </div>
      </div>
    </div>

  </main>

  <!-- Notification Toast -->
  <div id="toast" class="fixed bottom-5 left-1/2 transform -translate-x-1/2 bg-gray-800 text-white px-6 py-3 rounded shadow-lg transition-opacity duration-300 opacity-0 hidden flex items-center gap-3 z-50">
    <span id="toastIcon" class="material-symbols-outlined">info</span>
    <span id="toastMessage" class="text-sm">Message</span>
  </div>

  <!-- JavaScript -->
  <script>
    // --- AVATAR HELPERS (Moved to global scope) ---
    function escAttr(str) {
      return String(str).replace(/&/g, '&amp;').replace(/"/g, '&quot;').replace(/'/g, '&#39;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    }

    function getMomaImgClass(ldap, cls, size) {
      size = size || 100;
      var sLdap = escAttr(ldap);
      var p = 'https://moma-teams-photos.corp.google.com/photos/' + sLdap + '?sz=' + size;
      var f = p + '&type=SECURITY&type=SILHOUETTE';
      return '<img src="' + p + '" class="' + cls + '" onerror="this.onerror=null;this.src=\'' + f + '\';">';
    }
    function formatCaseIdsAsLinks(idsString) {
      if (!idsString) return '';
      return idsString.split(',').map(id => {
        let cleanId = escAttr(id.trim());
        // Added whitespace-nowrap to the class list below:
        return `<a href="http://cases.connect.corp.google.com/${cleanId}" target="_blank" class="text-blue-600 hover:underline hover:text-blue-800 whitespace-nowrap">${cleanId}</a>`;
      }).join(', ');
    }
    // ----------------------------------------------

    // 1. App Initialization (Run on load)
    document.addEventListener("DOMContentLoaded", function() {
      // Call backend to get role and LDAP
      google.script.run
        .withSuccessHandler(initializeUserUI)
        .withFailureHandler(function(err){ console.error("Failed to load profile", err); })
        .getUserProfile();
    });

    // 2. Setup UI based on profile
    function initializeUserUI(profile) {
      const chip = document.getElementById('identityChip');
      
      // Generate the avatar HTML using the helper function
      const avatarHTML = getMomaImgClass(profile.ldap, 'w-10 h-10 rounded-full border border-gray-200 object-cover', 100);
      
      // Stop pulsing, inject data & avatar
      const safeName = escAttr(profile.name);
      const safeSite = escAttr(profile.site);
      const safeLob = escAttr(profile.lob);
      const safeWorkflow = escAttr(profile.workflow);

      chip.classList.remove('animate-pulse');
      chip.classList.remove('py-1.5', 'px-3'); 
      chip.classList.add('pr-4', 'pl-1', 'py-1'); 
      chip.innerHTML = `
        ${avatarHTML}
        <div class="flex flex-col ml-2">
          <span class="text-sm font-medium text-gray-800 leading-tight">${safeName}</span>
          <span class="text-[10px] text-gray-500 uppercase tracking-wide">${safeSite} • ${safeLob} ${safeWorkflow ? `(${safeWorkflow})` : ''}</span>
        </div>
      `;

      // Manager Setup
      if (profile.isManager) {
        document.getElementById('managerOnlyTabs').classList.remove('hidden');
        
        // Update My Submissions to Agent Submissions
        document.getElementById('tab-mySubmissionsView').innerHTML = `Agent Submissions`;
        document.getElementById('mySubmissionsTitle').innerHTML = `
          <span class="material-symbols-outlined text-blue-500">manage_search</span>
          Agent Submissions
        `;
        document.getElementById('mySubmissionsDesc').innerText = "View tracked cases per interval for any agent.";
        
        const searchContainer = document.getElementById('managerLdapSearch');
        const searchInput = document.getElementById('targetAgentLdap');
        searchContainer.classList.remove('hidden');
        searchInput.value = profile.ldap; // Default to themselves
        
        // Load Agents for autocomplete
        google.script.run
          .withSuccessHandler(function(agents) {
             const dataList = document.getElementById('agentList');
             dataList.innerHTML = agents.map(agent => `<option value="${agent}">`).join('');
          })
          .getAllAgents();
      }
      
      // Set default date for my submissions
      if(!document.getElementById('mySubmissionsDate').value) {
        document.getElementById('mySubmissionsDate').valueAsDate = new Date();
      }
    }

    // 3. Tab Switching Logic
    function switchTab(viewId) {
      ['submissionView', 'mySubmissionsView', 'dashboardView', 'intervalView'].forEach(id => {
        let el = document.getElementById(id);
        let tabEl = document.getElementById('tab-' + id);
        if (el) el.classList.add('hidden');
        if (tabEl) tabEl.className = "px-4 py-3 text-sm nav-inactive rounded-t-md transition-colors flex items-center gap-1";
      });
      
      let viewEl = document.getElementById(viewId);
      let tabViewEl = document.getElementById('tab-' + viewId);
      if (viewEl) viewEl.classList.remove('hidden');
      if (tabViewEl) tabViewEl.className = "px-4 py-3 text-sm nav-active transition-colors flex items-center gap-1";
      
      if(viewId === 'dashboardView') loadDashboard();
      if(viewId === 'mySubmissionsView') loadMySubmissions();
      if(viewId === 'intervalView' && !document.getElementById('intervalDate').value) {
         // Auto-fill today's date if empty
         document.getElementById('intervalDate').valueAsDate = new Date();
      }
    }

    // 4. Form Submission Logic
    function handleActivityChange() {
      const activity = document.querySelector('input[name="intervalActivity"]:checked').value;
      const caseIdsTextarea = document.getElementById('caseIds');
      
      if (activity === 'Coaching/Training') {
        caseIdsTextarea.disabled = true;
        caseIdsTextarea.required = false;
        caseIdsTextarea.classList.add('bg-gray-100', 'cursor-not-allowed');
        caseIdsTextarea.value = '';
      } else {
        caseIdsTextarea.disabled = false;
        caseIdsTextarea.required = true;
        caseIdsTextarea.classList.remove('bg-gray-100', 'cursor-not-allowed');
      }
    }

    function handleFormSubmit(event) {
      event.preventDefault();
      
      const activity = document.querySelector('input[name="intervalActivity"]:checked').value;
      const caseIdsText = document.getElementById('caseIds').value;
      
      if (activity === 'Break/Lunch') {
        let rawIds = caseIdsText.split(/[\n,;\s]+/).map(id => id.trim()).filter(id => id !== '');
        let uniqueIds = [...new Set(rawIds)];
        if (uniqueIds.length < 3) {
          showToast("You must track at least 3 case IDs during a Break/Lunch interval.", 'warning');
          return;
        }
      }

      const btn = document.getElementById('submitBtn');
      const spinner = document.getElementById('spinner');
      
      // UI: Loading state
      btn.disabled = true;
      btn.classList.add('opacity-75', 'cursor-not-allowed');
      spinner.classList.remove('hidden');

      const payload = {
        shiftType: document.querySelector('input[name="shiftType"]:checked').value,
        caseType: document.querySelector('input[name="caseType"]:checked') ? document.querySelector('input[name="caseType"]:checked').value : 'N/A',
        intervalActivity: activity,
        caseIdsText: caseIdsText
      };

      // Send to Backend
      google.script.run
        .withSuccessHandler(function(response) {
          resetButton(btn, spinner);
          
          if (response.success) {
            if (response.flagged > 0) {
              showToast(`Warning: ${response.flagged} duplicate(s) flagged. ${response.valid} valid cases logged.`, 'warning');
            } else {
              showToast(`Success! ${response.valid} cases securely logged.`, 'success');
            }
            document.getElementById('trackerForm').reset();
          } else {
            showToast(response.error, 'error');
          }
        })
        .withFailureHandler(function(error) {
          resetButton(btn, spinner);
          showToast("Network error. Please try again.", 'error');
        })
        .submitCases(payload);
    }

    function resetButton(btn, spinner) {
      btn.disabled = false;
      btn.classList.remove('opacity-75', 'cursor-not-allowed');
      spinner.classList.add('hidden');
    }

    function showToast(message, type) {
      const toast = document.getElementById('toast');
      const toastMsg = document.getElementById('toastMessage');
      const toastIcon = document.getElementById('toastIcon');

      toastMsg.textContent = message;
      
      if (type === 'success') {
        toast.className = "fixed bottom-5 left-1/2 transform -translate-x-1/2 bg-green-700 text-white px-6 py-3 rounded-full shadow-lg flex items-center gap-3 z-50 transition-opacity duration-300";
        toastIcon.textContent = "check_circle";
      } else if (type === 'warning') {
        toast.className = "fixed bottom-5 left-1/2 transform -translate-x-1/2 bg-amber-600 text-white px-6 py-3 rounded-full shadow-lg flex items-center gap-3 z-50 transition-opacity duration-300";
        toastIcon.textContent = "warning";
      } else {
        toast.className = "fixed bottom-5 left-1/2 transform -translate-x-1/2 bg-red-600 text-white px-6 py-3 rounded-full shadow-lg flex items-center gap-3 z-50 transition-opacity duration-300";
        toastIcon.textContent = "error";
      }

      toast.classList.remove('hidden', 'opacity-0');
      
      setTimeout(() => {
        toast.classList.add('opacity-0');
        setTimeout(() => toast.classList.add('hidden'), 300);
      }, 6000);
    }

    // --- DASHBOARD LOGIC ---

    // Load data when tab is clicked
    function loadDashboard() {
      document.getElementById('auditTableBody').innerHTML = '<tr><td colspan="5" class="p-8 text-center text-gray-400">Loading data...</td></tr>';
      document.getElementById('metricsTableBody').innerHTML = '<tr><td colspan="3" class="p-8 text-center text-gray-400">Loading data...</td></tr>';
      
      google.script.run
        .withSuccessHandler(renderDashboard)
        .withFailureHandler(err => showToast("Failed to load dashboard data.", 'error'))
        .getDashboardData();
    }

    // Render the tables
    function renderDashboard(data) {
      const auditTbody = document.getElementById('auditTableBody');
      const metricsTbody = document.getElementById('metricsTableBody');
      
      // 1. Render Audit Queue
      if (data.audits.length === 0) {
        auditTbody.innerHTML = '<tr><td colspan="5" class="p-8 text-center text-gray-400">🎉 No pending audits! Queue is clean.</td></tr>';
      } else {
        auditTbody.innerHTML = data.audits.map(a => `
          <tr class="hover:bg-gray-50 transition" id="audit-row-${a.row}">
            <td class="p-4">
              <div class="flex items-center gap-3">
                ${getMomaImgClass(a.agent, 'w-8 h-8 rounded-full border border-gray-200 object-cover', 80)}
                <div class="flex flex-col">
                  <span class="font-medium text-gray-800">${a.agent}</span>
                  <span class="text-[10px] text-gray-500 uppercase">${a.site}</span>
                </div>
              </div>
            </td>
            <td class="p-4">
              <div class="flex flex-col">
                <span class="text-gray-800">${a.timestamp}</span>
                <span class="text-xs text-gray-500">${a.caseType}</span>
              </div>
            </td>
            <td class="p-4 align-top">
              <div class="font-mono text-xs bg-red-50/50 rounded p-2 flex flex-wrap gap-2 leading-relaxed inline-flex">
                ${formatCaseIdsAsLinks(a.flaggedIds)}
              </div>
            </td>
            <td class="p-4 text-xs text-gray-600">${a.reason}</td>
            <td class="p-4 text-right">
              <div class="flex items-center justify-end gap-2">
                <button onclick="handleAudit(${a.row}, ${a.rawRowRef}, 'Approve')" class="text-green-600 bg-green-50 hover:bg-green-100 p-1.5 rounded transition" title="Approve">
                  <span class="material-symbols-outlined text-[20px]">check</span>
                </button>
                <button onclick="handleAudit(${a.row}, ${a.rawRowRef}, 'Reject')" class="text-red-600 bg-red-50 hover:bg-red-100 p-1.5 rounded transition" title="Reject">
                  <span class="material-symbols-outlined text-[20px]">close</span>
                </button>
              </div>
            </td>
          </tr>
        `).join('');
      }

      // 2. Render Metrics
      if (data.metrics.length === 0) {
        metricsTbody.innerHTML = '<tr><td colspan="3" class="p-8 text-center text-gray-400">No cases logged today yet.</td></tr>';
      } else {
        metricsTbody.innerHTML = data.metrics.map(m => `
          <tr class="hover:bg-gray-50 transition">
            <td class="p-4">
              <div class="flex items-center gap-3">
                ${getMomaImgClass(m.agent, 'w-8 h-8 rounded-full border border-gray-200 object-cover', 80)}
                <div class="flex flex-col">
                  <span class="font-medium text-gray-800">${m.agent}</span>
                  <span class="text-[10px] text-gray-500 uppercase">${m.site}</span>
                </div>
              </div>
            </td>
            <td class="p-4 text-center font-medium text-lg ${m.totalValid > 0 ? 'text-green-600' : 'text-gray-400'}">${m.totalValid}</td>
            <td class="p-4 text-center font-medium ${m.totalFlagged > 0 ? 'text-amber-500' : 'text-gray-300'}">${m.totalFlagged}</td>
          </tr>
        `).join('');
      }
    }

    // --- MY SUBMISSIONS LOGIC ---
    function loadMySubmissions() {
      document.getElementById('mySubmissionsTableBody').innerHTML = '<tr><td colspan="5" class="p-8 text-center text-gray-400">Loading submissions...</td></tr>';
      const dateVal = document.getElementById('mySubmissionsDate').value;
      const targetLdap = document.getElementById('targetAgentLdap').value.trim();
      
      if (!dateVal) return showToast("Please select a date", "warning");

      // Convert HTML YYYY-MM-DD to Apps Script M/d/yyyy
      const [year, month, day] = dateVal.split('-');
      const formattedDate = `${parseInt(month)}/${parseInt(day)}/${year}`;

      google.script.run
        .withSuccessHandler(renderMySubmissions)
        .withFailureHandler(err => {
          document.getElementById('mySubmissionsTableBody').innerHTML =
            '<tr><td colspan="5" class="p-8 text-center text-red-400">Failed to load submissions: ' +
            (err && err.message ? escAttr(err.message) : 'Unknown error') +
            '</td></tr>';
          showToast("Failed to load submissions.", 'error');
        })
        .getMySubmissions(formattedDate, targetLdap);
    }

    function renderMySubmissions(data) {
      const tbody = document.getElementById('mySubmissionsTableBody');
      try {
      if (data.length === 0) {
        tbody.innerHTML = '<tr><td colspan="5" class="p-8 text-center text-gray-400">No submissions found for this date.</td></tr>';
        return;
      }
      
      tbody.innerHTML = data.map(sub => `
        <tr class="hover:bg-gray-50 transition">
          <td class="p-4 whitespace-nowrap text-gray-800">${sub.interval}</td>
          <td class="p-4 text-gray-600">${sub.activity}</td>
          <td class="p-4 text-gray-600">${sub.caseType}</td>
          <td class="p-4 text-center font-medium ${sub.validCount > 0 ? 'text-green-600' : 'text-gray-400'}">${sub.validCount}</td>
          <td class="p-4 align-top">
             <div class="font-mono text-xs bg-gray-50 border border-gray-200 rounded p-2 flex flex-wrap gap-2 leading-relaxed inline-flex">
               ${formatCaseIdsAsLinks(sub.caseIds)}
             </div>
          </td>
        </tr>
      `).join('');
      } catch (e) {
        tbody.innerHTML = '<tr><td colspan="5" class="p-8 text-center text-red-400">Render error: ' + escAttr(e.message) + '</td></tr>';
      }
    }

    // Handle Approve/Reject Action
    function handleAudit(auditRow, rawRowRef, resolution) {
      // Optimistically hide the row so the UI feels instant
      const rowEl = document.getElementById(`audit-row-${auditRow}`);
      if (rowEl) rowEl.style.opacity = '0.5';
      
      google.script.run
        .withSuccessHandler(res => {
          if (res.success) {
            showToast(`Successfully ${resolution.toLowerCase()}d cases.`, 'success');
            loadDashboard(); // Reload to refresh metrics
          } else {
            showToast(res.error, 'error');
            if (rowEl) rowEl.style.opacity = '1'; // Revert if failed
          }
        })
        .withFailureHandler(err => {
          showToast("Network error.", 'error');
          if (rowEl) rowEl.style.opacity = '1';
        })
        .resolveAudit(auditRow, rawRowRef, resolution);
    }
    // --- INTERVAL VIEW LOGIC ---
    function loadIntervalData() {
      const dateVal = document.getElementById('intervalDate').value;
      const timeVal = document.getElementById('intervalHour').value;
      
      if (!dateVal) return showToast("Please select a date", "warning");
      
      // Convert HTML YYYY-MM-DD to Apps Script M/d/yyyy
      const [year, month, day] = dateVal.split('-');
      const formattedDate = `${parseInt(month)}/${parseInt(day)}/${year}`;

      document.getElementById('intervalTableBody').innerHTML = '<tr><td colspan="6" class="p-8 text-center text-gray-400">Loading schedule...</td></tr>';
      
      google.script.run
        .withSuccessHandler(renderIntervalTable)
        .withFailureHandler(err => showToast("Error loading interval.", "error"))
        .getIntervalData(formattedDate, timeVal);
    }

    function renderIntervalTable(agents) {
      const tbody = document.getElementById('intervalTableBody');
      const copyBtn = document.getElementById('copyReportBtn'); // Grabs the button
      
      if (agents.length === 0) {
        tbody.innerHTML = '<tr><td colspan="6" class="p-8 text-center text-gray-400">No agents scheduled for this interval.</td></tr>';
        copyBtn.classList.add('hidden'); // Hide if empty
        return;
      }
      
      copyBtn.classList.remove('hidden'); // Unhide if data exists
      
      tbody.innerHTML = agents.map(a => `
        <tr class="hover:bg-gray-50 transition ${a.isOT ? 'bg-amber-50/30' : ''}">
          <td class="p-3">
            <div class="flex items-center gap-2">
              ${getMomaImgClass(a.ldap, 'w-6 h-6 rounded-full border border-gray-200', 50)}
              <span class="font-medium text-blue-600">${a.ldap}</span>
              ${a.isOT ? '<span class="text-[10px] bg-amber-200 text-amber-800 px-1.5 rounded">OT</span>' : ''}
            </div>
          </td>
          <td class="p-3 text-center text-gray-600">${a.sos}</td>
          <td class="p-3 text-center text-gray-600">${a.eos}</td>
          <td class="p-3 text-center text-gray-600">${a.site}</td>
          <td class="p-3 text-center">
            <!-- Dynamic Status Dropdown -->
            <select onchange="updateStatusColor(this)" data-computed="${a.computedStatus}" class="text-xs p-1.5 border border-gray-200 rounded shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-500 w-36 font-medium cursor-pointer transition-colors bg-white text-gray-700">
              <option value="" class="bg-white text-gray-800" ${a.computedStatus === '' ? 'selected' : ''}></option>
              <option value="Assigned - 7" class="bg-white text-gray-800" ${a.computedStatus === 'Assigned - 7' ? 'selected' : ''}>Assigned - 7</option>
              <option value="Break - 3" class="bg-white text-gray-800" ${a.computedStatus === 'Break - 3' ? 'selected' : ''}>Break - 3</option>
              <option value="VL/SL" class="bg-white text-gray-800" ${a.computedStatus === 'VL/SL' ? 'selected' : ''}>VL/SL</option>
              <option value="on Live Channel" class="bg-white text-gray-800" ${a.computedStatus === 'on Live Channel' ? 'selected' : ''}>on Live Channel</option>
              <option value="SKIP - EOS" class="bg-white text-gray-800" ${a.computedStatus === 'SKIP - EOS' ? 'selected' : ''}>SKIP - EOS</option>
              <option value="SKIP SOS" class="bg-white text-gray-800" ${a.computedStatus === 'SKIP SOS' ? 'selected' : ''}>SKIP SOS</option>
              <option value="Absent" class="bg-white text-gray-800" ${a.computedStatus === 'Absent' ? 'selected' : ''}>Absent</option>
              <option value="on Coaching/Training" class="bg-white text-gray-800" ${a.computedStatus === 'on Coaching/Training' ? 'selected' : ''}>on Coaching/Training</option>
              <option value="Closing Reopens" class="bg-white text-gray-800" ${a.computedStatus === 'Closing Reopens' ? 'selected' : ''}>Closing Reopens</option>
            </select>
          </td>
          <td class="p-3 text-center font-medium ${a.casesLogged > 0 ? 'text-green-600' : 'text-gray-400'}">${a.casesLogged}</td>
        </tr>
      `).join('');
      
      // Update colors for dynamically selected statuses
      const selects = tbody.querySelectorAll('select');
      selects.forEach(select => updateStatusColor(select));
    }
    // --- STATUS DROPDOWN COLOR HELPER ---
    function updateStatusColor(selectElement) {
      const val = selectElement.value;
      // Reset classes
      selectElement.className = "text-xs p-1.5 border border-gray-200 rounded shadow-sm focus:outline-none focus:ring-2 focus:ring-blue-500 w-36 font-medium cursor-pointer transition-colors ";
      
      // Apply Google Sheets Chip Colors
      if (val === 'Assigned - 7') selectElement.classList.add('bg-[#1e8e3e]', 'text-white');
      else if (val === 'Break - 3') selectElement.classList.add('bg-[#f29900]', 'text-white');
      else if (val === 'VL/SL') selectElement.classList.add('bg-[#d2e3fc]', 'text-gray-800'); // Light Blue
      else if (val === 'on Live Channel') selectElement.classList.add('bg-[#f3e8fd]', 'text-purple-900'); // Light Purple
      else if (val === 'SKIP - EOS') selectElement.classList.add('bg-[#cbf0f8]', 'text-teal-900'); // Light Teal
      else if (val === 'SKIP SOS') selectElement.classList.add('bg-[#fce8e6]', 'text-red-900'); // Light Pink
      else if (val === 'Absent') selectElement.classList.add('bg-[#d93025]', 'text-white');
      else if (val === 'on Coaching/Training') selectElement.classList.add('bg-[#681da8]', 'text-white');
      else if (val === 'Closing Reopens') selectElement.classList.add('bg-[#5d4037]', 'text-white');
      else selectElement.classList.add('bg-white', 'text-gray-700');
    }

    // --- COPY REPORT LOGIC ---
    function copyIntervalReport() {
      const dateVal = document.getElementById('intervalDate').value;
      const timeSelect = document.getElementById('intervalHour');
      const timeVal = timeSelect.options[timeSelect.selectedIndex].text;
      
      const tbody = document.getElementById('intervalTableBody');
      const rows = tbody.querySelectorAll('tr');
      
      // Build the text using Tab spacing (\t) so it pastes perfectly into Sheets/Chat
      let reportText = `Interval Report: ${dateVal} | ${timeVal}\n\n`;
      reportText += `LDAP\tSOS\tEOS\tSite\tStatus\tCases Logged\n`; 

      rows.forEach(row => {
        const cells = row.querySelectorAll('td');
        if(cells.length === 6) {
          // Clean up the LDAP text (removes the visual 'OT' tag if present)
          const ldap = cells[0].innerText.replace('OT', '').trim(); 
          const sos = cells[1].innerText.trim();
          const eos = cells[2].innerText.trim();
          const site = cells[3].innerText.trim();
          
          // Grab the live selected value from the dropdown box!
          const statusSelect = cells[4].querySelector('select');
          const status = statusSelect ? statusSelect.value : '';
          
          const cases = cells[5].innerText.trim();
          
          reportText += `${ldap}\t${sos}\t${eos}\t${site}\t${status}\t${cases}\n`;
        }
      });

      // Write to the user's clipboard
      navigator.clipboard.writeText(reportText).then(() => {
        showToast("Report copied to clipboard!", "success");
      }).catch(err => {
        showToast("Failed to copy. Please copy manually.", "error");
      });
    }
  </script>
</body>
</html>
