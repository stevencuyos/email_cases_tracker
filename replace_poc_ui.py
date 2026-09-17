import sys

filepath = 'Index.html'
with open(filepath, 'r') as f:
    content = f.read()

search_str = """
      pocContainer.innerHTML = `
        <!-- Previous -->
        <div class="flex items-center justify-between px-3 py-1.5 rounded-theme-sm bg-gray-100 dark:bg-gray-700/50 text-gray-600 dark:text-gray-400 mb-1">
          <span class="text-[10px] font-medium">${prevTarget}</span>
          <span class="text-xs font-medium truncate max-w-[100px]">${prevPOC.poc}</span>
        </div>
        <!-- Current -->
        <div class="flex items-center justify-between px-3 py-2.5 rounded-theme-sm bg-green-50 dark:bg-green-900/30 text-green-700 dark:text-green-400 mb-1 border border-green-200/50 dark:border-green-800/50 relative overflow-hidden">
          <div class="absolute left-0 top-0 bottom-0 w-1 bg-green-500"></div>
          <span class="text-[11px] font-medium flex items-center gap-1.5 ml-1">
             <span class="w-1.5 h-1.5 rounded-full bg-green-500 animate-pulse"></span>
             ${currTarget}
          </span>
          <span class="text-sm font-bold truncate max-w-[100px]">${currPOC.poc}</span>
        </div>
        <!-- Next -->
        <div class="flex items-center justify-between px-3 py-1.5 rounded-theme-sm bg-gray-100 dark:bg-gray-700/50 text-gray-600 dark:text-gray-400">
          <span class="text-[10px] font-medium">${nextTarget}</span>
          <span class="text-xs font-medium truncate max-w-[100px]">${nextPOC.poc}</span>
        </div>
      `;
    }
"""

replace_str = """
      pocContainer.innerHTML = `
        <div class="relative flex flex-col gap-2 pl-3">
          <!-- Vertical Timeline Line -->
          <div class="absolute left-[22px] top-3 bottom-3 w-px bg-gray-200 dark:bg-gray-700"></div>

          <!-- Previous -->
          <div class="flex items-center gap-3 relative z-10 opacity-60">
            <div class="w-8 flex-shrink-0 text-right text-[10px] text-gray-500 font-medium">${prevTarget.split(' ')[0]}</div>
            <div class="w-2 h-2 rounded-full bg-gray-300 dark:bg-gray-600 flex-shrink-0 ring-4 ring-theme-card"></div>
            <div class="text-xs text-gray-600 dark:text-gray-400 font-medium truncate">${prevPOC.poc}</div>
          </div>

          <!-- Current -->
          <div class="flex items-center gap-3 relative z-10 py-1">
            <div class="w-8 flex-shrink-0 text-right text-[11px] font-bold text-theme-primary">${currTarget.split(' ')[0]}</div>
            <div class="relative w-2 h-2 flex-shrink-0">
               <div class="absolute inset-0 rounded-full bg-theme-primary animate-ping opacity-75"></div>
               <div class="relative w-2 h-2 rounded-full bg-theme-primary ring-4 ring-theme-card"></div>
            </div>
            <div class="text-sm font-bold text-gray-900 dark:text-gray-100 truncate">${currPOC.poc}</div>
          </div>

          <!-- Next -->
          <div class="flex items-center gap-3 relative z-10 opacity-60">
            <div class="w-8 flex-shrink-0 text-right text-[10px] text-gray-500 font-medium">${nextTarget.split(' ')[0]}</div>
            <div class="w-2 h-2 rounded-full bg-gray-300 dark:bg-gray-600 flex-shrink-0 ring-4 ring-theme-card"></div>
            <div class="text-xs text-gray-600 dark:text-gray-400 font-medium truncate">${nextPOC.poc}</div>
          </div>
        </div>
      `;

      // Auto-update mechanism: Refresh every 30 minutes.
      // Calculate milliseconds until the next half-hour or hour mark.
      const msNow = now.getTime();
      const minutes = now.getMinutes();
      let nextMinuteTarget = minutes < 30 ? 30 : 60;

      const nextUpdateDate = new Date(now);
      nextUpdateDate.setMinutes(nextMinuteTarget, 0, 0);

      const msUntilNextUpdate = nextUpdateDate.getTime() - msNow;

      // Clear any existing timeout to avoid duplicates
      if (window.pocUpdateTimeout) {
         clearTimeout(window.pocUpdateTimeout);
      }

      window.pocUpdateTimeout = setTimeout(() => {
         loadPOCSchedule();
      }, msUntilNextUpdate);
    }
"""

if search_str in content:
    content = content.replace(search_str, replace_str)
    with open(filepath, 'w') as f:
        f.write(content)
    print("Success")
else:
    print("Search string not found")
