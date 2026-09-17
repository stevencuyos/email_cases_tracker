import sys

filepath = 'Index.html'
with open(filepath, 'r') as f:
    content = f.read()

search_str = """
          <tr class="row-hover transition ${a.isOT ? 'bg-amber-50/30 dark:bg-amber-900/10' : ''} ${needsAttention ? 'bg-red-50/40 dark:bg-red-900/10' : ''}">
            <td class="p-3">
              <div class="flex items-center gap-2">
                ${getMomaImgClass(a.ldap, 'w-6 h-6 rounded-full border border-gray-200 dark:border-gray-600', 50)}
                <span class="font-medium case-link">${a.ldap}</span>
                ${a.isOT ? '<span class="text-[10px] bg-amber-200 text-amber-800 px-1.5 rounded">OT</span>' : ''}
              </div>
            </td>
"""

replace_str = """
          <tr class="row-hover transition ${a.isOT ? 'bg-amber-50/30 dark:bg-amber-900/10' : ''} ${needsAttention ? 'bg-red-50/40 dark:bg-red-900/10' : ''}">
            <td class="p-3">
              <div class="flex items-center gap-2 relative group cursor-pointer">
                ${getMomaImgClass(a.ldap, 'w-6 h-6 rounded-full border border-gray-200 dark:border-gray-600 shrink-0', 50)}
                <a href="https://cases.connect.corp.google.com/#/queues/org/gup/assignee%3A%20${a.ldap}%20state%3Aassigned" target="_blank" class="font-medium text-theme-primary hover:underline truncate z-10 block">${a.ldap}</a>
                ${a.isOT ? '<span class="text-[10px] bg-amber-200 text-amber-800 px-1.5 rounded z-10">OT</span>' : ''}
                <!-- Tooltip -->
                <div class="absolute left-full ml-2 top-1/2 -translate-y-1/2 z-50 invisible opacity-0 group-hover:visible group-hover:opacity-100 transition-all duration-200 pointer-events-none w-max max-w-xs">
                   <div class="bg-gray-800 dark:bg-gray-700 text-white text-xs rounded shadow-lg p-2 flex flex-col gap-1 border border-gray-700 dark:border-gray-600">
                     <div class="flex items-center gap-1"><span class="text-gray-400">Supervisor:</span> <span class="font-medium">${a.supervisor || 'Unknown'}</span></div>
                     <div class="flex items-center gap-1"><span class="text-gray-400">Site:</span> <span class="font-medium">${a.site || 'Unknown'}</span></div>
                   </div>
                </div>
              </div>
            </td>
"""

if search_str in content:
    content = content.replace(search_str, replace_str)
    with open(filepath, 'w') as f:
        f.write(content)
    print("Success")
else:
    print("Search string not found")
