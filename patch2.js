const fs = require('fs');
let code = fs.readFileSync('/app/Index.html', 'utf8');

const search = `              <div class="flex items-center gap-2">
                \${getMomaImgClass(a.ldap, 'w-6 h-6 rounded-full border border-gray-200 dark:border-gray-600', 50)}
                <span class="font-medium text-theme-primary whitespace-nowrap">\${ldapLink(a.ldap)}</span>
                \${a.isOT ? '<span class="text-[10px] bg-amber-200 text-amber-800 px-1.5 rounded">OT</span>' : ''}
              </div>`;

const replace = `              <div class="flex items-center gap-2">
                \${getMomaImgClass(a.ldap, 'w-6 h-6 rounded-full border border-gray-200 dark:border-gray-600', 50)}
                <span class="font-medium text-theme-primary whitespace-nowrap">\${ldapLink(a.ldap)}</span>
                \${a.isOT ? '<span class="text-[10px] bg-amber-200 text-amber-800 px-1.5 rounded">OT</span>' : ''}
                \${a.pendingAudits && a.pendingAudits.length > 0 ? \`<span class="material-symbols-outlined text-[16px] text-red-500 cursor-pointer hover:text-red-700 transition" title="\${a.pendingAudits.length} flagged case(s) pending approval" onclick="openAgentFlaggedModal('\${a.ldap}')">error</span>\` : ''}
              </div>`;

if (code.includes(search)) {
  code = code.replace(search, replace);
  fs.writeFileSync('/app/Index.html', code);
  console.log('Success');
} else {
  console.log('Search block not found!');
}
