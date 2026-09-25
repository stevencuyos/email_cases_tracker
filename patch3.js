const fs = require('fs');
let code = fs.readFileSync('/app/Index.html', 'utf8');

// I also need to pass the JSON string of pendingAudits to the modal so we don't have to fetch again.
const search = `                \${a.pendingAudits && a.pendingAudits.length > 0 ? \`<span class="material-symbols-outlined text-[16px] text-red-500 cursor-pointer hover:text-red-700 transition" title="\${a.pendingAudits.length} flagged case(s) pending approval" onclick="openAgentFlaggedModal('\${a.ldap}')">error</span>\` : ''}`;
const replace = `                \${a.pendingAudits && a.pendingAudits.length > 0 ? \`<span class="material-symbols-outlined text-[16px] text-red-500 cursor-pointer hover:text-red-700 transition" title="\${a.pendingAudits.length} flagged case(s) pending approval" onclick="openAgentFlaggedModal('\${a.ldap}', '\${escAttr(JSON.stringify(a.pendingAudits))}')">error</span>\` : ''}`;

if (code.includes(search)) {
  code = code.replace(search, replace);
  fs.writeFileSync('/app/Index.html', code);
  console.log('Success');
} else {
  console.log('Search block not found!');
}
