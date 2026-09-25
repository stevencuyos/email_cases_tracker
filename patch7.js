const fs = require('fs');
let code = fs.readFileSync('/app/Index.html', 'utf8');

const search = `onclick="openAgentFlaggedModal('\${a.ldap}', '\${escAttr(JSON.stringify(a.pendingAudits))}')"`;
const replace = `data-audits="\${escAttr(JSON.stringify(a.pendingAudits))}" onclick="openAgentFlaggedModal('\${a.ldap}', this.getAttribute('data-audits'))"`;

if (code.includes(search)) {
  code = code.replace(search, replace);
  fs.writeFileSync('/app/Index.html', code);
  console.log('Success');
} else {
  console.log('Search block not found!');
}
