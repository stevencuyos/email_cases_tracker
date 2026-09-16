const fs = require('fs');
let codeGs = fs.readFileSync('Code.gs', 'utf8');

const regex = /let isManager = false;\s*for \(let i = 0; i < masterData\.length; i\+\+\) {/;
if (regex.test(codeGs)) {
  console.log('Old auth logic found');

  // Apply changes to Code.gs using replace_with_git_merge_diff tool
}
