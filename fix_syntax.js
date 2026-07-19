const fs = require('fs');
let code = fs.readFileSync('Code.gs', 'utf8');
// Find the error line using a node syntax check that gives a line number
