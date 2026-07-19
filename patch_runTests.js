const fs = require('fs');
let code = fs.readFileSync('Code.gs', 'utf8');

code = code.replace(
  'testDashboardDataExtraction();\n  runThemeColorTests();',
  'testDashboardDataExtraction();\n  runThemeColorTests();\n  runSmsTests();'
);

fs.writeFileSync('Code.gs', code);
