const fs = require('fs');

let code = fs.readFileSync('Code.gs', 'utf8');

const setupStart = code.indexOf('function setupSpreadsheetSchema() {');
let setupEnd = setupStart;
let braceCount = 0;
let inFunction = false;

for (let i = setupStart; i < code.length; i++) {
  if (code[i] === '{') {
    braceCount++;
    inFunction = true;
  } else if (code[i] === '}') {
    braceCount--;
    if (inFunction && braceCount === 0) {
      setupEnd = i + 1;
      break;
    }
  }
}

const originalSetup = code.substring(setupStart, setupEnd);

let newSetup = originalSetup.replace(
  'return modifications ? "Schema updated successfully." : "Schema already up to date.";',
  `// 4. Add PhoneNumber if missing
  if (headers.indexOf('PhoneNumber') === -1) {
    let newCol = headers.length + 1;
    turnSheet.getRange(1, newCol).setValue('PhoneNumber');
    headers.push('PhoneNumber');
    modifications = true;
  }

  // 5. Setup Notification Log sheet
  let logSheet = ss.getSheetByName('Notification Log');
  let logModifications = false;
  if (!logSheet) {
    logSheet = ss.insertSheet('Notification Log');
    logModifications = true;
  }

  const logHeaders = ['Timestamp', 'DedupeKey', 'ParticipantName', 'Round', 'CalculatedRole', 'Status', 'TwilioMessageSid', 'Error'];
  let currentLogHeaders = [];
  if (logSheet.getLastColumn() > 0) {
     currentLogHeaders = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0];
  }

  for (let i = 0; i < logHeaders.length; i++) {
     if (currentLogHeaders.indexOf(logHeaders[i]) === -1) {
        logSheet.getRange(1, logSheet.getLastColumn() + 1 || 1).setValue(logHeaders[i]);
        currentLogHeaders.push(logHeaders[i]);
        logModifications = true;
     }
  }

  if (logModifications) {
      logSheet.setFrozenRows(1);
      logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).setFontWeight('bold').setBackground('#f3f4f6');
  }

  return (modifications || logModifications) ? "Schema updated successfully." : "Schema already up to date.";`
);

let finalCode = code.substring(0, setupStart) + newSetup + code.substring(setupEnd);
fs.writeFileSync('Code.gs', finalCode);
