const fs = require('fs');

let code = fs.readFileSync('Code.gs', 'utf8');

const processSelectionStart = code.indexOf('function processSelection(selectionData) {');
let processSelectionEnd = processSelectionStart;
let braceCount = 0;
let inFunction = false;

for (let i = processSelectionStart; i < code.length; i++) {
  if (code[i] === '{') {
    braceCount++;
    inFunction = true;
  } else if (code[i] === '}') {
    braceCount--;
    if (inFunction && braceCount === 0) {
      processSelectionEnd = i + 1;
      break;
    }
  }
}

const originalProcessSelection = code.substring(processSelectionStart, processSelectionEnd);

// Replace returns with returning { coreResult, createdRowIndices }
let coreFunction = originalProcessSelection.replace('function processSelection(selectionData) {', 'function _processSelectionCore(selectionData) {');

// Add before/after window capture and notification generation.
// Find the `return { success: ... };` statements and replace them to ensure notifications are created.

let modifiedCore = coreFunction;

// At the beginning, compute the beforeWindow:
modifiedCore = modifiedCore.replace(
  'const currentRound = configSheet.getRange("B2").getValue();',
  `const currentRound = configSheet.getRange("B2").getValue();
    const beforeWindowRaw = turnSheet.getDataRange().getValues();
    const beforeWindow = calculateQueueWindow(beforeWindowRaw, currentRound);`
);

// We need to replace all successful returns in the core with a helper that creates notifications.
// The return statements in question:
// return { success: true, message: "Selection recorded. Selection process is now complete." };
// return { success: true, message: "Selection recorded. Round 1 complete — awaiting lottery setup." };
// return { success: true, message: "Selection recorded." };

// Instead of string replacement which can be brittle, let's inject a wrapper inside _processSelectionCore:
modifiedCore = modifiedCore.replace('try {', `try {
    let coreResult = null;
    const executeLogic = () => {`);

// Close executeLogic before catch and call it.
modifiedCore = modifiedCore.replace(/} catch \(e\) {/g, `    }; // end executeLogic
    coreResult = executeLogic();
    if (!coreResult || !coreResult.success) {
      return { coreResult: coreResult || { success: false, message: "Unknown error" }, createdRowIndices: [] };
    }

    // Success path: compute after window and generated notifications
    const afterRound = configSheet.getRange("B2").getValue();
    const afterWindowRaw = turnSheet.getDataRange().getValues();
    const afterWindow = calculateQueueWindow(afterWindowRaw, afterRound);

    const createdRowIndices = computePendingNotifications(beforeWindow, afterWindow, afterRound, currentRound, afterWindowRaw);
    return { coreResult, createdRowIndices };
  } catch (e) {`);


let wrapperFunction = `
function processSelection(selectionData) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let pendingRowIndices = [];
  let finalResult = null;
  try {
    const res = _processSelectionCore(selectionData);
    finalResult = res.coreResult;
    pendingRowIndices = res.createdRowIndices || [];
  } finally {
    lock.releaseLock();
  }

  if (pendingRowIndices && pendingRowIndices.length > 0) {
    try {
      _processPendingNotifications(pendingRowIndices);
    } catch (e) {
      console.error("SMS notification processing failed, but selection succeeded: " + e.message);
    }
  }

  return finalResult;
}
`;


let finalCode = code.substring(0, processSelectionStart) + wrapperFunction + '\n' + modifiedCore + code.substring(processSelectionEnd);

fs.writeFileSync('Code.gs', finalCode);
