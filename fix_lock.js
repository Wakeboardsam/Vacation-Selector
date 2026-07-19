const fs = require('fs');

let code = fs.readFileSync('Code.gs', 'utf8');

// 1. Remove the lock acquisition and try/catch around the lock inside _processSelectionCore
// Original:
//   const lock = LockService.getScriptLock();
//   lock.waitLock(30000);
//   try {
//     let coreResult = null;
//     const executeLogic = () => {

// Let's replace the inner lock acquisition with just the try block setup.
code = code.replace(
  /const lock = LockService\.getScriptLock\(\);\s*lock\.waitLock\(30000\);\s*try {/g,
  `// Inner lock removed, managed by wrapper
  try {`
);

// We need to be careful as there might be a finally block releasing the lock in the inner function
// Let's find the inner function end.
//     return { success: true, message: "Selection recorded." };
//   } catch (e) {
//     return { success: false, message: "An error occurred: " + e.message };
//   } finally {
//     lock.releaseLock();
//   }

code = code.replace(
  /return \{ success: false, message: "An error occurred: " \+ e\.message \};\s*\} finally \{\s*lock\.releaseLock\(\);\s*\}/g,
  `return { coreResult: { success: false, message: "An error occurred: " + e.message } };
  }` // End of catch, no finally releasing the lock in the inner function
);

fs.writeFileSync('Code.gs', code);
