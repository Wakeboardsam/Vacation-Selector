const fs = require('fs');

const code = fs.readFileSync('Code.gs', 'utf8');

// Quick mock of SpreadsheetApp for tests
const mockSpreadsheetApp = {
    getActiveSpreadsheet: () => ({
        getSheetByName: () => null
    })
};

const mockHtmlService = {
    createHtmlOutputFromFile: () => ({ setTitle: () => {} })
};

const LockService = {
    getScriptLock: () => ({ waitLock: () => {}, releaseLock: () => {} })
};

const context = {
    SpreadsheetApp: mockSpreadsheetApp,
    HtmlService: mockHtmlService,
    LockService: LockService,
    console: console,
    Error: Error
};

const vm = require('vm');
vm.createContext(context);
vm.runInContext(code, context);
vm.runInContext('testQueueWindowBehavior()', context);
vm.runInContext('testQueueWindowSkipNextTurn()', context);
vm.runInContext('testRound1SelectableWeeks()', context);
vm.runInContext('testInvalidClassification()', context);
console.log('All tests passed.');
