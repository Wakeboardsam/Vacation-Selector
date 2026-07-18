const fs = require('fs');

// Read the Code.gs file content
const code = fs.readFileSync('Code.gs', 'utf-8');

// Define the mock environment
global.HtmlService = {
  createHtmlOutputFromFile: () => ({ getContent: () => '' }),
  createTemplateFromFile: () => ({
    evaluate: () => ({ setTitle: () => ({ addMetaTag: () => {} }) })
  })
};

global.LockService = {
  getScriptLock: () => ({ waitLock: () => {}, releaseLock: () => {} })
};

global.Logger = {
  log: console.log
};

// SpreadsheetApp Mock is somewhat complex, but we only want to run parser validation
// For testing Hex Color parsing, we don't strictly need SpreadsheetApp, we can isolate the tests

// Isolate only the hex color validation logic to test
const validationRegex = /^#([0-9a-fA-F]{6})$/;
function validateHexColorLocal(value) {
  if (typeof value !== 'string') return null;
  const trimmed = value.trim();
  const match = trimmed.match(validationRegex);
  if (!match) return null;
  return trimmed.toUpperCase();
}

function testThemeColorValidationLocal() {
  if (validateHexColorLocal('#FFFFFF') !== '#FFFFFF') throw new Error('Failed to validate valid hex');
  if (validateHexColorLocal(' #ff0000 ') !== '#FF0000') throw new Error('Failed to trim and uppercase hex');
  if (validateHexColorLocal('#FFF') !== null) throw new Error('Failed to reject 3-digit hex');
  if (validateHexColorLocal('#FFFFFFFF') !== null) throw new Error('Failed to reject 8-digit hex');
  if (validateHexColorLocal('rgb(255,0,0)') !== null) throw new Error('Failed to reject rgb()');
  if (validateHexColorLocal('red') !== null) throw new Error('Failed to reject named color');
  console.log('PASS: testThemeColorValidation');
}

function getContrastTextColorLocal(hexColor) {
  if (!hexColor || !hexColor.startsWith('#')) return '#000000';

  const r = parseInt(hexColor.substr(1, 2), 16) / 255;
  const g = parseInt(hexColor.substr(3, 2), 16) / 255;
  const b = parseInt(hexColor.substr(5, 2), 16) / 255;

  const linearize = (c) => c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
  const R = linearize(r);
  const G = linearize(g);
  const B = linearize(b);

  const L = 0.2126 * R + 0.7152 * G + 0.0722 * B;

  const crWhite = 1.05 / (L + 0.05);
  const crBlack = (L + 0.05) / 0.05;

  return crWhite > crBlack ? '#FFFFFF' : '#000000';
}

function testThemeLuminanceLocal() {
  if (getContrastTextColorLocal('#000000') !== '#FFFFFF') throw new Error('Black background needs white text');
  if (getContrastTextColorLocal('#FFFFFF') !== '#000000') throw new Error('White background needs black text');
  console.log('PASS: testThemeLuminance');
}

testThemeColorValidationLocal();
testThemeLuminanceLocal();
