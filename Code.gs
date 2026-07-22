// Final complete Code.gs file - Adds turn data to public view

// Theme color configuration constants
const ALLOWED_THEME_VARIABLES = [
  '--bg-color', '--surface-main', '--text-primary', '--text-secondary',
  '--accent-color', '--accent-hover', '--prime-tint', '--prime-text',
  '--available-tint', '--available-text', '--limited-tint', '--limited-text',
  '--full-tint', '--full-text', '--disabled-bg', '--disabled-text', '--skeleton-bg'
];

// Fallback configuration
const THEME_COLOR_DEFAULTS = {
  boring: {
    '--bg-color': '#F6F7FA', '--surface-main': '#FFFFFF', '--text-primary': '#171A22',
    '--text-secondary': '#6B7280', '--accent-color': '#5267E8', '--accent-hover': '#4052BF',
    '--prime-tint': '#FFF3D6', '--prime-text': '#92400E', '--available-tint': '#EAF8F0',
    '--available-text': '#166534', '--limited-tint': '#FFF4E5', '--limited-text': '#9A3412',
    '--full-tint': '#FDECEC', '--full-text': '#B91C1C', '--disabled-bg': '#E5E7EB',
    '--disabled-text': '#9CA3AF', '--skeleton-bg': '#E2E8F0'
  },
  anesthesia: {
    '--bg-color': '#E6F3F7', '--surface-main': '#FFFFFF', '--text-primary': '#1B2C33',
    '--text-secondary': '#566E7A', '--accent-color': '#039BE5', '--accent-hover': '#0277BD',
    '--prime-tint': '#FFF59D', '--prime-text': '#F57F17', '--available-tint': '#A5D6A7',
    '--available-text': '#1B5E20', '--limited-tint': '#FFCC80', '--limited-text': '#E65100',
    '--full-tint': '#EF9A9A', '--full-text': '#B71C1C', '--disabled-bg': '#CFD8DC',
    '--disabled-text': '#78909C', '--skeleton-bg': '#B0BEC5'
  },
  ketamine: {
    '--bg-color': '#1A0B2E', '--surface-main': '#2A1149', '--text-primary': '#E0E0E0',
    '--text-secondary': '#BDBDBD', '--accent-color': '#D500F9', '--accent-hover': '#AA00FF',
    '--prime-tint': '#C6FF00', '--prime-text': '#000000', '--available-tint': '#00E5FF',
    '--available-text': '#000000', '--limited-tint': '#FF3D00', '--limited-text': '#000000',
    '--full-tint': '#FF1744', '--full-text': '#000000', '--disabled-bg': '#4A148C',
    '--disabled-text': '#9C27B0', '--skeleton-bg': '#7B1FA2'
  }
};

const THEME_ROLE_LABELS = {
  '--bg-color': 'Page Background', '--surface-main': 'Card Surface',
  '--text-primary': 'Primary Text', '--text-secondary': 'Secondary Text',
  '--accent-color': 'Primary Accent', '--accent-hover': 'Accent Hover',
  '--prime-tint': 'Prime Background', '--prime-text': 'Prime Text',
  '--available-tint': 'Available Background', '--available-text': 'Available Text',
  '--limited-tint': 'Limited Background', '--limited-text': 'Limited Text',
  '--full-tint': 'Full Background', '--full-text': 'Full Text',
  '--disabled-bg': 'Disabled Background', '--disabled-text': 'Disabled Text',
  '--skeleton-bg': 'Loading Skeleton'
};

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Vacation Admin')
    .addItem('Set Up Admin Control', 'setupAdminControl')
    .addItem('Set Up Theme Colors', 'setupThemeColorsSheet')
    .addItem('Refresh Theme Swatches', 'refreshThemeColorSwatches')
    .addToUi();
}

/** Validates string strictly against #RRGGBB format */
function validateHexColor(value) {
  if (typeof value !== 'string') return null;
  const trimmed = value.trim();
  const match = trimmed.match(/^#([0-9a-fA-F]{6})$/);
  if (!match) return null;
  return trimmed.toUpperCase();
}

/** Gets best text color (black or white) for given valid #RRGGBB hex background using WCAG luminance */
function getContrastTextColor(hexColor) {
  if (!hexColor || !hexColor.startsWith('#')) return '#000000';

  const r = parseInt(hexColor.substr(1, 2), 16) / 255;
  const g = parseInt(hexColor.substr(3, 2), 16) / 255;
  const b = parseInt(hexColor.substr(5, 2), 16) / 255;

  const linearize = (c) => c <= 0.03928 ? c / 12.92 : Math.pow((c + 0.055) / 1.055, 2.4);
  const R = linearize(r);
  const G = linearize(g);
  const B = linearize(b);

  const L = 0.2126 * R + 0.7152 * G + 0.0722 * B;

  // Lw (white) = 1.0, Lb (black) = 0.0
  const crWhite = 1.05 / (L + 0.05);
  const crBlack = (L + 0.05) / 0.05;

  return crWhite > crBlack ? '#FFFFFF' : '#000000';
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function doGet(e) {
  const template = HtmlService.createTemplateFromFile('Index');

  const themeConfig = getThemeColorConfig();
  template.themeOverridesCss = buildThemeOverridesCss(themeConfig);

  return template
    .evaluate()
    .setTitle('Vacation Week Selection System')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function normalizeClassification(value) {
  const normalized = String(value || '').trim().toLowerCase();

  if (normalized === 'prime') return 'Prime';
  if (
    normalized === 'non-prime' ||
    normalized === 'non prime' ||
    normalized === 'nonprime'
  ) {
    return 'Non-Prime';
  }

  return null;
}

function buildAvailableWeekData(rows) {
  return rows
    .filter(row => Number(row[6]) > 0)
    .map(row => ({
      displayDate: row[0] instanceof Date ? row[0].toLocaleDateString("en-US", { timeZone: "UTC", month: 'short', day: 'numeric' }) : String(row[0]),
      valueDate: row[0] instanceof Date ? row[0].getTime() : null,
      classification: normalizeClassification(row[1]),
      spotsRemaining: Number(row[6]),
      originalClassification: row[1]
    }));
}

function getThemeColorConfig() {
  const config = {
    boring: { ...THEME_COLOR_DEFAULTS.boring },
    anesthesia: { ...THEME_COLOR_DEFAULTS.anesthesia },
    ketamine: { ...THEME_COLOR_DEFAULTS.ketamine }
  };

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName('Theme Colors');
    if (!sheet) return config;

    const data = sheet.getDataRange().getValues();
    if (data.length === 0) return config;

    const headers = data[0].map(h => String(h || '').trim().toLowerCase());

    // Find column indexes for each theme based on header name
    const themeColumns = {
      boring: headers.indexOf('boring'),
      anesthesia: headers.indexOf('anesthesia'),
      ketamine: headers.indexOf('ketamine')
    };

    // Track processed variables independently per theme
    const processedByTheme = {
      boring: new Set(),
      anesthesia: new Set(),
      ketamine: new Set()
    };

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const cssVarRaw = String(row[1] || '').trim();

      // Ignore unknown variables
      if (!ALLOWED_THEME_VARIABLES.includes(cssVarRaw)) {
        continue;
      }

      // Extract colors for each theme
      ['boring', 'anesthesia', 'ketamine'].forEach(theme => {
        const colIdx = themeColumns[theme];
        if (colIdx !== -1 && colIdx < row.length) {
          const rawValue = row[colIdx];
          const validColor = validateHexColor(rawValue);

          if (validColor && !processedByTheme[theme].has(cssVarRaw)) {
            config[theme][cssVarRaw] = validColor;
            processedByTheme[theme].add(cssVarRaw);
          }
        }
      });
    }

  } catch (e) {
    console.error("Failed to read theme colors:", e);
    // On error, fall back to default
  }

  return config;
}

function buildThemeOverridesCss(themeConfig) {
  let css = '';
  const themes = ['boring', 'anesthesia', 'ketamine'];

  themes.forEach(theme => {
    const selector = theme === 'boring' ? 'html[data-theme="boring"]' : `html[data-theme="${theme}"]`;
    css += `${selector} {\n`;

    ALLOWED_THEME_VARIABLES.forEach(cssVar => {
      const candidate = themeConfig && themeConfig[theme] && themeConfig[theme][cssVar];
      const value = validateHexColor(candidate) || THEME_COLOR_DEFAULTS[theme][cssVar];

      if (value) {
        css += `  ${cssVar}: ${value};\n`;
      }
    });

    css += `}\n\n`;
  });

  return css;
}

function setupThemeColorsSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Theme Colors');

  if (sheet) {
    return "Theme Colors tab already exists. Will not overwrite.";
  }

  sheet = ss.insertSheet('Theme Colors');

  // Setup headers
  const headers = ['Color Role', 'CSS Variable', 'Boring', 'Anesthesia', 'Ketamine'];
  sheet.appendRow(headers);
  sheet.setFrozenRows(1);

  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setFontWeight('bold');
  headerRange.setBackground('#f3f4f6');

  // Populate default variables
  const rows = [];
  ALLOWED_THEME_VARIABLES.forEach(cssVar => {
    const role = THEME_ROLE_LABELS[cssVar] || cssVar;
    const boringColor = THEME_COLOR_DEFAULTS.boring[cssVar] || '';
    const anesthesiaColor = THEME_COLOR_DEFAULTS.anesthesia[cssVar] || '';
    const ketamineColor = THEME_COLOR_DEFAULTS.ketamine[cssVar] || '';

    rows.push([role, cssVar, boringColor, anesthesiaColor, ketamineColor]);
  });

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
  }

  // Adjust column widths
  sheet.autoResizeColumn(1);
  sheet.autoResizeColumn(2);
  sheet.setColumnWidths(3, 3, 120);

  // Apply warning-only protection to Columns A and B
  const protectionA = sheet.getRange("A:A").protect().setDescription("Theme role identifiers are managed by the vacation-selection app.");
  protectionA.setWarningOnly(true);

  const protectionB = sheet.getRange("B:B").protect().setDescription("CSS-variable identifiers are managed by the vacation-selection app.");
  protectionB.setWarningOnly(true);

  // Refresh formatting to apply swatches
  refreshThemeColorSwatches();

  return "Theme Colors tab created and populated successfully.";
}

function refreshThemeColorSwatches() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Theme Colors');
  if (!sheet) return "Theme Colors tab not found.";

  const data = sheet.getDataRange().getValues();
  if (data.length < 2) return "No data to refresh.";

  const headers = data[0].map(h => String(h || '').trim().toLowerCase());

  const themeColumns = [
    headers.indexOf('boring'),
    headers.indexOf('anesthesia'),
    headers.indexOf('ketamine')
  ];

  const backgrounds = sheet.getDataRange().getBackgrounds();
  const fontColors = sheet.getDataRange().getFontColors();
  const fontWeights = sheet.getDataRange().getFontWeights();
  const fontStyles = sheet.getDataRange().getFontStyles();

  let changesMade = false;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];

    themeColumns.forEach(colIdx => {
      if (colIdx !== -1 && colIdx < row.length) {
        const rawValue = String(row[colIdx] || '');
        const validColor = validateHexColor(rawValue);

        if (validColor) {
          // Valid color: set background and contrast text
          backgrounds[i][colIdx] = validColor;
          fontColors[i][colIdx] = getContrastTextColor(validColor);
          fontWeights[i][colIdx] = 'bold';
          fontStyles[i][colIdx] = 'normal';
          changesMade = true;
        } else if (rawValue.trim() !== '') {
          // Invalid color: neutral warning format
          backgrounds[i][colIdx] = '#ffffff';
          fontColors[i][colIdx] = '#dc2626'; // Red text for error
          fontWeights[i][colIdx] = 'bold';
          fontStyles[i][colIdx] = 'italic';
          changesMade = true;
        } else {
          // Blank
          backgrounds[i][colIdx] = '#ffffff';
          fontColors[i][colIdx] = '#000000';
          fontWeights[i][colIdx] = 'normal';
          fontStyles[i][colIdx] = 'normal';
          changesMade = true;
        }
      }
    });
  }

  if (changesMade) {
    const dataRange = sheet.getDataRange();
    dataRange.setBackgrounds(backgrounds);
    dataRange.setFontColors(fontColors);
    dataRange.setFontWeights(fontWeights);
    dataRange.setFontStyles(fontStyles);
  }

  return "Theme color swatches refreshed.";
}

function getParticipantNames() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Turn Management');
  const names = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  return names.flat();
}
function verifyUser(name, pin) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Turn Management');
  const data = sheet.getDataRange().getValues();
  const userRow = data.find(row => row[0] === name);
  if (userRow) {
    if (String(userRow[1]) === String(pin)) { return "Success"; }
  }
  return "Invalid PIN";
}
function calculateQueueWindow(turnDataRaw, currentRound) {
    const turnData = [...turnDataRaw];
    const turnHeaders = turnData.shift();

    const nameIdx = turnHeaders.indexOf('Name');
    const statusIdx = turnHeaders.indexOf('Status');
    const senPosIdx = turnHeaders.indexOf('SeniorityPosition');
    const lotPosIdx = turnHeaders.indexOf('LotteryPosition');
    const skipIdx = turnHeaders.indexOf('SkipNextTurn');

    // Create an array of objects to sort easily
    let queue = turnData.map((row, index) => ({
        originalIndex: index + 1, // +1 because we shifted headers, and another +1 for 1-based indexing in sheets if needed, but we'll use original row data
        name: row[nameIdx],
        status: row[statusIdx],
        skipNextTurn: row[skipIdx],
        seniorityPosition: row[senPosIdx],
        lotteryPosition: row[lotPosIdx],
        queuePosition: currentRound === 1 ? row[senPosIdx] : row[lotPosIdx],
        computedStatus: 'Waiting'
    }));

    // Sort queue based on current round rules
    if (currentRound === 1) {
        queue.sort((a, b) => a.seniorityPosition - b.seniorityPosition);
    } else {
        const isEvenRound = currentRound % 2 === 0;
        if (isEvenRound) {
            queue.sort((a, b) => a.lotteryPosition - b.lotteryPosition);
        } else {
            queue.sort((a, b) => b.lotteryPosition - a.lotteryPosition);
        }
    }

    // Find the anchor: the first person who is not 'Completed' and not 'skipped' (if we were consuming skips)
    // Wait, the requirement says: "Do not consume SkipNextTurn in read-only getter functions. Consume skips only under the script lock during state transitions."
    // And "A completed or skipped interior position leaves a hole; it must not pull a fourth person into the window while the earliest unfinished person remains."

    let anchorIndex = queue.findIndex(person => person.status !== 'Completed');

    if (anchorIndex !== -1) {
        const windowSlots = ['Active', 'Standby', 'Backup'];
        // The permitted window consists of that queue position and the next two queue positions only.
        for (let offset = 0; offset < 3; offset++) {
            const personIndex = anchorIndex + offset;
            if (personIndex < queue.length) {
                let person = queue[personIndex];
                if (person.status !== 'Completed' && person.skipNextTurn !== true) {
                    person.computedStatus = windowSlots[offset];
                } else if (person.skipNextTurn === true && person.status !== 'Completed') {
                    // If they are skipping, they leave a hole.
                    person.computedStatus = 'Waiting';
                } else if (person.status === 'Completed') {
                     // If they are completed but inside the window (can happen if someone completes out of order), they leave a hole
                    person.computedStatus = 'Completed';
                }
            }
        }
    } else {
        // Everyone is completed
        queue.forEach(p => p.computedStatus = 'Completed');
    }

    // Sort back to original order or keep sorted order for return. We'll return sorted queue.
    return queue;
}

function getDashboardData(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const turnSheet = ss.getSheetByName('Turn Management');
  const weekSheet = ss.getSheetByName('Week Availability');
  const configSheet = ss.getSheetByName('Config');
  const turnDataRaw = turnSheet.getDataRange().getValues();
  const weekData = weekSheet.getDataRange().getValues();
  const currentRound = configSheet.getRange("B2").getValue();

  // Find user details in Turn Management
  const turnHeaders = turnDataRaw[0];
  const nameIdx = turnHeaders.indexOf('Name');
  const weeksSelectedIdx = turnHeaders.indexOf('WeeksSelected');
  const skipNextTurnIdx = turnHeaders.indexOf('SkipNextTurn');

  let weeksSelected = 0;
  let skipNextTurn = false;

  if (nameIdx !== -1) {
    const userRow = turnDataRaw.slice(1).find(row => row[nameIdx] === name);
    if (userRow) {
      if (weeksSelectedIdx !== -1) weeksSelected = Number(userRow[weeksSelectedIdx]) || 0;
      if (skipNextTurnIdx !== -1) skipNextTurn = Boolean(userRow[skipNextTurnIdx]);
    }
  }

  // Derive selectedWeeks from Week Availability sheet
  const weekHeaders = weekData[0];
  const selectedWeeks = [];

  // Person 1-4 columns are at indexes 2, 3, 4, 5
  weekData.slice(1).forEach(row => {
    let hasSelected = false;
    for (let i = 2; i <= 5; i++) {
      if (row[i] === name) {
        hasSelected = true;
        break;
      }
    }

    if (hasSelected) {
      selectedWeeks.push({
        valueDate: row[0] instanceof Date ? row[0].getTime() : null,
        displayDate: row[0] instanceof Date ? row[0].toLocaleDateString("en-US", { timeZone: "UTC", month: 'short', day: 'numeric' }) : String(row[0]),
        classification: normalizeClassification(row[1]) || row[1]
      });
    }
  });

  // Sort chronologically
  selectedWeeks.sort((a, b) => {
    if (a.valueDate && b.valueDate) return a.valueDate - b.valueDate;
    return 0;
  });

  weekData.shift();

  const queueWindow = calculateQueueWindow(turnDataRaw, currentRound);

  const userObj = queueWindow.find(p => p.name === name);
  const currentUser = {
    name: userObj ? userObj.name : name,
    queuePosition: userObj ? userObj.queuePosition : null,
    status: userObj ? userObj.computedStatus : 'Unknown',
    weeksSelected: weeksSelected,
    skipNextTurn: skipNextTurn,
    selectedWeeks: selectedWeeks
  };

  const turnQueue = queueWindow.map(p => ({
      name: p.name,
      queuePosition: p.queuePosition,
      status: p.computedStatus
  }));

  const availableWeeks = buildAvailableWeekData(weekData);

  return {
    currentUser: currentUser,
    turnQueue: turnQueue,
    availableWeeks: availableWeeks,
    currentRound: currentRound
  };
}
function getPublicCalendarData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const weekSheet = ss.getSheetByName('Week Availability');
  const turnSheet = ss.getSheetByName('Turn Management');
  const configSheet = ss.getSheetByName('Config');
  const currentRound = configSheet.getRange("B2").getValue();
  const weekData = weekSheet.getDataRange().getValues();
  const turnDataRaw = turnSheet.getDataRange().getValues();
  weekData.shift();

  const calendarData = weekData.map(row => ({
    startDate: row[0] instanceof Date ? row[0].toLocaleDateString("en-US", { timeZone: "UTC", month: 'short', day: 'numeric' }) : String(row[0]),
    classification: normalizeClassification(row[1]) || row[1], // fallback for calendar display
    person1: row[2], person2: row[3], person3: row[4], person4: row[5],
    spotsRemaining: row[6]
  }));

  const queueWindow = calculateQueueWindow(turnDataRaw, currentRound);

  const turnQueue = queueWindow.map(p => ({
    name: p.name,
    queuePosition: p.queuePosition,
    status: p.computedStatus
  }));
  
  return {
      calendarData: calendarData,
      turnQueue: turnQueue,
      currentRound: currentRound
  };
}
function setupSpreadsheetSchema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const turnSheet = ss.getSheetByName('Turn Management');
  if (!turnSheet) return "Turn Management sheet not found.";

  const headersRange = turnSheet.getRange(1, 1, 1, turnSheet.getLastColumn());
  let headers = headersRange.getValues()[0];

  let modifications = false;

  // 1. Rename QueuePosition to SeniorityPosition
  let qIndex = headers.indexOf('QueuePosition');
  if (qIndex !== -1) {
    turnSheet.getRange(1, qIndex + 1).setValue('SeniorityPosition');
    headers[qIndex] = 'SeniorityPosition';
    modifications = true;
  }

  // 2. Add LotteryPosition if missing
  if (headers.indexOf('LotteryPosition') === -1) {
    let newCol = headers.length + 1;
    turnSheet.getRange(1, newCol).setValue('LotteryPosition');
    headers.push('LotteryPosition');
    modifications = true;
  }

  // 3. Add SkipNextTurn if missing
  if (headers.indexOf('SkipNextTurn') === -1) {
    let newCol = headers.length + 1;
    turnSheet.getRange(1, newCol).setValue('SkipNextTurn');

    // Initialize to false
    if (turnSheet.getLastRow() > 1) {
      turnSheet.getRange(2, newCol, turnSheet.getLastRow() - 1, 1).setValue(false);
    }
    headers.push('SkipNextTurn');
    modifications = true;
  }

  // 4. Add PhoneNumber if missing
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

  return (modifications || logModifications) ? "Schema updated successfully." : "Schema already up to date.";
}

function validateSchema(turnData, currentRound) {
    const headers = turnData[0];
    const required = ['Name', 'PIN', 'SeniorityPosition', 'Status', 'WeeksSelected', 'LotteryPosition', 'SkipNextTurn'];
    for (let req of required) {
        if (headers.indexOf(req) === -1) {
            return { valid: false, message: "Missing required column: " + req + ". Please run setupSpreadsheetSchema()." };
        }
    }
    return { valid: true };
}

function checkLotteryReady(turnData) {
    const headers = turnData[0];
    const lotIdx = headers.indexOf('LotteryPosition');
    if (lotIdx === -1) return false;
    let usedPositions = new Set();
    for (let i = 1; i < turnData.length; i++) {
        let val = turnData[i][lotIdx];
        if (val === "" || val === null || val === undefined) return false;
        if (usedPositions.has(val)) return false;
        usedPositions.add(val);
    }
    return true;
}

function initializeLotteryRound() {
    setupSpreadsheetSchema();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const turnSheet = ss.getSheetByName('Turn Management');
    const configSheet = ss.getSheetByName('Config');

    const currentRound = configSheet.getRange("B2").getValue();
    if (currentRound !== 1) {
        return "Not in Round 1. No action taken.";
    }

    const turnDataRaw = turnSheet.getDataRange().getValues();
    const result = _transitionToRound2(turnSheet, configSheet, turnDataRaw);
    return result.message;
}


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

function _processSelectionCore(selectionData) {
  // Inner lock removed, managed by wrapper
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const turnSheet = ss.getSheetByName('Turn Management');
    const weekSheet = ss.getSheetByName('Week Availability');
    const configSheet = ss.getSheetByName('Config');

    // Capture before state outside of the inner closure
    const beforeRound = configSheet.getRange("B2").getValue();
    const beforeWindowRaw = turnSheet.getDataRange().getValues();
    const beforeWindow = calculateQueueWindow(beforeWindowRaw, beforeRound);

    let coreResult = null;
    const executeLogic = () => {
    let turnDataRaw = turnSheet.getDataRange().getValues();
    const currentRound = beforeRound;

    // Check if selection is started
    if (!isSelectionStarted()) {
        return { success: false, message: "The vacation selection process has not started yet." };
    }

    // Schema validation
    const schemaCheck = validateSchema(turnDataRaw, currentRound);
    if (!schemaCheck.valid) { return { success: false, message: "System setup error: " + schemaCheck.message }; }

    // Check if system is completely full
    const weekDataRaw = weekSheet.getDataRange().getValues();
    const weekData = weekDataRaw.slice(1);
    let totalSpotsRemaining = 0;
    weekData.forEach(row => {
        let empty = 0;
        for(let i=2; i<=5; i++) { if (!row[i]) empty++; }
        totalSpotsRemaining += empty;
    });
    if (totalSpotsRemaining === 0) {
        return { success: false, message: "Selection Complete: All available vacation slots are filled." };
    }

    const turnHeaders = turnDataRaw[0];
    const turnData = turnDataRaw.slice(1);

    const nameIdx = turnHeaders.indexOf('Name');
    const statusIdx = turnHeaders.indexOf('Status');
    const weeksSelectedIdx = turnHeaders.indexOf('WeeksSelected');
    const senPosIdx = turnHeaders.indexOf('SeniorityPosition');
    const lotPosIdx = turnHeaders.indexOf('LotteryPosition');
    const skipIdx = turnHeaders.indexOf('SkipNextTurn');

    const queueWindow = calculateQueueWindow(turnDataRaw, currentRound);
    const userObj = queueWindow.find(p => p.name === selectionData.name);

    if (!userObj) { return { success: false, message: "User not found." }; }

    if (!['Active', 'Standby', 'Backup'].includes(userObj.computedStatus)) {
        return { success: false, message: "It is not your turn to make a selection." };
    }

    // We also need the userRowIndex in the original turnData array
    const userRowIndex = turnData.findIndex(row => row[nameIdx] === selectionData.name);

    if (!selectionData.week1) {
        return { success: false, message: "Missing primary week selection." };
    }

    if (currentRound === 1 && selectionData.week2) {
        return { success: false, message: "Invalid selection. You can only select exactly ONE week during Round 1." };
    }

    if (selectionData.week1 === selectionData.week2) {
        return { success: false, message: "You cannot select the same week twice in one submission." };
    }

    let w1Index = weekData.findIndex(row => row[0].getTime() == selectionData.week1);
    let w2Index = selectionData.week2 ? weekData.findIndex(row => row[0].getTime() == selectionData.week2) : -1;

    if (w1Index === -1 || (selectionData.week2 && w2Index === -1)) {
        return { success: false, message: "One of the selected weeks does not exist." };
    }

    let w1Data = weekData[w1Index];
    let w2Data = selectionData.week2 ? weekData[w2Index] : null;

    // Validate classifications tightly on the backend before making any writes
    const class1 = normalizeClassification(w1Data[1]);
    const class2 = w2Data ? normalizeClassification(w2Data[1]) : null;

    if (!class1 || (w2Data && !class2)) {
        return { success: false, message: "Selected week has an invalid classification." };
    }

    if (w2Data && (class1 !== "Non-Prime" || class2 !== "Non-Prime")) {
        return { success: false, message: "Two-week selections must both be Non-Prime." };
    }

    // Check max capacity and existing spots
    let w1EmptySlots = 0;
    for (let i=2; i<=5; i++) { if (!w1Data[i]) w1EmptySlots++; }
    if (w1EmptySlots === 0) return { success: false, message: "Primary week is full." };

    if (w2Data) {
        let w2EmptySlots = 0;
        for (let i=2; i<=5; i++) { if (!w2Data[i]) w2EmptySlots++; }
        if (w2EmptySlots === 0) return { success: false, message: "Secondary week is full." };
    }

    // Check double booking in the same week
    for(let i=2; i<=5; i++){
        if (w1Data[i] === selectionData.name) return { success: false, message: "You are already booked for the primary week." };
        if (w2Data && w2Data[i] === selectionData.name) return { success: false, message: "You are already booked for the secondary week." };
    }

    // If validation passes, apply changes atomically.
    let w1TargetCol = -1;
    for(let i=2; i<=5; i++){
        if(!weekSheet.getRange(w1Index + 2, i + 1).getValue()) {
            w1TargetCol = i + 1;
            break;
        }
    }
    if (w1TargetCol === -1) throw new Error("Concurrency error: primary week filled up.");

    let w2TargetCol = -1;
    if (selectionData.week2) {
        for(let i=2; i<=5; i++){
            if(!weekSheet.getRange(w2Index + 2, i + 1).getValue()) {
                w2TargetCol = i + 1;
                break;
            }
        }
        if (w2TargetCol === -1) throw new Error("Concurrency error: secondary week filled up.");
    }

    // Perform writes
    weekSheet.getRange(w1Index + 2, w1TargetCol).setValue(selectionData.name);
    weekSheet.getRange(w1Index + 2, 7).setValue(w1EmptySlots - 1);
    totalSpotsRemaining -= 1;

    if (selectionData.week2) {
        weekSheet.getRange(w2Index + 2, w2TargetCol).setValue(selectionData.name);
        let currentSpots = (w2Data ? (4 - (w2Data.filter((_, idx) => idx >= 2 && idx <= 5 && w2Data[idx]).length)) : 0);
        let newW2Spots = Math.max(0, currentSpots - 1);
        weekSheet.getRange(w2Index + 2, 7).setValue(newW2Spots);
        totalSpotsRemaining -= 1;
    }

    // Update Turn Sheet for current user
    const weeksPickedCount = selectionData.week2 ? 2 : 1;
    const currentWeeksSelected = turnData[userRowIndex][weeksSelectedIdx];
    turnSheet.getRange(userRowIndex + 2, weeksSelectedIdx + 1).setValue(currentWeeksSelected + weeksPickedCount);
    turnSheet.getRange(userRowIndex + 2, statusIdx + 1).setValue('Completed');

    if (selectionData.week2) {
        turnSheet.getRange(userRowIndex + 2, skipIdx + 1).setValue(true);
    }

    // If we just filled the last spot in the whole sheet, clear queue and exit early
    if (totalSpotsRemaining === 0) {
        return { success: true, message: "Selection recorded. Selection process is now complete." };
    }

    // Since authorization relies on computed queue window and 'Completed' status,
    // we don't strictly need to write 'Waiting', 'Active', etc. to the sheet anymore
    // except for skips. Let's consume skips that fall within the *new* window.
    // Wait, the rule says: "Consume skips only under the script lock during state transitions."
    // If a person inside the new 3-person window has SkipNextTurn = true, we consume it and mark them Completed.

    let loopGuard = 0;
    let nextRound = currentRound;

    while (loopGuard < 100) {
        loopGuard++;
        let currentTurnDataRaw = turnSheet.getDataRange().getValues();
        let queueWindowData = calculateQueueWindow(currentTurnDataRaw, nextRound);

        // Find if anyone in the new window (offset 0, 1, 2 from anchor) has skipNextTurn = true
        let anchorIndex = queueWindowData.findIndex(person => person.status !== 'Completed');

        if (anchorIndex === -1) {
             // Round is over
             if (nextRound === 1) {
                 // Try to automatically transition to Round 2
                 const transitionResult = _transitionToRound2(turnSheet, configSheet, turnSheet.getDataRange().getValues());
                 if (!transitionResult.success) {
                     // Transition failed, stay in Round 1, but selection was successful
                     return { success: true, message: "Selection recorded. Round 1 complete. Could not auto-start Round 2: " + transitionResult.message };
                 }
                 // Transition succeeded, move to Round 2 processing
                 nextRound++;
                 turnDataRaw = turnSheet.getDataRange().getValues(); // Refresh stale memory data
                 continue;
             } else {
                 nextRound++;
                 configSheet.getRange("B2").setValue(nextRound);
                 let rows = turnSheet.getDataRange().getValues();
                 rows.shift();
                 rows.forEach((row, index) => {
                     turnSheet.getRange(index + 2, statusIdx + 1).setValue('Waiting');
                 });
                 turnDataRaw = turnSheet.getDataRange().getValues(); // Refresh stale memory data
                 continue; // re-evaluate for the new round
             }
        }

        let skippedSomeone = false;
        for (let offset = 0; offset < 3; offset++) {
            const personIndex = anchorIndex + offset;
            if (personIndex < queueWindowData.length) {
                let person = queueWindowData[personIndex];
                if (person.status !== 'Completed' && person.skipNextTurn === true) {
                    // Consume skip
                    let originalDataRow = turnData.findIndex(row => row[nameIdx] === person.name);
                    turnSheet.getRange(originalDataRow + 2, skipIdx + 1).setValue(false);
                    turnSheet.getRange(originalDataRow + 2, statusIdx + 1).setValue('Completed');
                    skippedSomeone = true;
                }
            }
        }

        if (!skippedSomeone) {
            break; // stable state
        }
    }

    return { success: true, message: "Selection recorded." };
      }; // end executeLogic
    coreResult = executeLogic();
    if (!coreResult || !coreResult.success) {
      return { coreResult: coreResult || { success: false, message: "Unknown error" }, createdRowIndices: [] };
    }

    // Success path: compute after window and generated notifications
    const afterRound = configSheet.getRange("B2").getValue();
    const afterWindowRaw = turnSheet.getDataRange().getValues();
    const afterWindow = calculateQueueWindow(afterWindowRaw, afterRound);

    let createdRowIndices = [];
    try {
      createdRowIndices = computePendingNotifications(beforeWindow, afterWindow, afterRound, beforeRound, afterWindowRaw);
    } catch (err) {
      console.error("Failed to compute pending notifications: " + err.message);
      createdRowIndices = [];
    }
    return { coreResult, createdRowIndices };
  } catch (e) {
    return { coreResult: { success: false, message: "An error occurred: " + e.message } };
  }
}

function testQueueWindowBehavior() {
    // 1, 2, 3 Active
    const headers = ['Name', 'PIN', 'SeniorityPosition', 'Status', 'WeeksSelected', 'LotteryPosition', 'SkipNextTurn'];
    const turnDataRaw = [
        headers,
        ['Person1', '1234', 1, 'Waiting', 0, 1, false],
        ['Person2', '1234', 2, 'Waiting', 0, 2, false],
        ['Person3', '1234', 3, 'Waiting', 0, 3, false],
        ['Person4', '1234', 4, 'Waiting', 0, 4, false],
        ['Person5', '1234', 5, 'Waiting', 0, 5, false],
    ];

    let computed = calculateQueueWindow(turnDataRaw, 1);
    if (computed[0].computedStatus !== 'Active' || computed[1].computedStatus !== 'Standby' || computed[2].computedStatus !== 'Backup') {
        throw new Error("Initial window incorrect");
    }
    if (computed[3].computedStatus !== 'Waiting') throw new Error("Person 4 should be waiting");

    // Person 2 completes
    turnDataRaw[2][3] = 'Completed';
    computed = calculateQueueWindow(turnDataRaw, 1);

    if (computed[0].computedStatus !== 'Active') throw new Error("Person 1 should be Active");
    if (computed[1].computedStatus !== 'Completed') throw new Error("Person 2 should be Completed");
    if (computed[2].computedStatus !== 'Backup') throw new Error("Person 3 should be Backup");
    if (computed[3].computedStatus !== 'Waiting') throw new Error("Person 4 should be Waiting");

    // Person 1 completes
    turnDataRaw[1][3] = 'Completed';
    computed = calculateQueueWindow(turnDataRaw, 1);

    if (computed[2].computedStatus !== 'Active') throw new Error("Person 3 should be Active");
    if (computed[3].computedStatus !== 'Standby') throw new Error("Person 4 should be Standby");
    if (computed[4].computedStatus !== 'Backup') throw new Error("Person 5 should be Backup");

    console.log('PASS: testQueueWindowBehavior');
}

function testQueueWindowSkipNextTurn() {
    const headers = ['Name', 'PIN', 'SeniorityPosition', 'Status', 'WeeksSelected', 'LotteryPosition', 'SkipNextTurn'];
    const turnDataRaw = [
        headers,
        ['Person1', '1234', 1, 'Waiting', 0, 1, false],
        ['Person2', '1234', 2, 'Waiting', 0, 2, true], // Skipping
        ['Person3', '1234', 3, 'Waiting', 0, 3, false],
        ['Person4', '1234', 4, 'Waiting', 0, 4, false],
    ];

    let computed = calculateQueueWindow(turnDataRaw, 1);
    if (computed[0].computedStatus !== 'Active') throw new Error("Person 1 should be Active");
    if (computed[1].computedStatus !== 'Waiting') throw new Error("Person 2 should be Waiting (skipped)");
    if (computed[2].computedStatus !== 'Backup') throw new Error("Person 3 should be Backup");
    if (computed[3].computedStatus !== 'Waiting') throw new Error("Person 4 should be Waiting (hole is left)");

    console.log('PASS: testQueueWindowSkipNextTurn');
}
function testThemeColorValidation() {
  if (validateHexColor('#FFFFFF') !== '#FFFFFF') throw new Error('Failed to validate valid hex');
  if (validateHexColor(' #ff0000 ') !== '#FF0000') throw new Error('Failed to trim and uppercase hex');
  if (validateHexColor('#FFF') !== null) throw new Error('Failed to reject 3-digit hex');
  if (validateHexColor('#FFFFFFFF') !== null) throw new Error('Failed to reject 8-digit hex');
  if (validateHexColor('rgb(255,0,0)') !== null) throw new Error('Failed to reject rgb()');
  if (validateHexColor('red') !== null) throw new Error('Failed to reject named color');
  console.log('PASS: testThemeColorValidation');
}

function testThemeLuminance() {
  if (getContrastTextColor('#000000') !== '#FFFFFF') throw new Error('Black background needs white text');
  if (getContrastTextColor('#FFFFFF') !== '#000000') throw new Error('White background needs black text');
  console.log('PASS: testThemeLuminance');
}

function testThemeParserRobustness() {
  let ss;
  let originalGetActive;
  try {
    ss = setupMockSpreadsheet();
    originalGetActive = SpreadsheetApp.getActiveSpreadsheet;
    SpreadsheetApp.getActiveSpreadsheet = () => ss;

    // Test missing tab
    let config = getThemeColorConfig();
    if (config.boring['--bg-color'] !== THEME_COLOR_DEFAULTS.boring['--bg-color']) {
      throw new Error("Missing tab should fall back to default");
    }

    // Insert tab with missing themes
    let sheet = ss.insertSheet('Theme Colors');
    sheet.appendRow(['Role', 'CSS Variable', 'Boring', 'Anesthesia']); // Add Anesthesia column for subsequent tests
    sheet.appendRow(['BG', '--bg-color', '#111111', 'invalid']); // Boring gets #111111, Anesthesia gets invalid

    config = getThemeColorConfig();
    if (config.boring['--bg-color'] !== '#111111') throw new Error("Should parse valid Boring color");
    if (config.ketamine['--bg-color'] !== THEME_COLOR_DEFAULTS.ketamine['--bg-color']) {
      throw new Error("Missing Ketamine column should fall back to default");
    }

    // Unknown variable and duplicate handling
    sheet.appendRow(['Unknown', '--unknown-var', '#222222']);
    sheet.appendRow(['BG Dup', '--bg-color', '#333333', '#222222']); // Boring ignores #333333 (already parsed valid row 2), Anesthesia gets #222222 (first valid)

    config = getThemeColorConfig();
    if (config.boring['--unknown-var'] !== undefined) throw new Error("Should ignore unknown variables");
    if (config.boring['--bg-color'] !== '#111111') throw new Error("Should ignore duplicate variables if already processed valid");
    if (config.anesthesia['--bg-color'] !== '#222222') throw new Error("Should accept later duplicate if first was not valid");

    // Invalid value fallback
    sheet.appendRow(['Surface', '--surface-main', 'invalid']);
    config = getThemeColorConfig();
    if (config.boring['--surface-main'] !== THEME_COLOR_DEFAULTS.boring['--surface-main']) {
      throw new Error("Invalid value should fall back to default individually");
    }

    // Invalid first, valid second duplicate
    sheet.appendRow(['Accent', '--accent-color', 'invalid', '#222222']); // Boring invalid, Anesthesia valid
    sheet.appendRow(['Accent', '--accent-color', '#333333', '#444444']); // Boring gets #333333 (first valid), Anesthesia ignores #444444 (already processed valid)
    config = getThemeColorConfig();

    if (config.boring['--accent-color'] !== '#333333') throw new Error("Boring should accept second row valid if first was invalid");
    if (config.anesthesia['--accent-color'] !== '#222222') throw new Error("Anesthesia should keep first row valid, ignore second");

    console.log('PASS: testThemeParserRobustness');
  } finally {
    if (originalGetActive) {
      SpreadsheetApp.getActiveSpreadsheet = originalGetActive;
    }
    if (ss) {
      const files = DriveApp.getFilesByName(ss.getName());
      while (files.hasNext()) files.next().setTrashed(true);
    }
  }
}

function testThemeCssGeneration() {
  const mockConfig = {
    boring: { '--bg-color': '#111111', '--accent-color': '#FFFFFF; } body { display:none' }, // Malicious injection attempt
    anesthesia: {},
    ketamine: {}
  };

  // Test builder does not throw when passed incomplete or malformed config
  let css;
  try {
     css = buildThemeOverridesCss(mockConfig);
  } catch(e) {
     throw new Error('Builder threw when passed incomplete config');
  }

  if (!css.includes('html[data-theme="boring"] {') || !css.includes('--bg-color: #111111;')) {
    throw new Error('Failed to generate safe CSS from config');
  }

  if (css.includes('display:none')) {
    throw new Error('Failed to reject injected malicious CSS');
  }

  if (!css.includes(`--accent-color: ${THEME_COLOR_DEFAULTS.boring['--accent-color']};`)) {
    throw new Error('Failed to fall back to default when injected CSS is invalid');
  }

  console.log('PASS: testThemeCssGeneration');
}

function runThemeColorTests() {
  testThemeColorValidation();
  testThemeLuminance();
  testThemeParserRobustness();
  testThemeCssGeneration();
}

function runTests() {
  testQueueWindowBehavior();
  testQueueWindowSkipNextTurn();
  testDashboardDataExtraction();
  runThemeColorTests();
  runSmsTests();
  runAdminTests();
}

function testRound1SelectableWeeks() {
  const rows = [
    [new Date('2026-06-01'), ' Prime ', '', '', '', '', 4],
    [new Date('2026-10-05'), 'non-prime', '', '', '', '', 3]
  ];

  const weeks = buildAvailableWeekData(rows);

  const primeCount = weeks.filter(
    week => week.classification === 'Prime'
  ).length;

  const nonPrimeCount = weeks.filter(
    week => week.classification === 'Non-Prime'
  ).length;

  if (primeCount !== 1) {
    throw new Error('Expected one selectable Prime week.');
  }

  if (nonPrimeCount !== 1) {
    throw new Error('Expected one selectable Non-Prime week.');
  }

  console.log('PASS: Round 1 includes Prime and Non-Prime options.');
}

function testInvalidClassification() {
  const rows = [
    [new Date('2026-06-01'), ' Prime ', '', '', '', '', 4],
    [new Date('2026-10-05'), 'UnknownType', '', '', '', '', 3]
  ];

  const weeks = buildAvailableWeekData(rows);
  const invalid = weeks.filter(w => w.classification === null);

  if (invalid.length !== 1) {
    throw new Error('Expected one invalid classification.');
  }

  console.log('PASS: Invalid classification properly detected.');
}

function testDashboardDataExtraction() {
    let ss;
    try {
        ss = setupMockSpreadsheet();
        const originalGetActive = SpreadsheetApp.getActiveSpreadsheet;
        SpreadsheetApp.getActiveSpreadsheet = () => ss;

        try {
            // Setup some test data in mock spreadsheet
            const turnSheet = ss.getSheetByName('Turn Management');
            const weekSheet = ss.getSheetByName('Week Availability');

            // Give Person1 2 weeks selected, and SkipNextTurn true
            // Headers: Name(1), PIN(2), SeniorityPosition(3), Status(4), WeeksSelected(5), LotteryPosition(6), SkipNextTurn(7)
            turnSheet.getRange(2, 5).setValue(2);
            turnSheet.getRange(2, 7).setValue(true);

            // Assign Person1 to week 1 (Prime) in P1 col
            weekSheet.getRange(2, 3).setValue('Person1');

            // Assign Person1 to week 2 (Non-Prime) in P2 col
            weekSheet.getRange(3, 4).setValue('Person1');

            // Assign Person2 to week 2 (Non-Prime) in P3 col
            weekSheet.getRange(3, 5).setValue('Person2');

            // Assign Person3 to week 2 (Non-Prime) in P4 col
            weekSheet.getRange(3, 6).setValue('Person3');

            // Test Person1 extraction (P1 and P2 cols)
            let dashboard1 = getDashboardData('Person1');
            if (dashboard1.currentUser.weeksSelected !== 2) throw new Error("Expected weeksSelected to be 2 for Person1");
            if (dashboard1.currentUser.skipNextTurn !== true) throw new Error("Expected skipNextTurn to be true for Person1");
            if (dashboard1.currentUser.selectedWeeks.length !== 2) throw new Error("Expected selectedWeeks length to be 2 for Person1");
            if (dashboard1.currentUser.selectedWeeks[0].classification !== 'Prime') throw new Error("Expected first selected week to be Prime");

            // Test Person2 extraction (P3 col)
            let dashboard2 = getDashboardData('Person2');
            if (dashboard2.currentUser.selectedWeeks.length !== 1) throw new Error("Expected selectedWeeks length to be 1 for Person2");

            // Test Person3 extraction (P4 col)
            let dashboard3 = getDashboardData('Person3');
            if (dashboard3.currentUser.selectedWeeks.length !== 1) throw new Error("Expected selectedWeeks length to be 1 for Person3");

            // Test Person4 extraction (no weeks)
            let dashboard4 = getDashboardData('Person4');
            if (dashboard4.currentUser.selectedWeeks.length !== 0) throw new Error("Expected selectedWeeks length to be 0 for Person4");

            console.log("PASS: testDashboardDataExtraction");

        } finally {
            SpreadsheetApp.getActiveSpreadsheet = originalGetActive;
        }
    } finally {
        if (ss) {
            const files = DriveApp.getFilesByName(ss.getName());
            while (files.hasNext()) files.next().setTrashed(true);
        }
    }
}

function runMoreTests() {
  testRound1SelectableWeeks();
  testInvalidClassification();
}

// ============================================================================
// INTEGRATION TESTS
// ============================================================================

function setupMockSpreadsheet() {
    // We create a temporary spreadsheet for true end-to-end testing
    const ss = SpreadsheetApp.create('Temp Test SS - Vacation Selection');

    const turnSheet = ss.insertSheet('Turn Management');
    turnSheet.appendRow(['Name', 'PIN', 'SeniorityPosition', 'Status', 'WeeksSelected', 'LotteryPosition', 'SkipNextTurn']);
    turnSheet.appendRow(['Person1', '1234', 1, 'Waiting', 0, 1, false]);
    turnSheet.appendRow(['Person2', '1234', 2, 'Waiting', 0, 2, false]);
    turnSheet.appendRow(['Person3', '1234', 3, 'Waiting', 0, 3, false]);
    turnSheet.appendRow(['Person4', '1234', 4, 'Waiting', 0, 4, false]);
    turnSheet.appendRow(['Person5', '1234', 5, 'Waiting', 0, 5, false]);

    const weekSheet = ss.insertSheet('Week Availability');
    weekSheet.appendRow(['WeekStartDate', 'Classification', 'Person1', 'Person2', 'Person3', 'Person4', 'SpotsRemaining']);
    weekSheet.appendRow([new Date('2026-06-01'), 'Prime', '', '', '', '', 4]);
    weekSheet.appendRow([new Date('2026-10-05'), 'Non-Prime', '', '', '', '', 4]);
    weekSheet.appendRow([new Date('2026-11-02'), 'Non-Prime', '', '', '', '', 4]);
    weekSheet.appendRow([new Date('2026-12-07'), 'Non-Prime', '', '', '', '', 4]); // Valid classification for preflight tests

    const configSheet = ss.insertSheet('Config');
    configSheet.getRange('A2').setValue('CurrentRound');
    configSheet.getRange('B2').setValue(1);
    configSheet.getRange('A3').setValue('SelectionStarted');
    configSheet.getRange('B3').setValue(false);

    // Delete the default 'Sheet1'
    const sheet1 = ss.getSheetByName('Sheet1');
    if (sheet1) ss.deleteSheet(sheet1);

    return ss;
}

function runIntegrationTests() {
    console.log("Running Integration Tests...");

    let ss;
    try {
        ss = setupMockSpreadsheet();
        // Temporarily override SpreadsheetApp.getActiveSpreadsheet to return our temp sheet
        const originalGetActive = SpreadsheetApp.getActiveSpreadsheet;
        SpreadsheetApp.getActiveSpreadsheet = () => ss;

        try {
            // Turn on SelectionStarted for general integration tests
            ss.getSheetByName('Config').getRange('B3').setValue(true);

            // TEST 1: Person 4 rejected initially
            let res = processSelection({ name: 'Person4', week1: ss.getSheetByName('Week Availability').getRange(2, 1).getValue().getTime() });
            if (res.success) throw new Error("Person 4 should have been rejected (outside window).");
            console.log("PASS: Person 4 rejected initially.");

            // TEST 2: Person 2 can select while Person 1 unfinished
            let weekTime2 = ss.getSheetByName('Week Availability').getRange(3, 1).getValue().getTime();
            res = processSelection({ name: 'Person2', week1: weekTime2 });
            if (!res.success) throw new Error("Person 2 should succeed.");
            console.log("PASS: Person 2 selected while Person 1 is unfinished.");

            // TEST 3: Person 4 STILL rejected
            res = processSelection({ name: 'Person4', week1: ss.getSheetByName('Week Availability').getRange(2, 1).getValue().getTime() });
            if (res.success) throw new Error("Person 4 should STILL be rejected (window anchored at Person 1).");
            console.log("PASS: Person 4 remains rejected because window is 1, 3, 4 (2 is completed). Wait... 1, 3, 4 is the window? No. 1, 2, 3 is the original window. 2 is completed. The permitted window consists of the anchor and the next TWO positions. So anchor=1, offset=0 (1), offset=1 (2-completed), offset=2 (3). So Person 4 is still offset 3, thus outside the window!");

            // TEST 4: Person 1 submits
            res = processSelection({ name: 'Person1', week1: ss.getSheetByName('Week Availability').getRange(2, 1).getValue().getTime() });
            if (!res.success) throw new Error("Person 1 should succeed.");
            console.log("PASS: Person 1 selected. Anchor should move to 3.");

            // TEST 5: Person 4 now accepted (Anchor=3, window=3,4,5)
            res = processSelection({ name: 'Person4', week1: ss.getSheetByName('Week Availability').getRange(2, 1).getValue().getTime() });
            if (!res.success) throw new Error("Person 4 should succeed now that anchor moved to 3.");
            console.log("PASS: Person 4 accepted in new window.");

            // TEST 6: Invalid Classification Request Rejected
            ss.getSheetByName('Week Availability').getRange(5, 2).setValue('UnknownType');
            let invalidTime = ss.getSheetByName('Week Availability').getRange(5, 1).getValue().getTime();
            res = processSelection({ name: 'Person3', week1: invalidTime });
            if (res.success) throw new Error("Person 3 selecting invalid classification should fail.");
            console.log("PASS: Invalid classification rejected.");

            // TEST 7: Descending Lottery Round
            ss.getSheetByName('Config').getRange('B2').setValue(3); // Round 3 is descending
            // Reset statuses to Waiting
            let turnSheet = ss.getSheetByName('Turn Management');
            for(let i=2; i<=6; i++) turnSheet.getRange(i, 4).setValue('Waiting');

            // Queue should be: Person5 (anchor), Person4, Person3, Person2, Person1
            // Person 2 should fail initially
            res = processSelection({ name: 'Person2', week1: weekTime2 });
            if (res.success) throw new Error("Person 2 should be rejected in descending round 3 (window is 5,4,3).");
            console.log("PASS: Descending round window calculation correct.");

        } finally {
            // Restore SpreadsheetApp
            SpreadsheetApp.getActiveSpreadsheet = originalGetActive;
        }
    } catch(e) {
        console.error("Integration test failed:", e.message);
        throw e;
    } finally {
        if (ss) {
            // Clean up temporary spreadsheet
            const files = DriveApp.getFilesByName(ss.getName());
            while (files.hasNext()) {
                files.next().setTrashed(true);
            }
        }
    }
}

// ============================================================================
// SMS NOTIFICATIONS
// ============================================================================

/** Dependencies mapping for testing */
const _smsDependencies = {
  getProperties: () => PropertiesService.getScriptProperties().getProperties(),
  fetch: (url, params) => UrlFetchApp.fetch(url, params)
};

/**
 * Validates SMS configuration
 * @returns {object} { valid: boolean, message: string }
 */
function checkSmsConfiguration() {
  const props = _smsDependencies.getProperties();
  const required = ['TWILIO_ACCOUNT_SID', 'TWILIO_AUTH_TOKEN', 'TWILIO_FROM_NUMBER', 'VACATION_SELECTOR_URL'];
  const missing = required.filter(key => !props[key] || String(props[key]).trim() === '');

  if (missing.length > 0) {
    return { valid: false, message: "Missing SMS configuration: " + missing.join(', ') };
  }
  return { valid: true, message: "SMS configuration is complete." };
}

/**
 * Checks if SMS notifications are globally enabled
 */
function isSmsEnabled() {
  const props = _smsDependencies.getProperties();
  const val = String(props['SMS_NOTIFICATIONS_ENABLED'] || '').toLowerCase();
  return val === 'true' || val === '1' || val === 'yes';
}

/**
 * Sends an SMS via Twilio
 * @param {string} to - E.164 phone number
 * @param {string} body - SMS content
 * @returns {object} { success: boolean, messageSid?: string, error?: string }
 */
function sendSmsViaTwilio(to, body) {
  const props = _smsDependencies.getProperties();
  const sid = props['TWILIO_ACCOUNT_SID'];
  const token = props['TWILIO_AUTH_TOKEN'];
  const from = props['TWILIO_FROM_NUMBER'];

  if (!sid || !token || !from) {
    return { success: false, error: "Missing Twilio credentials" };
  }

  const twilioUrl = `https://api.twilio.com/2010-04-01/Accounts/${sid}/Messages.json`;

  const payload = {
    "To": to,
    "From": from,
    "Body": body
  };

  const options = {
    method: "post",
    headers: {
      "Authorization": "Basic " + Utilities.base64Encode(`${sid}:${token}`)
    },
    payload: payload,
    muteHttpExceptions: true
  };

  try {
    const response = _smsDependencies.fetch(twilioUrl, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();

    let responseJson = {};
    try {
      responseJson = JSON.parse(responseText);
    } catch (e) {}

    if (responseCode >= 200 && responseCode < 300) {
      return { success: true, messageSid: responseJson.sid || "Unknown SID" };
    } else {
      return { success: false, error: `Twilio Error ${responseCode}: ${responseJson.message || responseText}` };
    }
  } catch (e) {
    return { success: false, error: "Request failed: " + e.message };
  }
}

/**
 * Computes pending notifications for new entrants to the window.
 * Writes PENDING rows to the Notification Log.
 * Must be called under the script lock.
 * @returns {number[]} Array of row indices created in the Notification Log
 */
function computePendingNotifications(beforeWindow, afterWindow, afterRound, currentRound, afterWindowRaw) {
  if (!isSmsEnabled()) return [];

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('Notification Log');
  if (!logSheet) return []; // Schema not set up yet

  const logData = logSheet.getDataRange().getValues();
  const logHeaders = logData[0] || [];
  const dedupeIdx = logHeaders.indexOf('DedupeKey');

  // If no dedupe col, we can't safely notify
  if (dedupeIdx === -1) return [];

  // Get existing dedupe keys to prevent duplicates
  const existingKeys = new Set();
  for (let i = 1; i < logData.length; i++) {
    if (logData[i][dedupeIdx]) {
      existingKeys.add(String(logData[i][dedupeIdx]));
    }
  }

  // Find participants who are in the Active, Standby, or Backup roles NOW
  const targetRoles = ['Active', 'Standby', 'Backup'];
  const newEntrants = [];

  afterWindow.forEach(afterPerson => {
    if (targetRoles.includes(afterPerson.computedStatus)) {
      // Were they in a target role BEFORE?
      // Wait, the requirement says "A participant should receive only one window-entry SMS per round, even if ... they move from Backup to Standby ... Standby to Active"
      // Therefore, the dedupe key is all that matters.
      // But to be clean, let's also check if they weren't in a target role before OR if round changed.
      const beforePerson = beforeWindow.find(p => p.name === afterPerson.name);

      let newlyEntered = false;
      if (!beforePerson) {
        newlyEntered = true;
      } else if (afterRound !== currentRound) {
        newlyEntered = true;
      } else if (!targetRoles.includes(beforePerson.computedStatus)) {
        newlyEntered = true;
      }

      const dedupeKey = `ROUND:${afterRound}|ENTERED_WINDOW|NAME:${afterPerson.name}`;

      // Even if newlyEntered is false, if they somehow lack a notification for this round's entry, send it.
      // The dedupe key is the ultimate source of truth.
      if (newlyEntered && !existingKeys.has(dedupeKey)) {
        newEntrants.push({
          name: afterPerson.name,
          role: afterPerson.computedStatus,
          dedupeKey: dedupeKey,
          round: afterRound
        });
        existingKeys.add(dedupeKey); // prevent dupes in the same batch
      }
    }
  });

  if (newEntrants.length === 0) return [];

  // We need phone numbers
  const turnHeaders = afterWindowRaw[0];
  const nameIdx = turnHeaders.indexOf('Name');
  const phoneIdx = turnHeaders.indexOf('PhoneNumber');

  const createdRowIndices = [];

  // Columns: Timestamp, DedupeKey, ParticipantName, Round, CalculatedRole, Status, TwilioMessageSid, Error
  const tsIdx = logHeaders.indexOf('Timestamp');
  const nameLogIdx = logHeaders.indexOf('ParticipantName');
  const roundIdx = logHeaders.indexOf('Round');
  const roleIdx = logHeaders.indexOf('CalculatedRole');
  const statusIdx = logHeaders.indexOf('Status');

  const nextRowIndex = logSheet.getLastRow() + 1;
  let currentRowOffset = 0;

  newEntrants.forEach(entrant => {
    let phoneNum = null;
    if (phoneIdx !== -1 && nameIdx !== -1) {
      const pRow = afterWindowRaw.find(r => r[nameIdx] === entrant.name);
      if (pRow) phoneNum = String(pRow[phoneIdx] || '').trim();
    }

    // Determine initial status based on phone number presence
    let initialStatus = 'PENDING';
    if (!phoneNum) {
      initialStatus = 'SKIPPED_NO_PHONE';
      console.log(`Skipping SMS for ${entrant.name} - no phone number.`);
    }

    const rowData = new Array(logHeaders.length).fill('');
    if (tsIdx !== -1) rowData[tsIdx] = new Date();
    if (dedupeIdx !== -1) rowData[dedupeIdx] = entrant.dedupeKey;
    if (nameLogIdx !== -1) rowData[nameLogIdx] = entrant.name;
    if (roundIdx !== -1) rowData[roundIdx] = entrant.round;
    if (roleIdx !== -1) rowData[roleIdx] = entrant.role;
    if (statusIdx !== -1) rowData[statusIdx] = initialStatus;

    logSheet.appendRow(rowData);

    if (initialStatus === 'PENDING') {
      createdRowIndices.push(nextRowIndex + currentRowOffset);
    }
    currentRowOffset++;
  });

  return createdRowIndices;
}

/**
 * Processes PENDING notifications, making external calls to Twilio.
 * Must run OUTSIDE the script lock.
 * @param {number[]} rowIndices - Indices in the Notification Log sheet
 */
function _processPendingNotifications(rowIndices) {
  if (!rowIndices || rowIndices.length === 0) return;

  const configCheck = checkSmsConfiguration();
  const props = _smsDependencies.getProperties();
  const vacationUrl = props['VACATION_SELECTOR_URL']; // No fallback

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('Notification Log');
  if (!logSheet) return;

  const logHeaders = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0];
  const nameIdx = logHeaders.indexOf('ParticipantName');
  const roundIdx = logHeaders.indexOf('Round');
  const roleIdx = logHeaders.indexOf('CalculatedRole');
  const statusIdx = logHeaders.indexOf('Status');
  const sidIdx = logHeaders.indexOf('TwilioMessageSid');
  const errorIdx = logHeaders.indexOf('Error');

  const turnSheet = ss.getSheetByName('Turn Management');
  const turnData = turnSheet.getDataRange().getValues();
  const tNameIdx = turnData[0].indexOf('Name');
  const tPhoneIdx = turnData[0].indexOf('PhoneNumber');

  rowIndices.forEach(rowIdx => {
    // Re-read row to ensure it's still PENDING
    const rowRange = logSheet.getRange(rowIdx, 1, 1, logHeaders.length);
    const rowValues = rowRange.getValues()[0];

    if (rowValues[statusIdx] !== 'PENDING') return;

    // Mark PROCESSING
    logSheet.getRange(rowIdx, statusIdx + 1).setValue('PROCESSING');

    const pName = rowValues[nameIdx];
    const pRound = rowValues[roundIdx];
    const pRole = rowValues[roleIdx];

    // Find phone number
    let phoneNum = null;
    if (tNameIdx !== -1 && tPhoneIdx !== -1) {
      const pRow = turnData.find(r => r[tNameIdx] === pName);
      if (pRow) phoneNum = String(pRow[tPhoneIdx] || '').trim();
    }

    if (!phoneNum) {
      logSheet.getRange(rowIdx, statusIdx + 1).setValue('SKIPPED_NO_PHONE');
      if (errorIdx !== -1) logSheet.getRange(rowIdx, errorIdx + 1).setValue('No phone number found');
      return;
    }

    if (!configCheck.valid) {
      logSheet.getRange(rowIdx, statusIdx + 1).setValue('FAILED');
      if (errorIdx !== -1) logSheet.getRange(rowIdx, errorIdx + 1).setValue(configCheck.message);
      return;
    }

    // Build and send SMS
    const msg = `Vacation Week Selection: You are now in the Round ${pRound} selection window as ${pRole}. Make your selection here: ${vacationUrl}`;
    const result = sendSmsViaTwilio(phoneNum, msg);

    if (result.success) {
      logSheet.getRange(rowIdx, statusIdx + 1).setValue('SENT');
      if (sidIdx !== -1) logSheet.getRange(rowIdx, sidIdx + 1).setValue(result.messageSid);
    } else {
      logSheet.getRange(rowIdx, statusIdx + 1).setValue('FAILED');
      if (errorIdx !== -1) logSheet.getRange(rowIdx, errorIdx + 1).setValue(result.error);
    }
  });
}

/**
 * Admin function to manually notify the current window
 */
function sendCurrentWindowNotifications() {
  if (!isSmsEnabled()) {
    return "SMS notifications are disabled in configuration.";
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  let pendingRowIndices = [];
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName('Config');
    const turnSheet = ss.getSheetByName('Turn Management');
    const currentRound = configSheet.getRange("B2").getValue();
    const turnDataRaw = turnSheet.getDataRange().getValues();

    // Simulate before window (empty) and after window (current) to trigger notifications for everyone in the window
    const currentWindow = calculateQueueWindow(turnDataRaw, currentRound);

    pendingRowIndices = computePendingNotifications([], currentWindow, currentRound, currentRound, turnDataRaw);
  } finally {
    lock.releaseLock();
  }

  if (pendingRowIndices && pendingRowIndices.length > 0) {
    try {
      _processPendingNotifications(pendingRowIndices);
      return `Processed ${pendingRowIndices.length} notifications.`;
    } catch (e) {
      return "Error processing notifications: " + e.message;
    }
  }
  return "No new notifications to send for the current window.";
}

/**
 * Admin function to send a test SMS explicitly bypassing some checks
 */
function sendTestSms(name) {
  const check = checkSmsConfiguration();
  if (!check.valid) {
    return "Cannot send test SMS: " + check.message;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const turnSheet = ss.getSheetByName('Turn Management');
  const turnData = turnSheet.getDataRange().getValues();
  const tNameIdx = turnData[0].indexOf('Name');
  const tPhoneIdx = turnData[0].indexOf('PhoneNumber');

  if (tNameIdx === -1 || tPhoneIdx === -1) {
    return "Phone number schema is not setup.";
  }

  const pRow = turnData.find(r => r[tNameIdx] === name);
  if (!pRow) return "Participant not found.";

  const phoneNum = String(pRow[tPhoneIdx] || '').trim();
  if (!phoneNum) return "Participant has no phone number on record.";

  const msg = "TEST MESSAGE: Vacation Week Selection SMS Configuration is working.";
  const result = sendSmsViaTwilio(phoneNum, msg);

  if (result.success) {
    return "Test SMS sent successfully. SID: " + result.messageSid;
  } else {
    return "Test SMS failed: " + result.error;
  }
}

// ============================================================================
// SMS TESTS
// ============================================================================

function testSmsConfigurationValidation() {
  const originalGet = _smsDependencies.getProperties;
  try {
    // Missing all
    _smsDependencies.getProperties = () => ({});
    if (checkSmsConfiguration().valid) throw new Error("Should fail when missing properties");

    // Has all
    _smsDependencies.getProperties = () => ({
      TWILIO_ACCOUNT_SID: '123',
      TWILIO_AUTH_TOKEN: '456',
      TWILIO_FROM_NUMBER: '789',
      VACATION_SELECTOR_URL: 'http://test.com'
    });
    if (!checkSmsConfiguration().valid) throw new Error("Should pass with all properties");

    console.log("PASS: testSmsConfigurationValidation");
  } finally {
    _smsDependencies.getProperties = originalGet;
  }
}

function testSmsEnabledFlag() {
  const originalGet = _smsDependencies.getProperties;
  try {
    _smsDependencies.getProperties = () => ({ SMS_NOTIFICATIONS_ENABLED: 'true' });
    if (!isSmsEnabled()) throw new Error("Should be enabled");

    _smsDependencies.getProperties = () => ({ SMS_NOTIFICATIONS_ENABLED: 'False' });
    if (isSmsEnabled()) throw new Error("Should be disabled");

    console.log("PASS: testSmsEnabledFlag");
  } finally {
    _smsDependencies.getProperties = originalGet;
  }
}

function testComputePendingNotifications() {
  const originalGet = _smsDependencies.getProperties;
  let ss;
  const originalGetActive = SpreadsheetApp.getActiveSpreadsheet;
  try {
    _smsDependencies.getProperties = () => ({ SMS_NOTIFICATIONS_ENABLED: 'true' });

    ss = SpreadsheetApp.create('Temp Test SS - SMS Notif');
    SpreadsheetApp.getActiveSpreadsheet = () => ss;

    // Missing Notification Log
    const indices1 = computePendingNotifications([], [], 1, 1, []);
    if (indices1.length > 0) throw new Error("Should handle missing log sheet safely");

    ss.insertSheet('Notification Log').appendRow(['Timestamp', 'DedupeKey', 'ParticipantName', 'Round', 'CalculatedRole', 'Status']);

    const beforeWindow = [
      { name: 'P1', computedStatus: 'Active' },
      { name: 'P2', computedStatus: 'Standby' }
    ];

    const afterWindow = [
      { name: 'P1', computedStatus: 'Completed' },
      { name: 'P2', computedStatus: 'Active' },
      { name: 'P3', computedStatus: 'Standby' },
      { name: 'P4', computedStatus: 'Backup' } // new entrants
    ];

    const turnHeaders = ['Name', 'PhoneNumber'];
    const turnDataRaw = [
      turnHeaders,
      ['P1', '111'],
      ['P2', '222'],
      ['P3', '333'],
      ['P4', '444'] // newly entered
    ];

    const indices2 = computePendingNotifications(beforeWindow, afterWindow, 1, 1, turnDataRaw);
    if (indices2.length !== 2) throw new Error("Expected exactly 2 new notifications (for P3 and P4)");

    // Deduplication check
    const indices3 = computePendingNotifications(beforeWindow, afterWindow, 1, 1, turnDataRaw);
    if (indices3.length !== 0) throw new Error("Should not create duplicates in same round");

    console.log("PASS: testComputePendingNotifications");
  } finally {
    _smsDependencies.getProperties = originalGet;
    SpreadsheetApp.getActiveSpreadsheet = originalGetActive;
    if (ss) {
      const files = DriveApp.getFilesByName(ss.getName());
      while (files.hasNext()) files.next().setTrashed(true);
    }
  }
}

function testProcessPendingNotificationsMissingPhone() {
  const originalGet = _smsDependencies.getProperties;
  let ss;
  const originalGetActive = SpreadsheetApp.getActiveSpreadsheet;
  try {
    _smsDependencies.getProperties = () => ({ SMS_NOTIFICATIONS_ENABLED: 'true', VACATION_SELECTOR_URL: 'http' });
    ss = SpreadsheetApp.create('Temp Test SS - Process SMS');
    SpreadsheetApp.getActiveSpreadsheet = () => ss;

    const logSheet = ss.insertSheet('Notification Log');
    logSheet.appendRow(['Timestamp', 'DedupeKey', 'ParticipantName', 'Round', 'CalculatedRole', 'Status', 'TwilioMessageSid', 'Error']);
    logSheet.appendRow(['', 'k1', 'NoPhonePerson', 1, 'Active', 'PENDING', '', '']);

    const turnSheet = ss.insertSheet('Turn Management');
    turnSheet.appendRow(['Name', 'PhoneNumber']);
    turnSheet.appendRow(['NoPhonePerson', '']); // blank phone

    _processPendingNotifications([2]); // row index 2

    const status = logSheet.getRange(2, 6).getValue();
    if (status !== 'SKIPPED_NO_PHONE') throw new Error("Should skip if no phone number");

    console.log("PASS: testProcessPendingNotificationsMissingPhone");
  } finally {
    _smsDependencies.getProperties = originalGet;
    SpreadsheetApp.getActiveSpreadsheet = originalGetActive;
    if (ss) {
      const files = DriveApp.getFilesByName(ss.getName());
      while (files.hasNext()) files.next().setTrashed(true);
    }
  }
}

function runSmsTests() {
  testSmsConfigurationValidation();
  testSmsEnabledFlag();
  testComputePendingNotifications();
  testProcessPendingNotificationsMissingPhone();
  testProcessSelectionSmsIntegration();
  testProcessPendingNotificationsMissingConfig();
}

function testProcessSelectionSmsIntegration() {
  const originalGet = _smsDependencies.getProperties;
  let ss;
  const originalGetActive = SpreadsheetApp.getActiveSpreadsheet;
  try {
    // Enable SMS
    _smsDependencies.getProperties = () => ({ SMS_NOTIFICATIONS_ENABLED: 'true', VACATION_SELECTOR_URL: 'http' });

    ss = setupMockSpreadsheet();
    SpreadsheetApp.getActiveSpreadsheet = () => ss;

    // Setup Notification Log & PhoneNumber schema
    setupSpreadsheetSchema();
    const turnSheet = ss.getSheetByName('Turn Management');
    const logSheet = ss.getSheetByName('Notification Log');

    // Set Phone numbers for test
    const turnData = turnSheet.getDataRange().getValues();
    const headers = turnData[0];
    const phoneIdx = headers.indexOf('PhoneNumber');
    for (let i = 1; i <= 5; i++) {
        turnSheet.getRange(i + 1, phoneIdx + 1).setValue('555-1234');
    }

    // Simulate Person 1 selection (moving window)
    const weekTime = ss.getSheetByName('Week Availability').getRange(2, 1).getValue().getTime();
    const res = processSelection({ name: 'Person1', week1: weekTime });

    if (!res.success) throw new Error("Selection should be successful");

    const logData = logSheet.getDataRange().getValues();
    const logHeaders = logData[0];

    // When Person 1 completes, the window moves to 2, 3, 4.
    // Person 4 should be the new entrant (Backup).
    const nameLogIdx = logHeaders.indexOf('ParticipantName');

    const newEntrantsLogs = logData.slice(1).filter(r => r[nameLogIdx] === 'Person4');
    if (newEntrantsLogs.length !== 1) throw new Error("Expected exactly one notification log for the new entrant (Person 4)");

    // Ensure failure to log doesn't fail processSelection
    // Sabotage computePendingNotifications by renaming the sheet so it throws or fails gracefully
    logSheet.setName('HiddenLog');

    const weekTime2 = ss.getSheetByName('Week Availability').getRange(3, 1).getValue().getTime();
    const res2 = processSelection({ name: 'Person2', week1: weekTime2 });

    if (!res2.success) throw new Error("Selection should be successful even if SMS/logging fails");

    console.log("PASS: testProcessSelectionSmsIntegration");
  } finally {
    _smsDependencies.getProperties = originalGet;
    SpreadsheetApp.getActiveSpreadsheet = originalGetActive;
    if (ss) {
      const files = DriveApp.getFilesByName(ss.getName());
      while (files.hasNext()) files.next().setTrashed(true);
    }
  }
}

function testProcessPendingNotificationsMissingConfig() {
  const originalGet = _smsDependencies.getProperties;
  const originalFetch = _smsDependencies.fetch;
  let ss;
  const originalGetActive = SpreadsheetApp.getActiveSpreadsheet;
  let fetchCalled = false;

  try {
    // Missing VACATION_SELECTOR_URL
    _smsDependencies.getProperties = () => ({
      SMS_NOTIFICATIONS_ENABLED: 'true',
      TWILIO_ACCOUNT_SID: '123',
      TWILIO_AUTH_TOKEN: '456',
      TWILIO_FROM_NUMBER: '789'
    });

    _smsDependencies.fetch = () => {
       fetchCalled = true;
       return { getResponseCode: () => 200, getContentText: () => '{}' };
    };

    ss = SpreadsheetApp.create('Temp Test SS - Config Missing');
    SpreadsheetApp.getActiveSpreadsheet = () => ss;

    const logSheet = ss.insertSheet('Notification Log');
    logSheet.appendRow(['Timestamp', 'DedupeKey', 'ParticipantName', 'Round', 'CalculatedRole', 'Status', 'TwilioMessageSid', 'Error']);
    logSheet.appendRow(['', 'k1', 'PersonWithPhone', 1, 'Active', 'PENDING', '', '']);

    const turnSheet = ss.insertSheet('Turn Management');
    turnSheet.appendRow(['Name', 'PhoneNumber']);
    turnSheet.appendRow(['PersonWithPhone', '555-1234']);

    _processPendingNotifications([2]); // row index 2

    const status = logSheet.getRange(2, 6).getValue();
    const errorMsg = logSheet.getRange(2, 8).getValue();

    if (fetchCalled) throw new Error("Should not call Twilio fetch if config is invalid");
    if (status !== 'FAILED') throw new Error("Status should be FAILED");
    if (String(errorMsg).indexOf('Missing SMS configuration') === -1) throw new Error("Should log configuration error");

    console.log("PASS: testProcessPendingNotificationsMissingConfig");
  } finally {
    _smsDependencies.getProperties = originalGet;
    _smsDependencies.fetch = originalFetch;
    SpreadsheetApp.getActiveSpreadsheet = originalGetActive;
    if (ss) {
      const files = DriveApp.getFilesByName(ss.getName());
      while (files.hasNext()) files.next().setTrashed(true);
    }
  }
}
function setupAdminControl(skipTrigger = false) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName('Admin Control');

  if (!sheet) {
    sheet = ss.insertSheet('Admin Control');
  } else {
    sheet.clear();
  }

  // Set column widths
  sheet.setColumnWidth(1, 60);
  sheet.setColumnWidth(2, 300);

  // A1:B1 - VACATION ADMIN CONTROL
  sheet.getRange('A1:B1').merge().setValue('VACATION ADMIN CONTROL')
    .setFontWeight('bold').setFontSize(14).setHorizontalAlignment('center');

  // A2:B2 - Short instruction
  sheet.getRange('A2:B2').merge().setValue('Complete the checklist, then check Start Round 1.')
    .setFontStyle('italic').setHorizontalAlignment('center').setWrap(true);

  // A4:B4 - PRE-START CHECKLIST
  sheet.getRange('A4:B4').merge().setValue('PRE-START CHECKLIST')
    .setFontWeight('bold').setBackground('#f3f4f6');

  // A5:B9 - Checkboxes and labels
  sheet.getRange('A5:A9').insertCheckboxes();
  sheet.getRange('B5').setValue('Roster reviewed');
  sheet.getRange('B6').setValue('Seniority order reviewed');
  sheet.getRange('B7').setValue('Lottery order reviewed');
  sheet.getRange('B8').setValue('Week calendar Dates reviewed');
  sheet.getRange('B9').setValue('Phone numbers and PINs reviewed');

  // A11:B11 - START SELECTION
  sheet.getRange('A11:B11').merge().setValue('START SELECTION')
    .setFontWeight('bold').setBackground('#f3f4f6');

  // A12:B12 - Action checkbox
  sheet.getRange('A12').insertCheckboxes();
  sheet.getRange('B12').setValue('START ROUND 1').setFontWeight('bold');

  // A14:B14 - LAST ACTION RESULT
  sheet.getRange('A14:B14').merge().setValue('LAST ACTION RESULT')
    .setFontWeight('bold').setBackground('#f3f4f6');

  // A15:B17 - Status, Timestamp, Details
  sheet.getRange('A15').setValue('Status:');
  sheet.getRange('A16').setValue('Timestamp:');
  sheet.getRange('A17').setValue('Details:');

  sheet.getRange('B15').setValue('NOT STARTED');
  sheet.getRange('B16').setValue('-');
  sheet.getRange('B17').setValue('-');
  sheet.getRange('B15:B17').setWrap(true);

  // Increase row heights for mobile readability
  for (let r = 1; r <= 17; r++) {
    sheet.setRowHeight(r, 40);
  }

  // Install trigger
  if (!skipTrigger) {
    installAdminControlTrigger();
  }

  return "Admin Control tab created successfully.";
}

function installAdminControlTrigger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const triggers = ScriptApp.getUserTriggers(ss);

  let triggerExists = false;
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'adminControlOnEdit') {
      triggerExists = true;
      break;
    }
  }

  if (!triggerExists) {
    ScriptApp.newTrigger('adminControlOnEdit')
      .forSpreadsheet(ss)
      .onEdit()
      .create();
  }
}

function adminControlOnEdit(e) {
  if (!e || !e.range) return;
  const range = e.range;
  const sheet = range.getSheet();

  if (sheet.getName() !== 'Admin Control') return;

  // Check if edited cell is A12 (Start Round 1 checkbox)
  if (range.getRow() === 12 && range.getColumn() === 1) {
    const isChecked = range.getValue() === true;
    if (!isChecked) return; // Only process when checked

    // We should lock here
    const lock = LockService.getScriptLock();
    // Wait for up to 30 seconds
    const locked = lock.tryLock(30000);
    if (!locked) {
      sheet.getRange('B15').setValue('❌ ERROR');
      sheet.getRange('B16').setValue(new Date().toLocaleString());
      sheet.getRange('B17').setValue('Could not acquire system lock. Please try again.');
      range.setValue(false); // Reset
      return;
    }

    let pendingRowIndices = [];

    try {
      // Re-read A12 to ensure it wasn't double tapped
      if (range.getValue() !== true) return;

      // Verify A5:A9 are true
      const checklistValues = sheet.getRange('A5:A9').getValues();
      const allChecked = checklistValues.every(row => row[0] === true);
      if (!allChecked) {
        sheet.getRange('B15').setValue('❌ NOT STARTED');
        sheet.getRange('B16').setValue(new Date().toLocaleString());
        sheet.getRange('B17').setValue('All pre-start checklist items must be confirmed.');
        range.setValue(false);
        return;
      }

      // Run preflight checks
      const preflight = runPreflightChecks();
      if (!preflight.valid) {
        sheet.getRange('B15').setValue('❌ NOT STARTED');
        sheet.getRange('B16').setValue(new Date().toLocaleString());
        sheet.getRange('B17').setValue(preflight.message);
        range.setValue(false);
        return;
      }

      // Check if already running using SelectionStarted config
      const ss = SpreadsheetApp.getActiveSpreadsheet();
      const configSheet = ss.getSheetByName('Config');
      const turnSheet = ss.getSheetByName('Turn Management');
      const turnData = turnSheet.getDataRange().getValues();

      let selectionStarted = false;
      const configData = configSheet.getDataRange().getValues();
      let startedRowIdx = -1;
      for (let i = 0; i < configData.length; i++) {
        if (configData[i][0] === 'SelectionStarted') {
          selectionStarted = (configData[i][1] === true || String(configData[i][1]).toUpperCase() === 'TRUE');
          startedRowIdx = i + 1;
          break;
        }
      }

      if (selectionStarted) {
        sheet.getRange('B15').setValue('❌ NOT STARTED');
        sheet.getRange('B16').setValue(new Date().toLocaleString());
        sheet.getRange('B17').setValue('The selection process has already begun.');
        range.setValue(false);
        return;
      }

      // Start Round 1
      configSheet.getRange('B2').setValue(1);

      if (startedRowIdx !== -1) {
        configSheet.getRange(startedRowIdx, 2).setValue(true);
      } else {
        configSheet.appendRow(['SelectionStarted', true]);
      }

      // Calculate queue to get initial notifications
      const beforeWindow = []; // Empty since we just started
      const currentWindow = calculateQueueWindow(turnData, 1);

      pendingRowIndices = computePendingNotifications(beforeWindow, currentWindow, 1, 1, turnData);

      sheet.getRange('B15').setValue('✅ READY — ROUND 1 STARTED');
      sheet.getRange('B16').setValue(new Date().toLocaleString());
      sheet.getRange('B17').setValue(`${turnData.length - 1} participants validated. Initial notifications queued.`);

      range.setValue(false); // Reset checkbox

    } catch(err) {
      sheet.getRange('B15').setValue('❌ ERROR');
      sheet.getRange('B16').setValue(new Date().toLocaleString());
      sheet.getRange('B17').setValue(err.message);
      range.setValue(false);
    } finally {
      lock.releaseLock();
    }

    // Process SMS outside the lock
    if (pendingRowIndices && pendingRowIndices.length > 0) {
      try {
        _processPendingNotifications(pendingRowIndices);
      } catch (e) {
        console.error("SMS notification processing failed during Round 1 start: " + e.message);
      }
    }
  }
}

function runPreflightChecks() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const turnSheet = ss.getSheetByName('Turn Management');
  const weekSheet = ss.getSheetByName('Week Availability');
  const configSheet = ss.getSheetByName('Config');

  if (!turnSheet) return { valid: false, message: 'Turn Management sheet is missing.' };
  if (!weekSheet) return { valid: false, message: 'Week Availability sheet is missing.' };
  if (!configSheet) return { valid: false, message: 'Config sheet is missing.' };

  const turnData = turnSheet.getDataRange().getValues();
  if (turnData.length < 2) return { valid: false, message: 'Turn Management has no participants.' };

  const headers = turnData[0];

  // Required columns
  const required = ['Name', 'PIN', 'SeniorityPosition', 'Status', 'WeeksSelected', 'LotteryPosition', 'SkipNextTurn'];
  for (let req of required) {
    if (headers.indexOf(req) === -1) {
      return { valid: false, message: `Missing required column in Turn Management: ${req}` };
    }
  }

  const nameIdx = headers.indexOf('Name');
  const pinIdx = headers.indexOf('PIN');
  const senIdx = headers.indexOf('SeniorityPosition');
  const lotIdx = headers.indexOf('LotteryPosition');

  const names = new Set();
  const senPositions = new Set();
  const lotPositions = new Set();

  for (let i = 1; i < turnData.length; i++) {
    const row = turnData[i];
    const name = String(row[nameIdx] || '').trim();
    if (!name) return { valid: false, message: `Row ${i+1} has a blank name.` };
    if (names.has(name)) return { valid: false, message: `Duplicate name found: ${name}` };
    names.add(name);

    const pin = String(row[pinIdx] || '').trim();
    if (!pin) return { valid: false, message: `Participant ${name} is missing a PIN.` };

    const sen = row[senIdx];
    if (sen === '' || sen === null || sen === undefined) return { valid: false, message: `SeniorityPosition is blank for ${name}.` };
    if (senPositions.has(sen)) return { valid: false, message: `SeniorityPosition ${sen} is duplicated.` };
    senPositions.add(sen);

    const lot = row[lotIdx];
    if (lot === '' || lot === null || lot === undefined) return { valid: false, message: `LotteryPosition is blank for ${name}.` };
    if (lotPositions.has(lot)) return { valid: false, message: `LotteryPosition ${lot} is duplicated.` };
    lotPositions.add(lot);
  }

  // Week Availability Checks
  const weekData = weekSheet.getDataRange().getValues();
  if (weekData.length < 2) return { valid: false, message: 'No weeks defined in Week Availability.' };

  const wHeaders = weekData[0];
  if (wHeaders[0] !== 'WeekStartDate' || wHeaders[1] !== 'Classification' || wHeaders[6] !== 'SpotsRemaining') {
    return { valid: false, message: 'Week Availability headers are incorrect.' };
  }

  for (let i = 1; i < weekData.length; i++) {
    const row = weekData[i];
    const date = row[0];
    if (!date) return { valid: false, message: `Row ${i+1} in Week Availability is missing a Date.` };

    const cls = normalizeClassification(row[1]);
    if (!cls) return { valid: false, message: `Row ${i+1} has an invalid classification.` };

    const spots = row[6];
    if (spots === '' || spots === null || isNaN(spots) || spots < 0 || spots > 4) {
      return { valid: false, message: `Row ${i+1} has an invalid SpotsRemaining value.` };
    }
  }

  // Conditional SMS validation
  if (isSmsEnabled()) {
    const smsCheck = checkSmsConfiguration();
    if (!smsCheck.valid) return smsCheck;

    const phoneIdx = headers.indexOf('PhoneNumber');
    if (phoneIdx === -1) return { valid: false, message: 'PhoneNumber column is missing but SMS is enabled.' };

    for (let i = 1; i < turnData.length; i++) {
      const phone = String(turnData[i][phoneIdx] || '').trim();
      if (!phone) return { valid: false, message: `Participant ${turnData[i][nameIdx]} is missing a phone number, but SMS is enabled.` };
    }
  }

  return { valid: true };
}

/**
 * Shared helper to safely transition the system from Round 1 to Round 2.
 * Validates that all participants have finished Round 1 and have unique lottery positions.
 * @returns {object} { success: boolean, message: string }
 */
function _transitionToRound2(turnSheet, configSheet, turnDataRaw) {
  const headers = turnDataRaw[0];
  const statusIdx = headers.indexOf('Status');
  const lotPosIdx = headers.indexOf('LotteryPosition');

  if (statusIdx === -1) return { success: false, message: "Status column missing." };

  // Verify everyone is completed
  let isRound1Over = true;
  for (let i = 1; i < turnDataRaw.length; i++) {
    if (turnDataRaw[i][statusIdx] !== 'Completed') {
      isRound1Over = false;
      break;
    }
  }

  if (!isRound1Over) {
    return { success: false, message: "Round 1 is not fully complete. No action taken." };
  }

  // Verify lottery readiness
  if (!checkLotteryReady(turnDataRaw)) {
    return { success: false, message: "LotteryPosition is not correctly populated. Must have exactly one unique value per participant." };
  }

  // Validation passed, perform state transitions
  configSheet.getRange("B2").setValue(2);

  for (let i = 1; i < turnDataRaw.length; i++) {
    turnSheet.getRange(i + 1, statusIdx + 1).setValue('Waiting');
  }

  return { success: true, message: "Lottery Round 2 Initialized Successfully." };
}

function testAdminControlMissingCheckboxes() {
  let ss;
  const originalGetActive = SpreadsheetApp.getActiveSpreadsheet;
  try {
    ss = setupMockSpreadsheet();
    SpreadsheetApp.getActiveSpreadsheet = () => ss;

    // Create Admin Control
    setupAdminControl(true);

    const adminSheet = ss.getSheetByName('Admin Control');

    // Uncheck one requirement
    adminSheet.getRange('A5').setValue(false);

    // Trigger start
    adminSheet.getRange('A12').setValue(true);
    adminControlOnEdit({ range: adminSheet.getRange('A12') });

    const status = adminSheet.getRange('B15').getValue();
    if (status !== '❌ NOT STARTED') throw new Error("Should not start if checklist is incomplete.");

    console.log("PASS: testAdminControlMissingCheckboxes");
  } finally {
    SpreadsheetApp.getActiveSpreadsheet = originalGetActive;
    if (ss) {
      const files = DriveApp.getFilesByName(ss.getName());
      while (files.hasNext()) files.next().setTrashed(true);
    }
  }
}

function testAdminControlAlreadyRunning() {
  let ss;
  const originalGetActive = SpreadsheetApp.getActiveSpreadsheet;
  try {
    ss = setupMockSpreadsheet();
    SpreadsheetApp.getActiveSpreadsheet = () => ss;
    setupSpreadsheetSchema(); // Need skipNextTurn etc

    // Make it look running using persistent state
    const configSheet = ss.getSheetByName('Config');
    configSheet.getRange('B3').setValue(true);

    setupAdminControl(true);
    const adminSheet = ss.getSheetByName('Admin Control');
    adminSheet.getRange('A5:A9').setValue(true); // Check all
    adminSheet.getRange('A12').setValue(true); // Trigger

    adminControlOnEdit({ range: adminSheet.getRange('A12') });

    const status = adminSheet.getRange('B15').getValue();
    if (status !== '❌ NOT STARTED') throw new Error("Should not start if already running.");
    if (adminSheet.getRange('B17').getValue().indexOf('already begun') === -1) throw new Error("Wrong error message for already running.");

    console.log("PASS: testAdminControlAlreadyRunning");
  } finally {
    SpreadsheetApp.getActiveSpreadsheet = originalGetActive;
    if (ss) {
      const files = DriveApp.getFilesByName(ss.getName());
      while (files.hasNext()) files.next().setTrashed(true);
    }
  }
}

function testAdminControlSuccessfulStart() {
  let ss;
  const originalGetActive = SpreadsheetApp.getActiveSpreadsheet;
  try {
    ss = setupMockSpreadsheet();
    SpreadsheetApp.getActiveSpreadsheet = () => ss;
    setupSpreadsheetSchema(); // Need skipNextTurn etc

    // Turn off SMS to not fail on missing phone numbers
    const props = PropertiesService.getScriptProperties();
    props.setProperty('SMS_NOTIFICATIONS_ENABLED', 'false');

    setupAdminControl(true);
    const adminSheet = ss.getSheetByName('Admin Control');
    adminSheet.getRange('A5:A9').setValue(true); // Check all
    adminSheet.getRange('A12').setValue(true); // Trigger

    adminControlOnEdit({ range: adminSheet.getRange('A12') });

    const status = adminSheet.getRange('B15').getValue();
    if (status !== '✅ READY — ROUND 1 STARTED') throw new Error("Failed to start Round 1. Status: " + status);
    if (adminSheet.getRange('A12').getValue() === true) throw new Error("Checkbox should be reset");

    console.log("PASS: testAdminControlSuccessfulStart");
  } finally {
    SpreadsheetApp.getActiveSpreadsheet = originalGetActive;
    if (ss) {
      const files = DriveApp.getFilesByName(ss.getName());
      while (files.hasNext()) files.next().setTrashed(true);
    }
  }
}

function testAutomaticRound2Transition() {
  let ss;
  const originalGetActive = SpreadsheetApp.getActiveSpreadsheet;
  const props = PropertiesService.getScriptProperties();
  const originalSms = props.getProperty('SMS_NOTIFICATIONS_ENABLED');
  try {
    props.setProperty('SMS_NOTIFICATIONS_ENABLED', 'false'); // Disable SMS

    ss = setupMockSpreadsheet();
    SpreadsheetApp.getActiveSpreadsheet = () => ss;
    setupSpreadsheetSchema(); // Get all columns

    const configSheet = ss.getSheetByName('Config');
    const turnSheet = ss.getSheetByName('Turn Management');
    const weekSheet = ss.getSheetByName('Week Availability');

    // Enable started state for transition tests
    configSheet.getRange('B3').setValue(true);

    // Everyone except Person5 is Complete
    for (let i = 2; i <= 5; i++) {
        turnSheet.getRange(i, 4).setValue('Completed'); // Status
    }

    // Simulate Person5 finishing
    const res = processSelection({ name: 'Person5', week1: weekSheet.getRange(2, 1).getValue().getTime() });

    if (!res.success) throw new Error("Selection should have succeeded: " + res.message);

    const currentRound = configSheet.getRange('B2').getValue();
    if (currentRound !== 2) throw new Error("Should have automatically transitioned to Round 2. Instead in round " + currentRound);

    const p1Status = turnSheet.getRange(2, 4).getValue();
    if (p1Status !== 'Waiting') throw new Error("Statuses should be reset to Waiting.");

    console.log("PASS: testAutomaticRound2Transition");
  } finally {
    if (originalSms) {
      props.setProperty('SMS_NOTIFICATIONS_ENABLED', originalSms);
    } else {
      props.deleteProperty('SMS_NOTIFICATIONS_ENABLED');
    }
    SpreadsheetApp.getActiveSpreadsheet = originalGetActive;
    if (ss) {
      const files = DriveApp.getFilesByName(ss.getName());
      while (files.hasNext()) files.next().setTrashed(true);
    }
  }
}

function runAdminTests() {
  testAdminControlMissingCheckboxes();
  testAdminControlAlreadyRunning();
  testAdminControlSuccessfulStart();
  testAutomaticRound2Transition();
  testAdminControlMissingLottery();
  testAdminControlRepeatedTaps();
  testAutomaticRound2FailedSafely();
  testAutomaticRound2SmsTransition();
  testSelectionStartedEnforcement();
}

function testAdminControlMissingLottery() {
  let ss;
  const originalGetActive = SpreadsheetApp.getActiveSpreadsheet;
  try {
    ss = setupMockSpreadsheet();
    SpreadsheetApp.getActiveSpreadsheet = () => ss;
    setupSpreadsheetSchema(); // Need skipNextTurn etc

    // Mess up lottery position
    const turnSheet = ss.getSheetByName('Turn Management');
    turnSheet.getRange(2, 6).setValue(''); // Blank LotteryPosition for Person1

    // Turn off SMS to not fail on missing phone numbers
    const props = PropertiesService.getScriptProperties();
    props.setProperty('SMS_NOTIFICATIONS_ENABLED', 'false');

    setupAdminControl(true);
    const adminSheet = ss.getSheetByName('Admin Control');
    adminSheet.getRange('A5:A9').setValue(true); // Check all
    adminSheet.getRange('A12').setValue(true); // Trigger

    adminControlOnEdit({ range: adminSheet.getRange('A12') });

    const status = adminSheet.getRange('B15').getValue();
    const details = adminSheet.getRange('B17').getValue();

    if (status !== '❌ NOT STARTED') throw new Error("Should not start with missing lottery position.");
    if (details.indexOf('LotteryPosition is blank') === -1) throw new Error("Missing correct error message. Got: " + details);

    // Fix it, then make it duplicate
    turnSheet.getRange(2, 6).setValue(2);
    adminSheet.getRange('A12').setValue(true);
    adminControlOnEdit({ range: adminSheet.getRange('A12') });

    const statusDup = adminSheet.getRange('B15').getValue();
    const detailsDup = adminSheet.getRange('B17').getValue();

    if (statusDup !== '❌ NOT STARTED') throw new Error("Should not start with duplicate lottery position.");
    if (detailsDup.indexOf('is duplicated') === -1) throw new Error("Missing correct error message for duplicate. Got: " + detailsDup);

    console.log("PASS: testAdminControlMissingLottery");
  } finally {
    SpreadsheetApp.getActiveSpreadsheet = originalGetActive;
    if (ss) {
      const files = DriveApp.getFilesByName(ss.getName());
      while (files.hasNext()) files.next().setTrashed(true);
    }
  }
}

function testAdminControlRepeatedTaps() {
  let ss;
  const originalGetActive = SpreadsheetApp.getActiveSpreadsheet;
  try {
    ss = setupMockSpreadsheet();
    SpreadsheetApp.getActiveSpreadsheet = () => ss;
    setupSpreadsheetSchema(); // Need skipNextTurn etc

    // Turn off SMS to not fail on missing phone numbers
    const props = PropertiesService.getScriptProperties();
    props.setProperty('SMS_NOTIFICATIONS_ENABLED', 'false');

    setupAdminControl(true);
    const adminSheet = ss.getSheetByName('Admin Control');
    adminSheet.getRange('A5:A9').setValue(true); // Check all
    adminSheet.getRange('A12').setValue(true); // Trigger

    // First tap starts it
    adminControlOnEdit({ range: adminSheet.getRange('A12') });

    const status = adminSheet.getRange('B15').getValue();
    if (status !== '✅ READY — ROUND 1 STARTED') throw new Error("Failed to start Round 1 initially");

    // Next tap should not reset anything
    adminSheet.getRange('A12').setValue(true);
    adminControlOnEdit({ range: adminSheet.getRange('A12') });

    const status2 = adminSheet.getRange('B15').getValue();
    if (status2 !== '❌ NOT STARTED') throw new Error("Repeated tap should be blocked as 'NOT STARTED'. Status: " + status2);

    console.log("PASS: testAdminControlRepeatedTaps");
  } finally {
    SpreadsheetApp.getActiveSpreadsheet = originalGetActive;
    if (ss) {
      const files = DriveApp.getFilesByName(ss.getName());
      while (files.hasNext()) files.next().setTrashed(true);
    }
  }
}

function testAutomaticRound2FailedSafely() {
  let ss;
  const originalGetActive = SpreadsheetApp.getActiveSpreadsheet;
  const props = PropertiesService.getScriptProperties();
  const originalSms = props.getProperty('SMS_NOTIFICATIONS_ENABLED');
  try {
    props.setProperty('SMS_NOTIFICATIONS_ENABLED', 'false'); // Disable SMS

    ss = setupMockSpreadsheet();
    SpreadsheetApp.getActiveSpreadsheet = () => ss;
    setupSpreadsheetSchema(); // Get all columns

    const configSheet = ss.getSheetByName('Config');
    const turnSheet = ss.getSheetByName('Turn Management');
    const weekSheet = ss.getSheetByName('Week Availability');

    // Make LotteryPosition duplicated to fail transition validation
    turnSheet.getRange(2, 6).setValue(2);

    // Everyone except Person5 is Complete
    for (let i = 2; i <= 5; i++) {
        turnSheet.getRange(i, 4).setValue('Completed'); // Status
    }

    // Simulate Person5 finishing
    const res = processSelection({ name: 'Person5', week1: weekSheet.getRange(2, 1).getValue().getTime() });

    if (!res.success) throw new Error("Selection should have succeeded even if transition failed: " + res.message);
    if (res.message.indexOf('Could not auto-start') === -1) throw new Error("Should notify that auto-start failed.");

    const currentRound = configSheet.getRange('B2').getValue();
    if (currentRound !== 1) throw new Error("Should have stayed in round 1 because of validation failure. Currently in round: " + currentRound);

    console.log("PASS: testAutomaticRound2FailedSafely");
  } finally {
    if (originalSms) {
      props.setProperty('SMS_NOTIFICATIONS_ENABLED', originalSms);
    } else {
      props.deleteProperty('SMS_NOTIFICATIONS_ENABLED');
    }
    SpreadsheetApp.getActiveSpreadsheet = originalGetActive;
    if (ss) {
      const files = DriveApp.getFilesByName(ss.getName());
      while (files.hasNext()) files.next().setTrashed(true);
    }
  }
}


function testAutomaticRound2SmsTransition() {
  let ss;
  const originalGetActive = SpreadsheetApp.getActiveSpreadsheet;
  const props = _smsDependencies.getProperties;
  const originalFetch = _smsDependencies.fetch;

  try {
    // Enable SMS, and mock config for successful validation
    _smsDependencies.getProperties = () => ({
      SMS_NOTIFICATIONS_ENABLED: 'true',
      TWILIO_ACCOUNT_SID: '123',
      TWILIO_AUTH_TOKEN: '456',
      TWILIO_FROM_NUMBER: '789',
      VACATION_SELECTOR_URL: 'http'
    });

    // Mock the external network call to succeed and not throw
    _smsDependencies.fetch = () => {
      return {
        getResponseCode: () => 200,
        getContentText: () => JSON.stringify({ sid: 'SM123' })
      };
    };

    ss = setupMockSpreadsheet();
    SpreadsheetApp.getActiveSpreadsheet = () => ss;
    setupSpreadsheetSchema(); // Need skipNextTurn, Notifications sheet, etc.

    const configSheet = ss.getSheetByName('Config');
    const turnSheet = ss.getSheetByName('Turn Management');
    const weekSheet = ss.getSheetByName('Week Availability');
    const logSheet = ss.getSheetByName('Notification Log');

    // Set Phone numbers for test
    const turnData = turnSheet.getDataRange().getValues();
    const phoneIdx = turnData[0].indexOf('PhoneNumber');
    for (let i = 1; i <= 5; i++) {
        turnSheet.getRange(i + 1, phoneIdx + 1).setValue('555-1234');
    }

    // Enable started state for transition tests
    configSheet.getRange('B3').setValue(true);

    // Everyone except Person5 is Complete for Round 1
    for (let i = 2; i <= 5; i++) {
        turnSheet.getRange(i, 4).setValue('Completed'); // Status
    }

    // Wipe any existing logs (from other tests modifying mock sheet)
    if (logSheet.getLastRow() > 1) {
      logSheet.getRange(2, 1, logSheet.getLastRow() - 1, logSheet.getLastColumn()).clearContent();
    }

    // Simulate Person5 finishing Round 1
    const res = processSelection({ name: 'Person5', week1: weekSheet.getRange(2, 1).getValue().getTime() });

    if (!res.success) throw new Error("Selection should have succeeded: " + res.message);

    const currentRound = configSheet.getRange('B2').getValue();
    if (currentRound !== 2) throw new Error("Should have automatically transitioned to Round 2.");

    // Verify Notification Log row generation
    // Since Round 2 sorts by lottery position, people 1, 2, and 3 should be in the window
    // and thus exactly 3 logs should exist.
    const newLogs = logSheet.getDataRange().getValues();
    // length is 4 (header + 3 logs)
    if (newLogs.length !== 4) throw new Error("Expected exactly 3 notifications, got " + (newLogs.length - 1));

    const roundIdx = newLogs[0].indexOf('Round');
    const dedupeIdx = newLogs[0].indexOf('DedupeKey');
    const statusIdx = newLogs[0].indexOf('Status');

    for (let i = 1; i < newLogs.length; i++) {
      if (newLogs[i][roundIdx] !== 2) throw new Error("Log has wrong round number: " + newLogs[i][roundIdx]);
      if (newLogs[i][dedupeIdx].indexOf('ROUND:2') === -1) throw new Error("Log has wrong dedupe key: " + newLogs[i][dedupeIdx]);
      if (newLogs[i][statusIdx] !== 'SENT') throw new Error("Log status is not SENT: " + newLogs[i][statusIdx]);
    }

    // Verify no duplicates created when evaluating round again
    const windowRaw = turnSheet.getDataRange().getValues();
    const window = calculateQueueWindow(windowRaw, 2);
    const newIndices = computePendingNotifications([], window, 2, 2, windowRaw);
    if (newIndices.length !== 0) throw new Error("Duplicate notifications were queued!");

    console.log("PASS: testAutomaticRound2SmsTransition");
  } finally {
    _smsDependencies.getProperties = props;
    _smsDependencies.fetch = originalFetch;
    SpreadsheetApp.getActiveSpreadsheet = originalGetActive;
    if (ss) {
      const files = DriveApp.getFilesByName(ss.getName());
      while (files.hasNext()) files.next().setTrashed(true);
    }
  }
}


/**
 * Shared helper to check if the selection process has formally started
 */
function isSelectionStarted() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  if (!configSheet) return false;

  const configData = configSheet.getDataRange().getValues();
  for (let i = 0; i < configData.length; i++) {
    if (configData[i][0] === 'SelectionStarted') {
      return (configData[i][1] === true || String(configData[i][1]).toUpperCase() === 'TRUE');
    }
  }
  return false;
}

function testSelectionStartedEnforcement() {
  let ss;
  const originalGetActive = SpreadsheetApp.getActiveSpreadsheet;
  try {
    ss = setupMockSpreadsheet();
    SpreadsheetApp.getActiveSpreadsheet = () => ss;

    // Explicitly set SelectionStarted to FALSE
    const configSheet = ss.getSheetByName('Config');
    configSheet.getRange('B3').setValue(false);

    const weekSheet = ss.getSheetByName('Week Availability');
    const weekTime = weekSheet.getRange(2, 1).getValue().getTime();

    // Test 1: Should fail
    let res = processSelection({ name: 'Person1', week1: weekTime });
    if (res.success) throw new Error("Selection should have failed when SelectionStarted = FALSE.");
    if (res.message.indexOf("not started yet") === -1) throw new Error("Wrong error message: " + res.message);

    // Set SelectionStarted to TRUE
    configSheet.getRange('B3').setValue(true);

    // Test 2: Should succeed
    res = processSelection({ name: 'Person1', week1: weekTime });
    if (!res.success) throw new Error("Selection should have succeeded when SelectionStarted = TRUE. " + res.message);

    console.log("PASS: testSelectionStartedEnforcement");
  } finally {
    SpreadsheetApp.getActiveSpreadsheet = originalGetActive;
    if (ss) {
      const files = DriveApp.getFilesByName(ss.getName());
      while (files.hasNext()) files.next().setTrashed(true);
    }
  }
}
