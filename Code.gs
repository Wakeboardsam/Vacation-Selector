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
  // Now row schema: Date, Classif, MaxCap, SpotsRemaining, Special, AssignedTo
  if (rows.length === 0) return [];
  return rows.slice(1)
    .filter(row => Number(row[3]) > 0)
    .map(row => ({
      displayDate: row[0] instanceof Date ? row[0].toLocaleDateString("en-US", { timeZone: "UTC", month: 'short', day: 'numeric' }) : String(row[0]),
      valueDate: row[0] instanceof Date ? row[0].getTime() : null,
      classification: normalizeClassification(row[1]),
      spotsRemaining: Number(row[3]),
      specialWeek: row[4],
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
  const sheet = ss.getSheetByName('Participant Config');
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const nameIdx = data[0].indexOf('Name');
  if (nameIdx === -1) return [];
  const names = [];
  for (let i = 1; i < data.length; i++) {
     if (data[i][nameIdx]) names.push(data[i][nameIdx]);
  }
  return names;
}
function verifyUser(name, pin) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Participant Config');
  if (!sheet) return { success: false, message: "Participant Config not found." };

  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const nameIdx = headers.indexOf('Name');
  const pinIdx = headers.indexOf('PIN');
  const ackYearIdx = headers.indexOf('Rules Acknowledged Year');

  const userRow = data.find(row => row[nameIdx] === name);
  if (userRow) {
    if (String(userRow[pinIdx]) === String(pin)) {
        const activeYear = _getConfigValue('ActiveYear', new Date().getFullYear());
        const hasAcknowledged = String(userRow[ackYearIdx]) === String(activeYear);
        return { success: true, message: "Success", needsAcknowledgment: !hasAcknowledged };
    }
  }
  return { success: false, message: "Invalid PIN" };
}

function _getActiveWindowSizeForPhase(phase) {
  const options = getAdminOptions();
  if (phase === 'VACATION_SENIORITY' || phase === 'VACATION_RANDOM') return parseInt(options['VACATION_WINDOW_SIZE'] || 3);
  if (phase === 'WEEKEND') return parseInt(options['WEEKEND_WINDOW_SIZE'] || 2);
  if (phase === 'HOLIDAY_VOLUNTEER' || phase === 'HOLIDAY_MANDATORY') return parseInt(options['HOLIDAY_WINDOW_SIZE'] || 2);
  if (phase === 'TRANSFER_RECEIVER') return parseInt(options['TRANSFER_WINDOW_SIZE'] || 2);
  return 1;
}

function calculateQueueWindow(turnDataRaw) {
    const turnData = [...turnDataRaw];
    const turnHeaders = turnData.shift();

    const nameIdx = turnHeaders.indexOf('Name');
    const statusIdx = turnHeaders.indexOf('Status');
    const senPosIdx = turnHeaders.indexOf('SeniorityPosition');
    const lotPosIdx = turnHeaders.indexOf('LotteryPosition');
    const skipIdx = turnHeaders.indexOf('SkipNextTurn');

    const currentPhase = _getConfigValue('CurrentPhase', 'SETUP_EMPTY');
    const currentDirection = _getConfigValue('CurrentDirection', 'ASCENDING');
    const windowSize = _getActiveWindowSizeForPhase(currentPhase);

    let queue = turnData.map((row, index) => ({
        originalRowIndex: index + 2,
        name: row[nameIdx],
        status: String(row[statusIdx] || '').trim(),
        skipNextTurn: row[skipIdx],
        seniorityPosition: row[senPosIdx] !== '' ? Number(row[senPosIdx]) : 9999,
        lotteryPosition: row[lotPosIdx] !== '' ? Number(row[lotPosIdx]) : 9999,
        computedStatus: 'Waiting'
    }));

    if (currentPhase === 'VACATION_SENIORITY') {
        queue.sort((a, b) => a.seniorityPosition - b.seniorityPosition);
    } else {
        if (currentDirection === 'ASCENDING') queue.sort((a, b) => a.lotteryPosition - b.lotteryPosition);
        else queue.sort((a, b) => b.lotteryPosition - a.lotteryPosition);
    }

    // Mandatory Phase 3-Tier Filter
    if (currentPhase === 'HOLIDAY_MANDATORY') {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        const participantSheet = ss.getSheetByName('Participant Config');
        if (participantSheet) {
            const hData = ss.getSheetByName('Holiday Coverage').getDataRange().getValues();
            let remaining = 0;
            for (let i = 1; i < hData.length; i++) { if (hData[i][3] === '') remaining++; }
            if (remaining > 0) {
                const partData = participantSheet.getDataRange().getValues();
                let currentCounts = {};
                for (let i = 1; i < hData.length; i++) {
                    let p = hData[i][3];
                    if (p) currentCounts[p] = (currentCounts[p] || 0) + 1;
                }
                let tier1 = [], tier2 = [], tier3 = [];
                for (let p of queue) {
                    if (p.status === 'Completed' || p.status === 'Pass' || p.status === 'TargetReached' || p.status === 'None') continue;
                    let row = partData.find(r => r[partData[0].indexOf('Name')] === p.name);
                    if (!row || row[partData[0].indexOf('Mandatory Holiday Eligible')] !== true) {
                        p.computedStatus = 'Skipping'; continue;
                    }
                    let hasPrior = row[partData[0].indexOf('Worked Any Official Holiday Last Year')] === true;
                    let hasCurrent = (currentCounts[p.name] || 0) > 0;
                    tier3.push(p);
                    if (!hasCurrent) {
                        if (!hasPrior) tier1.push(p); else tier2.push(p);
                    }
                }
                let activeTier = tier3;
                if (tier1.length > 0) activeTier = tier1;
                else if (tier2.length > 0) activeTier = tier2;

                const activeNames = new Set(activeTier.map(p => p.name));
                for (let p of queue) {
                    if (p.computedStatus !== 'Completed' && p.computedStatus !== 'Pass' && p.computedStatus !== 'None' && p.computedStatus !== 'TargetReached' && !activeNames.has(p.name)) {
                        p.computedStatus = 'Skipping';
                    }
                }
            }
        }
    }

    let activeCount = 0;
    for (let i = 0; i < queue.length; i++) {
        let person = queue[i];
        if (person.status === 'Completed' || person.status === 'Pass' || person.status === 'TargetReached' || person.status === 'None') {
            person.computedStatus = 'Completed';
            continue;
        }
        if (person.skipNextTurn === true || person.computedStatus === 'Skipping') {
            person.computedStatus = 'Skipping';
            continue;
        }
        if (activeCount < windowSize) {
            person.computedStatus = 'Active';
            activeCount++;
        }
    }

    return queue;
}

function _advanceQueueDirectionIfComplete() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const turnSheet = ss.getSheetByName('Turn Management');
    const turnDataRaw = turnSheet.getDataRange().getValues();
    const currentPhase = _getConfigValue('CurrentPhase', '');

    if (currentPhase === 'VACATION_SENIORITY' || currentPhase === 'TRANSFER_OFFER_COLLECTION') return false;

    const queue = calculateQueueWindow(turnDataRaw);
    const hasUnfinished = queue.some(p => p.computedStatus === 'Active' || p.computedStatus === 'Waiting');

    if (!hasUnfinished) {
        const currentDirection = _getConfigValue('CurrentDirection', 'ASCENDING');
        const nextDirection = currentDirection === 'ASCENDING' ? 'DESCENDING' : 'ASCENDING';

        if (currentPhase === 'VACATION_RANDOM') {
            const currentRound = _getConfigValue('CurrentRound', 1);
            _setConfigValue('CurrentRound', currentRound + 1);
        }
        _setConfigValue('CurrentDirection', nextDirection);

        const statusIdx = turnDataRaw[0].indexOf('Status');
        const skipIdx = turnDataRaw[0].indexOf('SkipNextTurn');

        for (let i = 1; i < turnDataRaw.length; i++) {
            let rowIdx = i + 1;
            let currentStatus = turnDataRaw[i][statusIdx];
            let currentSkip = turnDataRaw[i][skipIdx];

            if (currentSkip === true) {
                turnSheet.getRange(rowIdx, skipIdx + 1).setValue(false);
            }
            if (currentStatus === 'Pass' || currentStatus === 'TargetReached' || currentStatus === 'None') {
                // Permanent
            } else if (currentStatus === 'Completed') {
                turnSheet.getRange(rowIdx, statusIdx + 1).setValue('');
            }
        }
        return true;
    }
    return false;
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
      console.error("SMS notification processing failed: " + e.message);
    }
  }
  return finalResult;
}

function _processSelectionCore(selectionData) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const currentPhase = _getConfigValue('CurrentPhase', 'SETUP_EMPTY');

  if (currentPhase === 'SETUP_EMPTY' || currentPhase === 'SETUP_REVIEW' || currentPhase === 'SETUP_CONFIRMED') {
      return { coreResult: { success: false, message: "System is in Setup phase. Selection is closed." } };
  }
  if (currentPhase === 'COMPLETE') {
      return { coreResult: { success: false, message: "Selection is complete for the year." } };
  }

  const turnSheet = ss.getSheetByName('Turn Management');
  const participantSheet = ss.getSheetByName('Participant Config');
  const beforeWindowRaw = turnSheet.getDataRange().getValues();
  const beforeWindow = calculateQueueWindow(beforeWindowRaw);

  let coreResult = null;

  if (currentPhase === 'VACATION_SENIORITY' || currentPhase === 'VACATION_RANDOM') {
      coreResult = _processVacationSelection(selectionData, ss, turnSheet, participantSheet);
  } else if (currentPhase === 'WEEKEND') {
      coreResult = _processWeekendSelection(selectionData, ss, turnSheet, participantSheet);
  } else if (currentPhase === 'HOLIDAY_VOLUNTEER' || currentPhase === 'HOLIDAY_MANDATORY') {
      coreResult = _processHolidaySelection(selectionData, ss, turnSheet, participantSheet, currentPhase);
  } else if (currentPhase === 'TRANSFER_RECEIVER') {
      coreResult = _processTransferSelection(selectionData, ss, turnSheet, participantSheet);
  } else if (currentPhase === 'TRANSFER_OFFER_COLLECTION') {
      coreResult = _processTransferOffer(selectionData, ss, turnSheet, participantSheet);
  } else {
      coreResult = { success: false, message: "Unknown phase: " + currentPhase };
  }

  let newIndices = [];
  if (coreResult && coreResult.success) {
      if (currentPhase !== 'TRANSFER_OFFER_COLLECTION') {
         _advanceQueueDirectionIfComplete();
      }
      const afterWindowRaw = turnSheet.getDataRange().getValues();
      const afterWindow = calculateQueueWindow(afterWindowRaw);
      newIndices = computePendingNotifications(beforeWindow, afterWindow, _getConfigValue('CurrentRound', 1), _getActiveWindowSizeForPhase(currentPhase), afterWindowRaw);
  }
  return { coreResult: coreResult, createdRowIndices: newIndices };
}

function _processVacationSelection(selectionData, ss, turnSheet, participantSheet) {
    const { name, week1, week2 } = selectionData;
    const turnDataRaw = turnSheet.getDataRange().getValues();
    const queue = calculateQueueWindow(turnDataRaw);
    const user = queue.find(p => p.name === name);
    if (!user) return { success: false, message: "Participant not found." };
    if (user.computedStatus !== 'Active') return { success: false, message: "It is not your turn to select." };

    const partData = participantSheet.getDataRange().getValues();
    const partHeaders = partData[0];
    const userConf = partData.find(row => row[partHeaders.indexOf('Name')] === name);

    const globalTarget = parseInt(getAdminOptions()['DEFAULT_VACATION_TARGET'] || 9);
    let targetStr = userConf[partHeaders.indexOf('Vacation Week Target Override')];
    const userTarget = targetStr !== '' ? parseInt(targetStr) : globalTarget;

    const hadSpringBreak = userConf[partHeaders.indexOf('Had Spring Break Last Year')] === true;
    const hadChristmas = userConf[partHeaders.indexOf('Had Christmas Week Last Year')] === true;
    const currentRound = _getConfigValue('CurrentRound', 1);

    const weekSheet = ss.getSheetByName('Week Availability');
    const weekDataRaw = weekSheet.getDataRange().getValues();

    let w1Row = null, w2Row = null, w1Idx = -1, w2Idx = -1;
    for (let i = 1; i < weekDataRaw.length; i++) {
        let t = new Date(weekDataRaw[i][0]).getTime();
        if (week1 && t === Number(week1)) { w1Row = weekDataRaw[i]; w1Idx = i; }
        if (week2 && t === Number(week2)) { w2Row = weekDataRaw[i]; w2Idx = i; }
    }

    if (week1 && !w1Row) return { success: false, message: "Selected week 1 not found." };
    if (week2 && !w2Row) return { success: false, message: "Selected week 2 not found." };
    if (!week1 && !week2) return { success: false, message: "No weeks selected." };
    if (week1 && week2 && week1 === week2) return { success: false, message: "Cannot select the same week twice." };

    let countPrime = 0, countNonPrime = 0;
    const checkWeek = (row, idx) => {
        let maxCap = parseInt(row[2]) || parseInt(globalCap);
        let assignedStr = String(row[5] || '');
        let assignedArr = assignedStr ? assignedStr.split(',').map(n => n.trim()) : [];
        let spots = maxCap - assignedArr.length;

        let classif = String(row[1] || '').trim().toLowerCase() === 'prime' ? 'Prime' : 'Non-Prime';
        let special = row[4];

        if (spots <= 0) return "Week has zero spots remaining.";
        if (assignedArr.includes(name)) return "You already hold a spot in this week.";

        if (classif === 'Prime') countPrime++; else countNonPrime++;
        if (currentRound <= 3) {
            if (special === 'Spring Break' && hadSpringBreak) return "You are restricted from selecting Spring Break until Round 4.";
            if (special === 'Christmas' && hadChristmas) return "You are restricted from selecting Christmas until Round 4.";
        }
        return null;
    };

    if (w1Row) { let err = checkWeek(w1Row, w1Idx); if (err) return { success: false, message: err }; }
    if (w2Row) { let err = checkWeek(w2Row, w2Idx); if (err) return { success: false, message: err }; }

    if (countPrime > 1) return { success: false, message: "You may only select ONE Prime week per turn." };
    if (countPrime === 1 && countNonPrime > 0) return { success: false, message: "A Prime selection must stand alone." };
    if (_getConfigValue('CurrentPhase') === 'VACATION_SENIORITY' && (countPrime + countNonPrime) > 1) {
        return { success: false, message: "During the Seniority Round, you may only select EXACTLY ONE week per turn." };
    }

    let totalAssignmentsNow = 0;
    for (let i = 1; i < weekDataRaw.length; i++) {
       let assignedStr = String(weekDataRaw[i][5] || '');
       if (assignedStr) {
           let assignedArr = assignedStr.split(',').map(n => n.trim());
           if (assignedArr.includes(name)) totalAssignmentsNow++;
       }
    }

    let attemptingToTake = (w1Row ? 1 : 0) + (w2Row ? 1 : 0);
    if (totalAssignmentsNow + attemptingToTake > userTarget) {
        return { success: false, message: "Selecting these weeks exceeds your personal target of " + userTarget + " weeks." };
    }

    const assignWeek = (row, idx) => {
        let maxCap = parseInt(row[2]) || parseInt(globalCap);
        let assignedStr = String(row[5] || '');
        let assignedArr = assignedStr ? assignedStr.split(',').map(n => n.trim()) : [];

        if (assignedArr.length < maxCap) {
            assignedArr.push(name);
            weekSheet.getRange(idx + 1, 6).setValue(assignedArr.join(', '));
            weekSheet.getRange(idx + 1, 4).setValue(maxCap - assignedArr.length);
            return true;
        }
        return false;
    };

    if (w1Row && !assignWeek(w1Row, w1Idx)) return { success: false, message: "Concurrency error: Week 1 filled up." };
    if (w2Row && !assignWeek(w2Row, w2Idx)) return { success: false, message: "Concurrency error: Week 2 filled up." };

    let userRowIdx = turnSheet.getDataRange().getValues().findIndex(row => row[0] === name) + 1;
    let turnStatusCol = turnSheet.getDataRange().getValues()[0].indexOf('Status') + 1;
    let skipTurnCol = turnSheet.getDataRange().getValues()[0].indexOf('SkipNextTurn') + 1;

    turnSheet.getRange(userRowIdx, turnStatusCol).setValue((totalAssignmentsNow + attemptingToTake >= userTarget) ? 'TargetReached' : 'Completed');
    if (countNonPrime === 2) turnSheet.getRange(userRowIdx, skipTurnCol).setValue(true);

    if (_getConfigValue('CurrentPhase') === 'VACATION_SENIORITY') {
       let allComplete = true;
       const turnDataNow = turnSheet.getDataRange().getValues();
       const statIdx = turnDataNow[0].indexOf('Status');
       for (let i = 1; i < turnDataNow.length; i++) {
           let st = String(turnDataNow[i][statIdx]);
           if (st !== 'Completed' && st !== 'TargetReached') { allComplete = false; break; }
       }
       if (allComplete) {
           _setConfigValue('CurrentPhase', 'VACATION_RANDOM');
           _setConfigValue('CurrentRound', 2);
           _resetQueueToLotteryPosition1();
       }
    }
    return { success: true, message: "Selection recorded successfully." };
}

function _processWeekendSelection(selectionData, ss, turnSheet, participantSheet) {
    const { name, weekendDate, weekendDay, optHolidayName, optHolidayDate, optHolidayCall } = selectionData;
    const turnDataRaw = turnSheet.getDataRange().getValues();
    const queue = calculateQueueWindow(turnDataRaw);
    const user = queue.find(p => p.name === name);
    if (!user) return { success: false, message: "Participant not found." };
    if (user.computedStatus !== 'Active') return { success: false, message: "It is not your turn to select." };

    const partData = participantSheet.getDataRange().getValues();
    const partHeaders = partData[0];
    const userConf = partData.find(row => row[partHeaders.indexOf('Name')] === name);
    if (userConf[partHeaders.indexOf('Weekend Phase Enabled')] !== true) return { success: false, message: "You are not enabled for the Weekend Phase." };

    let maxStr = userConf[partHeaders.indexOf('Weekend Assignment Maximum')];
    let hasCap = maxStr !== '';
    let maxCap = hasCap ? parseInt(maxStr) : 9999;

    const weekendSheet = ss.getSheetByName('Weekend Coverage');
    const weekendDataRaw = weekendSheet.getDataRange().getValues();
    const wDateIdx = weekendDataRaw[0].indexOf('Date');
    const wDayIdx = weekendDataRaw[0].indexOf('Day of Week');
    const wPartIdx = weekendDataRaw[0].indexOf('Participant');

    let userWeekendCount = 0;
    for (let i = 1; i < weekendDataRaw.length; i++) {
        if (weekendDataRaw[i][wPartIdx] === name) userWeekendCount++;
    }
    if (userWeekendCount >= maxCap) return { success: false, message: "You have reached your Weekend Assignment Maximum." };
    if (!weekendDate || !weekendDay) return { success: false, message: "Missing weekend selection data." };

    let targetRowIdx = -1;
    let targetDateEpoch = Number(weekendDate);
    for (let i = 1; i < weekendDataRaw.length; i++) {
        if (new Date(weekendDataRaw[i][wDateIdx]).getTime() === targetDateEpoch && weekendDataRaw[i][wDayIdx] === weekendDay) {
            targetRowIdx = i; break;
        }
    }

    if (targetRowIdx === -1) return { success: false, message: "Selected weekend position not found." };
    if (weekendDataRaw[targetRowIdx][wPartIdx] !== '') return { success: false, message: "That position was just selected by another participant. Please choose another available option." };

    let partnerDateEpoch = weekendDay === 'Saturday' ? targetDateEpoch + 86400000 : targetDateEpoch - 86400000;
    for (let i = 1; i < weekendDataRaw.length; i++) {
        if (new Date(weekendDataRaw[i][wDateIdx]).getTime() === partnerDateEpoch && weekendDataRaw[i][wPartIdx] === name) {
            return { success: false, message: "You cannot hold both Saturday and Sunday First Call positions for the same weekend." };
        }
    }

    weekendSheet.getRange(targetRowIdx + 1, wPartIdx + 1).setValue(name);
    let resultMsg = "Weekend selection recorded successfully.";

    if (optHolidayName && optHolidayDate && optHolidayCall) {
        const holidaySheet = ss.getSheetByName('Holiday Coverage');
        const holidayData = holidaySheet.getDataRange().getValues();
        let holRowIdx = -1, userHoldsOtherCall = false;

        for (let i = 1; i < holidayData.length; i++) {
            if (holidayData[i][0] === optHolidayName && new Date(holidayData[i][1]).getTime() === Number(optHolidayDate)) {
                if (holidayData[i][2] === optHolidayCall) holRowIdx = i;
                else if (holidayData[i][3] === name) userHoldsOtherCall = true;
            }
        }

        if (userHoldsOtherCall) {
             resultMsg += " However, you already hold the other call position for that holiday, so the holiday assignment was skipped.";
        } else if (holRowIdx !== -1) {
            if (holidayData[holRowIdx][3] === '') {
                holidaySheet.getRange(holRowIdx + 1, 4).setValue(name);
                resultMsg += " The nearby holiday (" + optHolidayName + " " + optHolidayCall + ") was also assigned to you.";
                ss.getSheetByName('Notification Log').appendRow([
                     new Date(), 'HOLIDAY_OPT_' + name + '_' + optHolidayName + '_' + optHolidayDate, name, _getConfigValue('CurrentRound', 1), 'Holiday Opt-in', 'PENDING', '', '', 'HOLIDAY_CONFIRM'
                ]);
            } else {
                resultMsg += " However, the nearby holiday position was taken by someone else and is no longer available.";
            }
        }
    }

    let userRowIdx = turnSheet.getDataRange().getValues().findIndex(row => row[0] === name) + 1;
    let turnStatusCol = turnSheet.getDataRange().getValues()[0].indexOf('Status') + 1;
    turnSheet.getRange(userRowIdx, turnStatusCol).setValue((userWeekendCount + 1 >= maxCap && hasCap) ? 'TargetReached' : 'Completed');
    return { success: true, message: resultMsg };
}

function _processHolidaySelection(selectionData, ss, turnSheet, participantSheet, currentPhase) {
    const { name, holidayName, holidayDate, holidayCall, action } = selectionData;
    const turnDataRaw = turnSheet.getDataRange().getValues();
    const queue = calculateQueueWindow(turnDataRaw);
    const user = queue.find(p => p.name === name);
    if (!user) return { success: false, message: "Participant not found." };
    if (user.computedStatus !== 'Active') return { success: false, message: "It is not your turn to select." };

    const partData = participantSheet.getDataRange().getValues();
    const partHeaders = partData[0];
    const userConf = partData.find(row => row[partHeaders.indexOf('Name')] === name);
    let userRowIdx = turnSheet.getDataRange().getValues().findIndex(row => row[0] === name) + 1;
    let turnStatusCol = turnSheet.getDataRange().getValues()[0].indexOf('Status') + 1;

    if (currentPhase === 'HOLIDAY_VOLUNTEER') {
        if (userConf[partHeaders.indexOf('Holiday Volunteer')] !== true) return { success: false, message: "You are not enrolled in the Holiday Volunteer phase." };
        if (action === 'Pass') {
            turnSheet.getRange(userRowIdx, turnStatusCol).setValue('Pass');
            return { success: true, message: "You have passed and are removed from the remainder of the volunteer phase." };
        }
    } else if (currentPhase === 'HOLIDAY_MANDATORY') {
        if (userConf[partHeaders.indexOf('Mandatory Holiday Eligible')] !== true) return { success: false, message: "You are not eligible for Mandatory Holiday assignments." };
        if (action === 'Pass') return { success: false, message: "You cannot pass during the Mandatory phase." };
    }

    if (!holidayName || !holidayDate || !holidayCall) return { success: false, message: "Missing holiday selection data." };

    const holidaySheet = ss.getSheetByName('Holiday Coverage');
    const holidayDataRaw = holidaySheet.getDataRange().getValues();
    let targetRowIdx = -1, userHoldsOtherCall = false, targetRowData = null;

    for (let i = 1; i < holidayDataRaw.length; i++) {
        if (holidayDataRaw[i][0] === holidayName && new Date(holidayDataRaw[i][1]).getTime() === Number(holidayDate)) {
            if (holidayDataRaw[i][2] === holidayCall) { targetRowIdx = i; targetRowData = holidayDataRaw[i]; }
            else if (holidayDataRaw[i][3] === name) { userHoldsOtherCall = true; }
        }
    }

    if (targetRowIdx === -1) return { success: false, message: "Selected holiday position not found." };
    if (targetRowData[3] !== '') return { success: false, message: "That position was just selected by another participant." };
    if (userHoldsOtherCall) return { success: false, message: "You cannot hold both call positions for the same holiday." };

    holidaySheet.getRange(targetRowIdx + 1, 4).setValue(name);
    turnSheet.getRange(userRowIdx, turnStatusCol).setValue('Completed');
    return { success: true, message: "Holiday selection recorded successfully." };
}

function _processTransferOffer(selectionData, ss, turnSheet, participantSheet) {
    const { name, offers } = selectionData;
    const partData = participantSheet.getDataRange().getValues();
    const userConf = partData.find(row => row[partData[0].indexOf('Name')] === name);
    if (!userConf) return { success: false, message: "Participant not found." };
    if (userConf[partData[0].indexOf('Transfer Giver')] !== true) return { success: false, message: "You are not enrolled as a Transfer Giver." };
    if (_getConfigValue('TransferLocked', false)) return { success: false, message: "The transfer pool is locked. No more offers can be submitted." };

    if (offers && offers.length > 0) {
        const transferSheet = ss.getSheetByName('Transfer Offers');
        const nextId = transferSheet.getLastRow();
        const rowsToWrite = [];

        const wSheet = ss.getSheetByName('Weekend Coverage');
        const wData = wSheet.getDataRange().getValues();
        const hSheet = ss.getSheetByName('Holiday Coverage');
        const hData = hSheet.getDataRange().getValues();

        for (let i = 0; i < offers.length; i++) {
            let o = offers[i];
            let isValid = false;
            if (o.type === 'Weekend') {
                let day = o.details.indexOf('Saturday') !== -1 ? 'Saturday' : 'Sunday';
                for (let k = 1; k < wData.length; k++) {
                    if (new Date(wData[k][0]).getTime() === o.dateEpoch && wData[k][1] === day && wData[k][3] === name) {
                        isValid = true; break;
                    }
                }
            } else if (o.type === 'Holiday') {
                for (let k = 1; k < hData.length; k++) {
                    if (new Date(hData[k][1]).getTime() === o.dateEpoch && hData[k][2] === o.details && hData[k][3] === name) {
                        isValid = true; break;
                    }
                }
            }
            if (!isValid) return { success: false, message: "You do not own one or more of the offered assignments: " + o.details };
            rowsToWrite.push(['OFFER-' + (nextId + i), name, o.type, o.dateEpoch, o.details, 'Open', '']);
        }

        transferSheet.getRange(transferSheet.getLastRow() + 1, 1, rowsToWrite.length, 7).setValues(rowsToWrite);
    }
    let userRowIdx = turnSheet.getDataRange().getValues().findIndex(row => row[0] === name) + 1;
    turnSheet.getRange(userRowIdx, turnSheet.getDataRange().getValues()[0].indexOf('Status') + 1).setValue('Pass');
    return { success: true, message: "Transfer offers submitted successfully and are irrevocable." };
}

function _processTransferSelection(selectionData, ss, turnSheet, participantSheet) {
    const { name, offerId, action } = selectionData;
    const turnDataRaw = turnSheet.getDataRange().getValues();
    const queue = calculateQueueWindow(turnDataRaw);
    const user = queue.find(p => p.name === name);
    if (!user) return { success: false, message: "Participant not found." };
    if (user.computedStatus !== 'Active') return { success: false, message: "It is not your turn to select." };

    const partData = participantSheet.getDataRange().getValues();
    if (partData.find(row => row[partData[0].indexOf('Name')] === name)[partData[0].indexOf('Transfer Receiver')] !== true) {
        return { success: false, message: "You are not enrolled as a Transfer Receiver." };
    }

    let userRowIdx = turnSheet.getDataRange().getValues().findIndex(row => row[0] === name) + 1;
    let turnStatusCol = turnSheet.getDataRange().getValues()[0].indexOf('Status') + 1;

    if (action === 'None') {
        turnSheet.getRange(userRowIdx, turnStatusCol).setValue('None');
        return { success: true, message: "You have chosen None and are removed from the remainder of the transfer round." };
    }

    if (!offerId) return { success: false, message: "No transfer offer selected." };
    const transferSheet = ss.getSheetByName('Transfer Offers');
    const transferData = transferSheet.getDataRange().getValues();
    let targetRowIdx = -1, offer = null;

    for (let i = 1; i < transferData.length; i++) {
        if (transferData[i][0] === offerId) { targetRowIdx = i; offer = transferData[i]; break; }
    }
    if (targetRowIdx === -1) return { success: false, message: "Offer not found." };
    if (offer[5] !== 'Open') return { success: false, message: "That offer was just accepted by another participant." };

    const type = offer[2], dateEpoch = Number(offer[3]), details = offer[4], giver = offer[1];

    // Verify giver still owns it before moving
    if (type === 'Weekend') {
        const wSheet = ss.getSheetByName('Weekend Coverage');
        const wData = wSheet.getDataRange().getValues();
        let day = details.indexOf('Saturday') !== -1 ? 'Saturday' : 'Sunday';
        let verifyTargetIdx = -1;
        for (let i = 1; i < wData.length; i++) {
            if (new Date(wData[i][0]).getTime() === dateEpoch && wData[i][1] === day) {
                if (wData[i][3] !== giver) return { success: false, message: "The original giver no longer holds this assignment." };
                verifyTargetIdx = i;
                break;
            }
        }
        if (verifyTargetIdx === -1) return { success: false, message: "Assignment no longer exists." };
    } else if (type === 'Holiday') {
        const hSheet = ss.getSheetByName('Holiday Coverage');
        const hData = hSheet.getDataRange().getValues();
        let verifyTargetIdx = -1;
        for (let i = 1; i < hData.length; i++) {
            if (new Date(hData[i][1]).getTime() === dateEpoch && hData[i][2] === details) {
                if (hData[i][3] !== giver) return { success: false, message: "The original giver no longer holds this assignment." };
                verifyTargetIdx = i;
                break;
            }
        }
        if (verifyTargetIdx === -1) return { success: false, message: "Assignment no longer exists." };
    }

    if (type === 'Weekend') {
        const wSheet = ss.getSheetByName('Weekend Coverage');
        const wData = wSheet.getDataRange().getValues();
        let day = details.indexOf('Saturday') !== -1 ? 'Saturday' : 'Sunday';
        let partnerEpoch = day === 'Saturday' ? dateEpoch + 86400000 : dateEpoch - 86400000;
        let wTargetIdx = -1;
        for (let i = 1; i < wData.length; i++) {
            let wd = new Date(wData[i][0]).getTime();
            if (wd === partnerEpoch && wData[i][3] === name) return { success: false, message: "You cannot hold both Saturday and Sunday First Call positions for the same weekend." };
            if (wd === dateEpoch && wData[i][1] === day) wTargetIdx = i;
        }
        if (wTargetIdx !== -1) wSheet.getRange(wTargetIdx + 1, 4).setValue(name);
    } else if (type === 'Holiday') {
        const hSheet = ss.getSheetByName('Holiday Coverage');
        const hData = hSheet.getDataRange().getValues();
        let hTargetIdx = -1;
        for (let i = 1; i < hData.length; i++) {
            let hd = new Date(hData[i][1]).getTime();
            if (hd === dateEpoch) {
                if (hData[i][2] !== details && hData[i][3] === name) return { success: false, message: "You cannot hold both call positions for the same holiday." };
                if (hData[i][2] === details) hTargetIdx = i;
            }
        }
        if (hTargetIdx !== -1) hSheet.getRange(hTargetIdx + 1, 4).setValue(name);
    }

    transferSheet.getRange(targetRowIdx + 1, 6).setValue('Accepted');
    transferSheet.getRange(targetRowIdx + 1, 7).setValue(name);
    ss.getSheetByName('Transfer History').appendRow([ new Date(), _getConfigValue('ActiveYear', new Date().getFullYear()), type, new Date(dateEpoch), details, offer[1], name ]);
    turnSheet.getRange(userRowIdx, turnStatusCol).setValue('Completed');
    return { success: true, message: "Transfer accepted successfully." };
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


// ============================================================================
// NEW TESTS
// ============================================================================

function runAllNewTests() {
    testStrictSerpentineBoundary();
    testTransferOfferLocking();
    testTwilioTokenNonExposure();
}

function testStrictSerpentineBoundary() {
    let ss;
    const originalGetActive = SpreadsheetApp.getActiveSpreadsheet;
    try {
        ss = setupMockSpreadsheet();
        SpreadsheetApp.getActiveSpreadsheet = () => ss;
        setupSpreadsheetSchema();

        const turnSheet = ss.getSheetByName('Turn Management');
        const configSheet = ss.getSheetByName('Config');
        _setConfigValue('CurrentPhase', 'VACATION_RANDOM');
        _setConfigValue('CurrentRound', 2);
        _setConfigValue('CurrentDirection', 'DESCENDING');

        // Mock Admin options for window size = 3
        const originalAdmin = getAdminOptions;
        getAdminOptions = () => ({ 'VACATION_WINDOW_SIZE': 3 });

        for (let i = 2; i <= 6; i++) {
            turnSheet.getRange(i, 6).setValue(i - 1); // Lottery: 1 to 5
        }

        turnSheet.getRange(6, 4).setValue('Completed');
        turnSheet.getRange(5, 4).setValue('Completed');
        turnSheet.getRange(4, 4).setValue('Completed');
        turnSheet.getRange(3, 4).setValue('Completed');

        let turnDataRaw = turnSheet.getDataRange().getValues();
        let q = calculateQueueWindow(turnDataRaw);
        let p1 = q.find(p => p.name === 'Person1');
        if (p1.computedStatus !== 'Active') throw new Error("Person1 should be active.");

        turnSheet.getRange(2, 4).setValue('Completed');
        let advanced = _advanceQueueDirectionIfComplete();
        if (!advanced) throw new Error("Boundary should have advanced.");
        if (_getConfigValue('CurrentDirection') !== 'ASCENDING') throw new Error("Direction didn't reverse");
        if (_getConfigValue('CurrentRound') !== 3) throw new Error("Round didn't increment");

        console.log("PASS: testStrictSerpentineBoundary");
        getAdminOptions = originalAdmin;
    } finally {
        SpreadsheetApp.getActiveSpreadsheet = originalGetActive;
        if (ss) {
            const files = DriveApp.getFilesByName(ss.getName());
            while (files.hasNext()) files.next().setTrashed(true);
        }
    }
}

function testTransferOfferLocking() {
    let ss;
    const originalGetActive = SpreadsheetApp.getActiveSpreadsheet;
    try {
        ss = setupMockSpreadsheet();
        SpreadsheetApp.getActiveSpreadsheet = () => ss;
        setupSpreadsheetSchema();

        const turnSheet = ss.getSheetByName('Turn Management');
        const partSheet = ss.getSheetByName('Participant Config');

        _setConfigValue('CurrentPhase', 'TRANSFER_OFFER_COLLECTION');
        _setConfigValue('TransferLocked', true);
        partSheet.getRange(2, partSheet.getDataRange().getValues()[0].indexOf('Transfer Giver') + 1).setValue(true);

        let res = _processTransferOffer({ name: 'Person1', offers: [{type: 'Weekend', dateEpoch: 123, details: 'test'}] }, ss, turnSheet, partSheet);
        if (res.success) throw new Error("Should not be able to offer when locked.");
        if (res.message.indexOf("locked") === -1) throw new Error("Wrong error message for lock.");

        console.log("PASS: testTransferOfferLocking");
    } finally {
        SpreadsheetApp.getActiveSpreadsheet = originalGetActive;
        if (ss) {
            const files = DriveApp.getFilesByName(ss.getName());
            while (files.hasNext()) files.next().setTrashed(true);
        }
    }
}

function testTwilioTokenNonExposure() {
    let ss;
    const originalGetActive = SpreadsheetApp.getActiveSpreadsheet;
    try {
        ss = setupMockSpreadsheet();
        SpreadsheetApp.getActiveSpreadsheet = () => ss;
        setupSpreadsheetSchema();

        const adminSheet = ss.getSheetByName('Admin Options');
        adminSheet.appendRow(['TWILIO_AUTH_TOKEN', 'SECRET123', 'desc']);

        let dbData = getDashboardData('Person1');
        let jsonStr = JSON.stringify(dbData);
        if (jsonStr.indexOf('SECRET123') !== -1) throw new Error("Dashboard payload leaked Twilio token!");

        console.log("PASS: testTwilioTokenNonExposure");
    } finally {
        SpreadsheetApp.getActiveSpreadsheet = originalGetActive;
        if (ss) {
            const files = DriveApp.getFilesByName(ss.getName());
            while (files.hasNext()) files.next().setTrashed(true);
        }
    }
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

function _getConfigValue(key, defaultValue) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  if (!configSheet) return defaultValue;
  const configData = configSheet.getDataRange().getValues();
  for (let i = 0; i < configData.length; i++) {
    if (configData[i][0] === key) {
      return configData[i][1];
    }
  }
  return defaultValue;
}

function _setConfigValue(key, value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName('Config');
  if (!configSheet) return false;
  const configData = configSheet.getDataRange().getValues();
  for (let i = 0; i < configData.length; i++) {
    if (configData[i][0] === key) {
      configSheet.getRange(i + 1, 2).setValue(value);
      return true;
    }
  }
  return false;
}

function getAdminOptions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const optionsSheet = ss.getSheetByName('Admin Options');
  if (!optionsSheet) return {};

  const data = optionsSheet.getDataRange().getValues();
  const options = {};
  for (let i = 1; i < data.length; i++) {
    const key = data[i][0];
    const value = data[i][1];
    if (key) options[key] = value;
  }
  return options;
}

function getAdminPhoneNumber() {
  const adminOptions = getAdminOptions();
  return adminOptions['ADMIN_PHONE'] || '';
}

// Override _smsDependencies
const _smsDependencies = {
  getProperties: () => {
    const adminOptions = getAdminOptions();
    return {
      SMS_NOTIFICATIONS_ENABLED: String(adminOptions['SMS_ENABLED'] || 'false').toLowerCase(),
      TWILIO_ACCOUNT_SID: adminOptions['TWILIO_ACCOUNT_SID'] || '',
      TWILIO_AUTH_TOKEN: adminOptions['TWILIO_AUTH_TOKEN'] || '',
      TWILIO_FROM_NUMBER: adminOptions['TWILIO_FROM_NUMBER'] || '',
      VACATION_SELECTOR_URL: ScriptApp.getService().getUrl()
    };
  },
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
  const vacationUrl = props['VACATION_SELECTOR_URL'] || 'https://script.google.com/...';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const logSheet = ss.getSheetByName('Notification Log');
  if (!logSheet) return;

  const logHeaders = logSheet.getRange(1, 1, 1, logSheet.getLastColumn()).getValues()[0];
  const nameIdx = logHeaders.indexOf('ParticipantName');
  const phaseIdx = logHeaders.indexOf('Round');
  const statusIdx = logHeaders.indexOf('Status');
  const sidIdx = logHeaders.indexOf('TwilioMessageSid');
  const errorIdx = logHeaders.indexOf('Error');
  const typeIdx = logHeaders.indexOf('Type');

  const partSheet = ss.getSheetByName('Participant Config');
  const partData = partSheet.getDataRange().getValues();
  const tNameIdx = partData[0].indexOf('Name');
  const tPhoneIdx = partData[0].indexOf('PhoneNumber');

  rowIndices.forEach(rowIdx => {
    const rowRange = logSheet.getRange(rowIdx, 1, 1, logHeaders.length);
    const rowValues = rowRange.getValues()[0];

    if (rowValues[statusIdx] !== 'PENDING') return;
    logSheet.getRange(rowIdx, statusIdx + 1).setValue('PROCESSING');

    const pName = rowValues[nameIdx];
    const pPhase = rowValues[phaseIdx];
    const type = typeIdx !== -1 ? rowValues[typeIdx] : 'INITIAL';

    let phoneNum = null;
    if (tNameIdx !== -1 && tPhoneIdx !== -1) {
      const pRow = partData.find(r => r[tNameIdx] === pName);
      if (pRow) phoneNum = String(pRow[tPhoneIdx] || '').trim();
    }

    if (type === 'ADMIN_ALERT') phoneNum = getAdminPhoneNumber();

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

    let msg = '';
    if (type === 'HOLIDAY_CONFIRM') msg = `Vacation Selector: You have successfully secured the nearby holiday position you requested.`;
    else if (type === 'ADMIN_ALERT') msg = `ADMIN ALERT: Participant ${pName} has been unresponsive in the Active window (${pPhase}) for the configured threshold limit.`;
    else if (type === 'REMINDER') msg = `Vacation Selector Reminder: You are STILL in the Active window for ${pPhase}. Please make your selection ASAP: ${vacationUrl}`;
    else msg = `Vacation Selector: It is your turn! You are in the Active window for ${pPhase}. Make your selection here: ${vacationUrl}`;

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
  const props = PropertiesService.getScriptProperties();
  const originalSms = props.getProperty('SMS_NOTIFICATIONS_ENABLED');
  try {
    props.setProperty('SMS_NOTIFICATIONS_ENABLED', 'false'); // Disable SMS
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
    if (originalSms !== null) {
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

function testAdminControlAlreadyRunning() {
  let ss;
  const originalGetActive = SpreadsheetApp.getActiveSpreadsheet;
  const props = PropertiesService.getScriptProperties();
  const originalSms = props.getProperty('SMS_NOTIFICATIONS_ENABLED');
  try {
    props.setProperty('SMS_NOTIFICATIONS_ENABLED', 'false'); // Disable SMS
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
    if (originalSms !== null) {
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

function testAdminControlSuccessfulStart() {
  let ss;
  const originalGetActive = SpreadsheetApp.getActiveSpreadsheet;
  const props = PropertiesService.getScriptProperties();
  const originalSms = props.getProperty('SMS_NOTIFICATIONS_ENABLED');
  try {
    props.setProperty('SMS_NOTIFICATIONS_ENABLED', 'false'); // Disable SMS
    ss = setupMockSpreadsheet();
    SpreadsheetApp.getActiveSpreadsheet = () => ss;
    setupSpreadsheetSchema(); // Need skipNextTurn etc

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
    if (originalSms !== null) {
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
  const props = PropertiesService.getScriptProperties();
  const originalSms = props.getProperty('SMS_NOTIFICATIONS_ENABLED');
  try {
    props.setProperty('SMS_NOTIFICATIONS_ENABLED', 'false'); // Disable SMS
    ss = setupMockSpreadsheet();
    SpreadsheetApp.getActiveSpreadsheet = () => ss;
    setupSpreadsheetSchema(); // Need skipNextTurn etc

    // Mess up lottery position
    const turnSheet = ss.getSheetByName('Turn Management');
    turnSheet.getRange(2, 6).setValue(''); // Blank LotteryPosition for Person1

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
    if (originalSms !== null) {
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

function testAdminControlRepeatedTaps() {
  let ss;
  const originalGetActive = SpreadsheetApp.getActiveSpreadsheet;
  const props = PropertiesService.getScriptProperties();
  const originalSms = props.getProperty('SMS_NOTIFICATIONS_ENABLED');
  try {
    props.setProperty('SMS_NOTIFICATIONS_ENABLED', 'false'); // Disable SMS
    ss = setupMockSpreadsheet();
    SpreadsheetApp.getActiveSpreadsheet = () => ss;
    setupSpreadsheetSchema(); // Need skipNextTurn etc

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
    if (originalSms !== null) {
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

    // Enable started state for transition tests
    configSheet.getRange('B3').setValue(true);

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
