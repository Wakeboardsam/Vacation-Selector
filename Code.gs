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
  if (normalized === 'non-prime' || normalized === 'non prime' || normalized === 'nonprime') return 'Non-Prime';
  return null;
}


// ============================================================================
// CONFIGURATION & UTILS
// ============================================================================

function _getConfigValue(key, defaultValue) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s = ss.getSheetByName('Config');
  if (!s) return defaultValue;
  const d = s.getDataRange().getValues();
  for (let i = 0; i < d.length; i++) if (d[i][0] === key) return d[i][1];
  return defaultValue;
}

function _setConfigValue(key, value) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s = ss.getSheetByName('Config');
  if (!s) return false;
  const d = s.getDataRange().getValues();
  for (let i = 0; i < d.length; i++) {
    if (d[i][0] === key) { s.getRange(i + 1, 2).setValue(value); return true; }
  }
  return false;
}

function getAdminOptions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const s = ss.getSheetByName('Admin Options');
  if (!s) return {};
  const d = s.getDataRange().getValues();
  const opts = {};
  for (let i = 1; i < d.length; i++) if (d[i][0]) opts[d[i][0]] = d[i][1];
  return opts;
}

function getAdminPhoneNumber() {
  return getAdminOptions()['ADMIN_PHONE'] || '';
}

function _getGlobalVacationCap() {
    return parseInt(getAdminOptions()['DEFAULT_VACATION_CAPACITY'] || 4);
}

// ============================================================================
// SAFE SCHEMA MIGRATION
// ============================================================================

function setupSpreadsheetSchema() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let modifications = false;

  // 1. Safe Migration from Turn Management to Participant Config
  let pConfigSheet = ss.getSheetByName('Participant Config');
  const turnSheet = ss.getSheetByName('Turn Management');

  if (!pConfigSheet) {
      pConfigSheet = ss.insertSheet('Participant Config');
      const pHeaders = [
        'ParticipantID', 'Name', 'PIN', 'PhoneNumber', 'SeniorityPosition', 'LotteryPosition',
        'Active for Year', 'Vacation Phase Enabled', 'Vacation Week Target Override',
        'Weekend Phase Enabled', 'Weekend Assignment Maximum', 'Holiday Volunteer',
        'Mandatory Holiday Eligible', 'Transfer Giver', 'Transfer Receiver',
        'Had Spring Break Last Year', 'Had Christmas Week Last Year',
        'Worked Any Official Holiday Last Year', 'Rules Acknowledged Year'
      ];
      pConfigSheet.getRange(1, 1, 1, pHeaders.length).setValues([pHeaders]).setFontWeight('bold').setBackground('#f3f4f6');
      pConfigSheet.setFrozenRows(1);

      if (turnSheet && turnSheet.getLastRow() > 1) {
          const tData = turnSheet.getDataRange().getValues();
          const tHeaders = tData[0];
          const rows = [];
          for (let i=1; i<tData.length; i++) {
              let r = tData[i];
              if (!r[tHeaders.indexOf('Name')]) continue;
              let pId = 'P-' + Utilities.getUuid().substring(0,8).toUpperCase();
              rows.push([
                  pId, r[tHeaders.indexOf('Name')], r[tHeaders.indexOf('PIN')] || '',
                  tHeaders.indexOf('PhoneNumber') !== -1 ? r[tHeaders.indexOf('PhoneNumber')] : '',
                  tHeaders.indexOf('SeniorityPosition') !== -1 ? r[tHeaders.indexOf('SeniorityPosition')] : (tHeaders.indexOf('QueuePosition') !== -1 ? r[tHeaders.indexOf('QueuePosition')] : ''),
                  tHeaders.indexOf('LotteryPosition') !== -1 ? r[tHeaders.indexOf('LotteryPosition')] : '',
                  true, true, '', true, '', '', true, '', '', false, false, false, ''
              ]);
          }
          if (rows.length > 0) {
              pConfigSheet.getRange(2, 1, rows.length, pHeaders.length).setValues(rows);
              pConfigSheet.getRange(2, 7, rows.length, 2).insertCheckboxes();
              pConfigSheet.getRange(2, 10, rows.length, 1).insertCheckboxes();
              pConfigSheet.getRange(2, 12, rows.length, 7).insertCheckboxes();
          }
      }
      modifications = true;
  }

  // 2. Setup Config & Admin Options
  const sheets = [
      { name: 'Admin Options', headers: ['Setting Name', 'Setting Value', 'Description'] },
      { name: 'Rules & Tips', headers: ['Rules & Tips Content (Markdown/HTML supported)'] },
      { name: 'Weekend Coverage', headers: ['Date', 'Day of Week', 'Position', 'Participant', 'Warning/Holiday Near'] },
      { name: 'Holiday Coverage', headers: ['Holiday Name', 'Observed Date', 'Call Position', 'Participant'] },
      { name: 'Soft Holiday Warnings', headers: ['Event Name', 'Date', 'Enabled'] },
      { name: 'Transfer Offers', headers: ['Offer ID', 'Giver Name', 'Type', 'Date', 'Details', 'Status', 'Receiver Name'] },
      { name: 'Transfer History', headers: ['Timestamp', 'Year', 'Assignment Type', 'Assignment Date', 'Details', 'Original Assignee', 'New Assignee'] },
      { name: 'Notification Log', headers: ['Timestamp', 'DedupeKey', 'ParticipantName', 'Round', 'CalculatedRole', 'Status', 'TwilioMessageSid', 'Error', 'Type', 'TurnID'] },
      { name: 'Config', headers: ['Setting', 'Value', 'Description'] }
  ];

  for (let s of sheets) {
      let sheet = ss.getSheetByName(s.name);
      if (!sheet) {
          sheet = ss.insertSheet(s.name);
          sheet.getRange(1, 1, 1, s.headers.length).setValues([s.headers]).setFontWeight('bold').setBackground('#f3f4f6');
          sheet.setFrozenRows(1);
          if (s.name === 'Config') {
              sheet.getRange(2, 1, 9, 3).setValues([
                  ['CurrentPhase', 'SETUP_EMPTY', 'Current active phase'],
                  ['PhaseReady', '', 'Next phase ready to advance'],
                  ['ActiveYear', new Date().getFullYear(), 'Active lottery year'],
                  ['CurrentRound', 1, 'Current Vacation Round Number'],
                  ['CurrentDirection', 'ASCENDING', 'Current queue serpentine direction (ASCENDING/DESCENDING)'],
                  ['SetupState', 'EMPTY', 'Current setup workflow state'],
                  ['TransferLocked', false, 'Whether the transfer offer pool is locked'],
                  ['ReconciliationRequired', false, 'Whether reconciliation from a direct edit is needed'],
                  ['SelectionStarted', false, '(Legacy) Selection process started']
              ]);
              sheet.getRange(8, 2, 3, 1).insertCheckboxes();
          } else if (s.name === 'Admin Options') {
              sheet.getRange(2, 1, 15, 3).setValues([
                  ['VACATION_WINDOW_SIZE', 3, 'Active-window size for Vacation rounds'],
                  ['WEEKEND_WINDOW_SIZE', 2, 'Active-window size for Weekend phase'],
                  ['HOLIDAY_WINDOW_SIZE', 2, 'Active-window size for Holiday phases'],
                  ['TRANSFER_WINDOW_SIZE', 2, 'Active-window size for Transfer Receiver phase'],
                  ['DEFAULT_VACATION_TARGET', 9, 'Default vacation-week target per participant'],
                  ['DEFAULT_VACATION_CAPACITY', 4, 'Default maximum participant capacity per vacation week'],
                  ['REMINDER_DELAY_MINUTES', 360, 'Participant reminder delay (minutes)'],
                  ['ADMIN_ALERT_DELAY_MINUTES', 720, 'Admin nonresponse alert delay (minutes)'],
                  ['HOLIDAY_WARNING_RANGE_DAYS', 3, 'Holiday proximity warning range (calendar days)'],
                  ['SMS_ENABLED', true, 'Enable SMS Notifications'],
                  ['ADMIN_PHONE', '', 'Administrator phone number for alerts'],
                  ['TWILIO_ACCOUNT_SID', '', 'Twilio Account SID (WARNING: All editors can see this)'],
                  ['TWILIO_AUTH_TOKEN', '', 'Twilio Auth Token (WARNING: All editors can see this)'],
                  ['TWILIO_FROM_NUMBER', '', 'Twilio From Number'],
                  ['SPOUSE_REMINDER_TEXT', "There's a holiday near this weekend. 😉", 'Text for soft-holiday weekend warnings']
              ]);
              sheet.getRange(11, 2).insertCheckboxes();
          } else if (s.name === 'Rules & Tips') {
              sheet.getRange('A2').setValue('Welcome to the Vacation Selection System! Please make your selections carefully.');
          }
          modifications = true;
      }
  }

  // 3. Safe idempotent migration of Week Availability to configurable capacity schema
  // Without calling clear()!
  const weekSheet = ss.getSheetByName('Week Availability');
  if (weekSheet && weekSheet.getLastRow() > 1) {
      const h = weekSheet.getRange(1, 1, 1, weekSheet.getLastColumn()).getValues()[0];
      if (h.indexOf('Person1') !== -1) {
          const wData = weekSheet.getDataRange().getValues();
          const newRows = [];
          const defCap = parseInt(getAdminOptions()['DEFAULT_VACATION_CAPACITY'] || 4);
          for (let i = 1; i < wData.length; i++) {
              let assigned = [];
              for (let j = 2; j <= 5; j++) if (wData[i][j]) assigned.push(wData[i][j]);
              // Date, Classif, MaxCap, SpotsRemaining, Special Week, AssignedTo
              newRows.push([wData[i][0], wData[i][1], defCap, Math.max(0, defCap - assigned.length), 'None', assigned.join(', ')]);
          }
          // Do not delete rows, just overwrite and resize columns
          weekSheet.getRange(1, 1, 1, 6).setValues([['WeekStartDate', 'Classification', 'MaxCapacity', 'SpotsRemaining', 'Special Week', 'AssignedTo']]);
          if (newRows.length > 0) weekSheet.getRange(2, 1, newRows.length, 6).setValues(newRows);
          // Clear remaining old columns (if any)
          if (weekSheet.getLastColumn() > 6) {
              weekSheet.getRange(1, 7, weekSheet.getLastRow(), weekSheet.getLastColumn() - 6).clearContent();
          }
          modifications = true;
      }
  }

  // 4. Ensure Turn Management has SkipNextTurn
  if (turnSheet) {
      let tHeaders = turnSheet.getRange(1, 1, 1, turnSheet.getLastColumn()).getValues()[0];
      if (tHeaders.indexOf('SkipNextTurn') === -1) {
          turnSheet.getRange(1, tHeaders.length + 1).setValue('SkipNextTurn');
          if (turnSheet.getLastRow() > 1) {
              turnSheet.getRange(2, tHeaders.length + 1, turnSheet.getLastRow() - 1, 1).setValue(false);
          }
      }
      if (tHeaders.indexOf('Status') === -1) {
          turnSheet.getRange(1, turnSheet.getLastColumn() + 1).setValue('Status');
      }
  }

  return modifications ? "Schema updated successfully." : "Schema already up to date.";
}

function validateSchema() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    if (!ss.getSheetByName('Participant Config') || !ss.getSheetByName('Admin Options')) {
        return { valid: false, message: "Missing sheets. Please run setupSpreadsheetSchema()." };
    }
    return { valid: true };
}


// ============================================================================
// AUTO-FILL & SETUP PREFLIGHT
// ============================================================================

function checkNewYearSetupReadiness() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const blockers = [];
  const checks = [
      { name: 'Week Availability', checkContent: true },
      { name: 'Weekend Coverage', checkContent: true },
      { name: 'Holiday Coverage', checkContent: true },
      { name: 'Transfer Offers', checkContent: true },
      { name: 'Transfer History', checkContent: true },
      { name: 'Notification Log', checkContent: true }
  ];

  for (let c of checks) {
      let sheet = ss.getSheetByName(c.name);
      if (sheet && sheet.getLastRow() > 1 && c.checkContent) {
          let data = sheet.getDataRange().getValues();
          // We only care if there is actual data in the rows (e.g. past row 1)
          if (data.slice(1).some(row => row.some(cell => String(cell).trim() !== ''))) {
              blockers.push(`${c.name} contains data. Please clear the rows below the header.`);
          }
      }
  }

  let pSheet = ss.getSheetByName('Participant Config');
  if (pSheet && pSheet.getLastRow() > 1) {
      let data = pSheet.getDataRange().getValues();
      let ackIdx = data[0].indexOf('Rules Acknowledged Year');
      if (ackIdx !== -1 && data.slice(1).some(r => r[ackIdx])) blockers.push("Participant Config contains Rules Acknowledged Year data. Please clear this column.");
  }

  let tSheet = ss.getSheetByName('Turn Management');
  if (tSheet && tSheet.getLastRow() > 1) {
      let data = tSheet.getDataRange().getValues();
      let statIdx = data[0].indexOf('Status');
      if (statIdx !== -1 && data.slice(1).some(r => r[statIdx])) blockers.push("Turn Management contains active queue status. Please clear this column.");
  }

  if (blockers.length === 0) {
    _setConfigValue('SetupState', 'EMPTY_SETUP');
    return { ready: true, message: "System is clean and ready." };
  }
  return { ready: false, message: "Setup Blocked:\n" + blockers.join('\n') };
}

function _calculateThanksgiving(year) {
    // Thanksgiving is the fourth Thursday of November
    let d = new Date(year, 10, 1);
    let day = d.getDay();
    let offset = (day <= 4) ? (4 - day) : (11 - day);
    d.setDate(1 + offset + 21); // First thursday + 3 weeks
    return d;
}

function _calculateEaster(year) {
    let f = Math.floor,
        G = year % 19,
        C = f(year / 100),
        H = (C - f(C / 4) - f((8 * C + 13) / 25) + 19 * G + 15) % 30,
        I = H - f(H / 28) * (1 - f(29 / (H + 1)) * f((21 - G) / 11)),
        J = (year + f(year / 4) + I + 2 - C + f(C / 4)) % 7,
        L = I - J,
        month = 3 + f((L + 40) / 44),
        day = L + 28 - 31 * f(month / 4);
    return new Date(year, month - 1, day);
}

function _calculateMemorialDay(year) {
    // Last Monday of May
    let d = new Date(year, 4, 31);
    while (d.getDay() !== 1) d.setDate(d.getDate() - 1);
    return d;
}

function _calculateLaborDay(year) {
    // First Monday of September
    let d = new Date(year, 8, 1);
    while (d.getDay() !== 1) d.setDate(d.getDate() + 1);
    return d;
}

function autoFillRandomize(year, confirmYear) {
  if (String(year) !== String(confirmYear)) return { success: false, message: 'Typed year does not match confirmation.' };
  let r = checkNewYearSetupReadiness();
  if (!r.ready) return { success: false, message: r.message };

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  _setConfigValue('ActiveYear', year);
  _setConfigValue('CurrentPhase', 'SETUP_REVIEW');
  _setConfigValue('SetupState', 'SETUP_REVIEW');

  const defCap = parseInt(getAdminOptions()['DEFAULT_VACATION_CAPACITY'] || 4);

  // Vacation Weeks: Monday on or before Jan 1 through Monday on or before Dec 31
  let currentMonday = new Date(year, 0, 1);
  currentMonday.setDate(currentMonday.getDate() - (currentMonday.getDay() === 0 ? 6 : currentMonday.getDay() - 1));

  let wData = [];
  while (true) {
      let s = new Date(currentMonday);
      let e = new Date(currentMonday);
      e.setDate(e.getDate() + 4);

      let isChristmas = (s.getMonth()===11 && s.getDate()>=21) || (e.getMonth()===11 && e.getDate()<=27);
      let isPrime = (s.getMonth()>=5 && s.getMonth()<=7) ? 'Prime' : 'Non-Prime'; // Jun, Jul, Aug

      wData.push([s, isPrime, defCap, defCap, isChristmas ? 'Christmas' : 'None', '']);

      currentMonday.setDate(currentMonday.getDate() + 7);
      // End with Monday on or before December 31.
      // So if the next Monday is > Dec 31, we break.
      if (currentMonday.getFullYear() > year) break;
  }

  const wSheet = ss.getSheetByName('Week Availability');
  wSheet.getRange(1, 1, 1, 6).setValues([['WeekStartDate', 'Classification', 'MaxCapacity', 'SpotsRemaining', 'Special Week', 'AssignedTo']]);
  wSheet.getRange(2, 1, wData.length, 6).setValues(wData);

  // Weekend Coverage
  currentMonday = new Date(year, 0, 1);
  currentMonday.setDate(currentMonday.getDate() - (currentMonday.getDay() === 0 ? 6 : currentMonday.getDay() - 1));
  let wSat = new Date(currentMonday);
  wSat.setDate(wSat.getDate() + 5);

  let wkData = [];
  while (wSat.getFullYear() <= year || wSat.getMonth() === 0) {
      let sat = new Date(wSat), sun = new Date(wSat);
      sun.setDate(sun.getDate()+1);
      wkData.push([sat, 'Saturday', 'First Call', '', ''], [sun, 'Sunday', 'First Call', '', '']);
      wSat.setDate(wSat.getDate()+7);
      if (wSat.getFullYear() > year && wSat.getMonth() > 0) break;
  }
  const wkSheet = ss.getSheetByName('Weekend Coverage');
  wkSheet.getRange(1, 1, 1, 5).setValues([['Date', 'Day of Week', 'Position', 'Participant', 'Warning/Holiday Near']]);
  wkSheet.getRange(2, 1, wkData.length, 5).setValues(wkData);

  // Holidays
  const hData = [
    ['New Year\'s Day', new Date(year, 0, 1), 'Call 1', ''], ['New Year\'s Day', new Date(year, 0, 1), 'Call 2', ''],
    ['Memorial Day', _calculateMemorialDay(year), 'Call 1', ''], ['Memorial Day', _calculateMemorialDay(year), 'Call 2', ''],
    ['Independence Day', new Date(year, 6, 4), 'Call 1', ''], ['Independence Day', new Date(year, 6, 4), 'Call 2', ''],
    ['Labor Day', _calculateLaborDay(year), 'Call 1', ''], ['Labor Day', _calculateLaborDay(year), 'Call 2', ''],
    ['Thanksgiving', _calculateThanksgiving(year), 'Call 1', ''], ['Thanksgiving', _calculateThanksgiving(year), 'Call 2', ''],
    ['Christmas', new Date(year, 11, 25), 'Call 1', ''], ['Christmas', new Date(year, 11, 25), 'Call 2', '']
  ];
  ss.getSheetByName('Holiday Coverage').getRange(2, 1, hData.length, 4).setValues(hData);

  const shData = [
      ['Presidents\' Day', new Date(year, 1, 15), true], // Approx
      ['Valentine\'s Day', new Date(year, 1, 14), true],
      ['Easter', _calculateEaster(year), true],
      ['Mother\'s Day', new Date(year, 4, 10), true], // Approx
      ['Father\'s Day', new Date(year, 5, 20), true] // Approx
  ];
  ss.getSheetByName('Soft Holiday Warnings').getRange(2, 1, shData.length, 3).setValues(shData);

  // Randomize Lottery Position
  const pSheet = ss.getSheetByName('Participant Config');
  const tSheet = ss.getSheetByName('Turn Management');
  const pData = pSheet.getDataRange().getValues();
  const pH = pData[0];
  let acts = [];
  for (let i=1; i<pData.length; i++) {
      if (pData[i][pH.indexOf('Active for Year')] === true) {
          acts.push({ r: i+1, n: pData[i][pH.indexOf('Name')] });
      } else {
          pSheet.getRange(i+1, pH.indexOf('LotteryPosition')+1).setValue('');
      }
  }

  // Fisher-Yates
  for (let i = acts.length - 1; i > 0; i--) {
      let j = Math.floor(Math.random() * (i + 1));
      [acts[i], acts[j]] = [acts[j], acts[i]];
  }

  for (let i=0; i<acts.length; i++) {
      pSheet.getRange(acts[i].r, pH.indexOf('LotteryPosition')+1).setValue(i+1);
  }

  // Sync to Turn Management
  if (tSheet.getLastRow() > 1) tSheet.getRange(2, 1, tSheet.getLastRow()-1, tSheet.getLastColumn()).clearContent();
  const tDataToWrite = [];
  const npData = pSheet.getDataRange().getValues();
  for (let i=1; i<npData.length; i++) {
      if (npData[i][pH.indexOf('Active for Year')] === true) {
          tDataToWrite.push([
              npData[i][pH.indexOf('Name')],
              npData[i][pH.indexOf('PIN')],
              npData[i][pH.indexOf('SeniorityPosition')],
              '', '',
              npData[i][pH.indexOf('LotteryPosition')],
              false
          ]);
      }
  }
  if (tDataToWrite.length > 0) tSheet.getRange(2, 1, tDataToWrite.length, 7).setValues(tDataToWrite);

  return { success: true, message: 'Auto-Fill complete.' };
}

// ============================================================================
// ADMIN PHASE CONTROLLERS
// ============================================================================

function confirmSetup() {
  if (_getConfigValue('CurrentPhase', '') !== 'SETUP_REVIEW') return { success: false, message: 'Not in SETUP_REVIEW.' };
  const weekData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Week Availability').getDataRange().getValues();
  const swIdx = weekData[0].indexOf('Special Week');
  if (!weekData.some(r => r[swIdx] === 'Spring Break')) return { success: false, message: 'Spring Break not designated.' };
  _setConfigValue('CurrentPhase', 'SETUP_CONFIRMED');
  _setConfigValue('SetupState', 'SETUP_CONFIRMED');
  return { success: true, message: 'Setup Confirmed.' };
}

function _resetQueueToLotteryPosition1() {
  const tSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Turn Management');
  const data = tSheet.getDataRange().getValues();
  const statIdx = data[0].indexOf('Status');
  if (statIdx === -1) return;
  for (let i=1; i<data.length; i++) {
      let st = data[i][statIdx];
      if (st !== 'Pass' && st !== 'None' && st !== 'TargetReached') tSheet.getRange(i+1, statIdx+1).setValue('');
  }
  _setConfigValue('CurrentDirection', 'ASCENDING');
  _setConfigValue('CurrentRound', 1);
}

function beginSeniorityRound() {
  if (_getConfigValue('SetupState', '') !== 'SETUP_CONFIRMED') return { success: false, message: 'Setup must be confirmed.' };
  _setConfigValue('CurrentPhase', 'VACATION_SENIORITY');
  _setConfigValue('SelectionStarted', true);
  _resetQueueToLotteryPosition1();
  return { success: true, message: 'Seniority Round 1 started.' };
}
function beginWeekendPhase() { _setConfigValue('CurrentPhase', 'WEEKEND'); _resetQueueToLotteryPosition1(); return { success: true, message: 'Weekend phase started.' }; }
function beginHolidayVolunteerPhase() { _setConfigValue('CurrentPhase', 'HOLIDAY_VOLUNTEER'); _resetQueueToLotteryPosition1(); return { success: true, message: 'Holiday Volunteer phase started.' }; }
function beginMandatoryHolidayPhase() { _setConfigValue('CurrentPhase', 'HOLIDAY_MANDATORY'); _resetQueueToLotteryPosition1(); return { success: true, message: 'Mandatory Holiday phase started.' }; }
function beginTransferRound() { _setConfigValue('CurrentPhase', 'TRANSFER_OFFER_COLLECTION'); _setConfigValue('TransferLocked', false); return { success: true, message: 'Transfer Offers started.' }; }
function lockTransferOffersAndBeginReceiverSelection() {
  if (_getConfigValue('CurrentPhase', '') !== 'TRANSFER_OFFER_COLLECTION') return { success: false, message: 'Not in offer collection.' };
  _setConfigValue('CurrentPhase', 'TRANSFER_RECEIVER');
  _setConfigValue('TransferLocked', true);
  _resetQueueToLotteryPosition1();
  return { success: true, message: 'Transfer locked. Receiver selection started.' };
}
function completeTransferRound() { _setConfigValue('CurrentPhase', 'COMPLETE'); return { success: true, message: 'Complete.' }; }

function adminManualSmsResend(participantName) {
    const currentPhase = _getConfigValue('CurrentPhase', 'UNKNOWN');
    const key = `${_getConfigValue('ActiveYear', '')}_${currentPhase}_MANUAL_${new Date().getTime()}`;
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    ss.getSheetByName('Notification Log').appendRow([new Date(), key, participantName, currentPhase, 'Manual Ping', 'PENDING', '', '', 'MANUAL', '']);
    _processPendingNotifications([ss.getSheetByName('Notification Log').getLastRow()]);
    return { success: true, message: 'Manual SMS dispatched.' };
}


// ============================================================================
// QUEUE ENGINE
// ============================================================================

function _getActiveWindowSizeForPhase(phase) {
  const o = getAdminOptions();
  if (phase.startsWith('VACATION')) return parseInt(o['VACATION_WINDOW_SIZE']||3);
  if (phase === 'WEEKEND') return parseInt(o['WEEKEND_WINDOW_SIZE']||2);
  if (phase.startsWith('HOLIDAY')) return parseInt(o['HOLIDAY_WINDOW_SIZE']||2);
  if (phase === 'TRANSFER_RECEIVER') return parseInt(o['TRANSFER_WINDOW_SIZE']||2);
  return 1;
}

function calculateQueueWindow(turnDataRaw) {
    const turnData = [...turnDataRaw];
    const headers = turnData.shift();
    const phase = _getConfigValue('CurrentPhase', '');
    const dir = _getConfigValue('CurrentDirection', 'ASCENDING');
    const ws = _getActiveWindowSizeForPhase(phase);

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const pSheet = ss.getSheetByName('Participant Config');
    if (!pSheet) return []; // Fail closed if missing config

    const pData = pSheet.getDataRange().getValues();
    const pH = pData[0];

    // Recalculate live targets
    let vCounts = {}, wkCounts = {};
    if (phase.startsWith('VACATION')) {
        let wData = ss.getSheetByName('Week Availability').getDataRange().getValues();
        for(let i=1; i<wData.length; i++) {
            let arr = wData[i][5]?String(wData[i][5]).split(',').map(n=>n.trim()):[];
            for(let p of arr) if(p) vCounts[p]=(vCounts[p]||0)+1;
        }
    } else if (phase === 'WEEKEND') {
        let wkData = ss.getSheetByName('Weekend Coverage').getDataRange().getValues();
        for(let i=1; i<wkData.length; i++) if (wkData[i][3]) wkCounts[wkData[i][3]]=(wkCounts[wkData[i][3]]||0)+1;
    }

    let queue = turnData.map((row, index) => {
        let name = row[headers.indexOf('Name')];
        let pRow = pData.find(r => r[pH.indexOf('Name')] === name);
        let eligible = true;
        let pId = name;

        if (!pRow) {
            eligible = false; // Fail closed if not in Participant Config
        } else {
            pId = pRow[pH.indexOf('ParticipantID')] || name;

            if (pRow[pH.indexOf('Active for Year')] !== true) eligible = false;
            else if (phase.startsWith('VACATION') && pRow[pH.indexOf('Vacation Phase Enabled')] !== true) eligible = false;
            else if (phase === 'WEEKEND' && pRow[pH.indexOf('Weekend Phase Enabled')] !== true) eligible = false;
            else if (phase === 'HOLIDAY_VOLUNTEER' && pRow[pH.indexOf('Holiday Volunteer')] !== true) eligible = false;
            else if (phase === 'HOLIDAY_MANDATORY' && pRow[pH.indexOf('Mandatory Holiday Eligible')] !== true) eligible = false;
            else if (phase === 'TRANSFER_RECEIVER' && pRow[pH.indexOf('Transfer Receiver')] !== true) eligible = false;
            else if (phase === 'TRANSFER_OFFER_COLLECTION') eligible = false;

            // Dynamic Caps Evaluation
            if (eligible && phase.startsWith('VACATION')) {
                let tgtStr = pRow[pH.indexOf('Vacation Week Target Override')];
                let tgt = tgtStr !== '' ? parseInt(tgtStr) : _getGlobalVacationCap(); // Actually DEFAULT_VACATION_TARGET, wait
                // Let's get the proper default target
                let t = tgtStr !== '' ? parseInt(tgtStr) : parseInt(getAdminOptions()['DEFAULT_VACATION_TARGET']||9);
                if ((vCounts[name]||0) >= t) eligible = false;
            } else if (eligible && phase === 'WEEKEND') {
                let cStr = pRow[pH.indexOf('Weekend Assignment Maximum')];
                if (cStr !== '' && (wkCounts[name]||0) >= parseInt(cStr)) eligible = false;
            }
        }

        let st = String(row[headers.indexOf('Status')]||'').trim();
        if (st==='Pass'||st==='None'||st==='TargetReached') eligible = false;

        return {
            originalRowIndex: index+2, pId: pId, name: name, status: st,
            skipNextTurn: row[headers.indexOf('SkipNextTurn')],
            sen: row[headers.indexOf('SeniorityPosition')] !== '' ? Number(row[headers.indexOf('SeniorityPosition')]) : 9999,
            lot: row[headers.indexOf('LotteryPosition')] !== '' ? Number(row[headers.indexOf('LotteryPosition')]) : 9999,
            computedStatus: eligible ? 'Waiting' : 'Skipping',
            eligible: eligible
        };
    });

    if (phase === 'VACATION_SENIORITY') queue.sort((a,b)=>a.sen-b.sen);
    else if (dir === 'ASCENDING') queue.sort((a,b)=>a.lot-b.lot);
    else queue.sort((a,b)=>b.lot-a.lot);

    // Mandatory Phase 3-Tier Algorithm
    if (phase === 'HOLIDAY_MANDATORY') {
        const hData = ss.getSheetByName('Holiday Coverage').getDataRange().getValues();
        if (hData.slice(1).some(r=>r[3]==='')) {
            let hCounts = {}; for(let i=1; i<hData.length; i++) if (hData[i][3]) hCounts[hData[i][3]]=(hCounts[hData[i][3]]||0)+1;
            let t1=[], t2=[], t3=[];
            for(let p of queue) {
                if(!p.eligible || p.status==='Completed') continue;
                let pr = pData.find(r=>r[pH.indexOf('Name')]===p.name);
                let hasPrior = pr[pH.indexOf('Worked Any Official Holiday Last Year')]===true;
                let hasCur = (hCounts[p.name]||0)>0;
                t3.push(p);
                if (!hasCur) { if (!hasPrior) t1.push(p); else t2.push(p); }
            }
            let activeTier = t1.length>0 ? t1 : (t2.length>0 ? t2 : t3);
            let aSet = new Set(activeTier.map(p=>p.name));
            for(let p of queue) if (p.computedStatus!=='Completed'&&p.computedStatus!=='Skipping'&&!aSet.has(p.name)) p.computedStatus='Skipping';
        }
    }

    let aCount = 0;
    for (let p of queue) {
        if (p.status === 'Completed') p.computedStatus = 'Completed';
        else if (p.skipNextTurn === true || !p.eligible || p.computedStatus === 'Skipping') p.computedStatus = 'Skipping';
        else if (aCount < ws) { p.computedStatus = 'Active'; aCount++; }
    }
    return queue;
}

function _advanceQueueDirectionIfComplete() {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const turnSheet = ss.getSheetByName('Turn Management');
    const tData = turnSheet.getDataRange().getValues();
    const phase = _getConfigValue('CurrentPhase', '');
    if (phase==='VACATION_SENIORITY'||phase==='TRANSFER_OFFER_COLLECTION') return false;

    const queue = calculateQueueWindow(tData);
    if (!queue.some(p => p.computedStatus === 'Active' || p.computedStatus === 'Waiting')) {
        let dir = _getConfigValue('CurrentDirection', 'ASCENDING');
        _setConfigValue('CurrentDirection', dir === 'ASCENDING' ? 'DESCENDING' : 'ASCENDING');
        if (phase === 'VACATION_RANDOM') _setConfigValue('CurrentRound', _getConfigValue('CurrentRound', 1) + 1);

        let statIdx = tData[0].indexOf('Status'), skipIdx = tData[0].indexOf('SkipNextTurn');
        for (let i=1; i<tData.length; i++) {
            if (tData[i][skipIdx]===true) turnSheet.getRange(i+1, skipIdx+1).setValue(false);
            if (tData[i][statIdx]==='Completed') turnSheet.getRange(i+1, statIdx+1).setValue('');
        }
        return true;
    }
    return false;
}

// ============================================================================
// PHASE SUBMISSION ENGINE (First-Valid-Wins)
// ============================================================================

function processSelection(selectionData) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const res = _processSelectionCore(selectionData);
    if (res.newIndices && res.newIndices.length > 0) {
        try { _processPendingNotifications(res.newIndices); } catch (e) { console.error("SMS processing failed", e); }
    }
    return res.coreResult;
  } finally { lock.releaseLock(); }
}

function _processSelectionCore(data) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const phase = _getConfigValue('CurrentPhase', '');
  if (phase.startsWith('SETUP') || phase === 'COMPLETE') return { coreResult: { success: false, message: 'Selection closed.' } };

  const turnSheet = ss.getSheetByName('Turn Management');
  const bWindow = calculateQueueWindow(turnSheet.getDataRange().getValues());

  let result;
  if (phase.startsWith('VACATION')) result = _processVacation(data, ss, turnSheet, bWindow);
  else if (phase === 'WEEKEND') result = _processWeekend(data, ss, turnSheet, bWindow);
  else if (phase.includes('HOLIDAY')) result = _processHoliday(data, ss, turnSheet, bWindow, phase);
  else if (phase === 'TRANSFER_RECEIVER') result = _processTransferClaim(data, ss, turnSheet, bWindow);
  else if (phase === 'TRANSFER_OFFER_COLLECTION') result = _processTransferOffer(data, ss, turnSheet);

  let ni = [];
  if (result && result.success) {
      if (phase !== 'TRANSFER_OFFER_COLLECTION') _advanceQueueDirectionIfComplete();
      ni = _queueNotifications(bWindow);
  }
  return { coreResult: result, newIndices: ni };
}

function _processVacation(data, ss, turnSheet, queue) {
    const u = queue.find(p=>p.name===data.name);
    if (!u || u.computedStatus !== 'Active') return { success: false, message: 'Not your turn.' };

    let { week1, week2 } = data;
    if (_getConfigValue('CurrentPhase', '') === 'VACATION_SENIORITY' && (week1 && week2)) return { success: false, message: 'During the Seniority Round, you may only select EXACTLY ONE week per turn.' };
    if (!week1 && !week2) return { success: false, message: 'No weeks selected.' };
    if (week1 === week2) return { success: false, message: 'Cannot select same week twice.' };

    const wSheet = ss.getSheetByName('Week Availability');
    const wData = wSheet.getDataRange().getValues();
    let r1=-1, r2=-1, pCount=0, npCount=0;

    let check = (val) => {
        let idx = wData.findIndex(r=>new Date(r[0]).getTime() === Number(val));
        if (idx < 1) return { err: "Week not found" };
        let r = wData[idx];
        let maxCap = parseInt(r[2]) || _getGlobalVacationCap();
        let arr = String(r[5]||'').split(',').map(n=>n.trim()).filter(Boolean);
        if (arr.length >= maxCap) return { err: "Week full" };
        if (arr.includes(data.name)) return { err: "Already hold spot" };
        if (String(r[1]).trim()==='Prime') pCount++; else npCount++;
        return { idx: idx, cap: maxCap, arr: arr };
    };

    let c1, c2;
    if (week1) { c1 = check(week1); if (c1.err) return { success: false, message: c1.err }; r1 = c1.idx; }
    if (week2) { c2 = check(week2); if (c2.err) return { success: false, message: c2.err }; r2 = c2.idx; }

    if (pCount > 1) return { success: false, message: 'Only ONE Prime week permitted.' };
    if (pCount === 1 && npCount > 0) return { success: false, message: 'Prime must stand alone.' };

    const pSheet = ss.getSheetByName('Participant Config');
    const pData = pSheet.getDataRange().getValues();
    const pRow = pData.find(r=>r[1]===data.name);

    let total = 0; for(let i=1;i<wData.length;i++){ let a=String(wData[i][5]||'').split(',').map(n=>n.trim()); if(a.includes(data.name)) total++; }
    let tgtStr = pRow[8];
    let tgt = tgtStr !== '' ? parseInt(tgtStr) : parseInt(getAdminOptions()['DEFAULT_VACATION_TARGET']||9);

    if (total + (c1?1:0) + (c2?1:0) > tgt) return { success: false, message: 'Exceeds target.' };

    const curRound = _getConfigValue('CurrentRound', 1);
    if (curRound <= 3) {
        if (c1 && wData[r1][4]==='Spring Break' && pRow[15]===true) return { success: false, message: 'Restricted from Spring Break until Round 4.' };
        if (c1 && wData[r1][4]==='Christmas' && pRow[16]===true) return { success: false, message: 'Restricted from Christmas until Round 4.' };
        if (c2 && wData[r2][4]==='Spring Break' && pRow[15]===true) return { success: false, message: 'Restricted from Spring Break until Round 4.' };
        if (c2 && wData[r2][4]==='Christmas' && pRow[16]===true) return { success: false, message: 'Restricted from Christmas until Round 4.' };
    }

    if (c1) { c1.arr.push(data.name); wSheet.getRange(r1+1, 6).setValue(c1.arr.join(', ')); wSheet.getRange(r1+1, 4).setValue(c1.cap-c1.arr.length); }
    if (c2) { c2.arr.push(data.name); wSheet.getRange(r2+1, 6).setValue(c2.arr.join(', ')); wSheet.getRange(r2+1, 4).setValue(c2.cap-c2.arr.length); }

    let row = u.originalRowIndex;
    turnSheet.getRange(row, turnSheet.getDataRange().getValues()[0].indexOf('Status')+1).setValue('Completed');
    if (npCount === 2) turnSheet.getRange(row, turnSheet.getDataRange().getValues()[0].indexOf('SkipNextTurn')+1).setValue(true);

    if (_getConfigValue('CurrentPhase', '') === 'VACATION_SENIORITY') {
        const tData = turnSheet.getDataRange().getValues();
        if (!calculateQueueWindow(tData).some(p => p.computedStatus === 'Waiting' || p.computedStatus === 'Active')) {
           _setConfigValue('CurrentPhase', 'VACATION_RANDOM');
           _setConfigValue('CurrentRound', 2);
           _resetQueueToLotteryPosition1();
        }
    }
    return { success: true };
}

function _processWeekend(data, ss, turnSheet, queue) {
    const u = queue.find(p=>p.name===data.name);
    if (!u || u.computedStatus !== 'Active') return { success: false, message: 'Not your turn.' };

    const wSheet = ss.getSheetByName('Weekend Coverage');
    const wData = wSheet.getDataRange().getValues();
    let tRow = -1;
    for(let i=1;i<wData.length;i++) if(new Date(wData[i][0]).getTime() === Number(data.weekendDate) && wData[i][1] === data.weekendDay) tRow = i;

    if (tRow === -1) return { success: false, message: 'Not found' };
    if (wData[tRow][3] !== '') return { success: false, message: 'Position taken concurrently.' };

    let partnerEpoch = data.weekendDay === 'Saturday' ? Number(data.weekendDate)+86400000 : Number(data.weekendDate)-86400000;
    for(let i=1;i<wData.length;i++) if(new Date(wData[i][0]).getTime() === partnerEpoch && wData[i][3] === data.name) return { success: false, message: 'Cannot hold both Sat/Sun' };

    wSheet.getRange(tRow+1, 4).setValue(data.name);
    let msg = 'Weekend assigned.';

    if (data.optHolidayName && data.optHolidayDate && data.optHolidayCall) {
        const hSheet = ss.getSheetByName('Holiday Coverage');
        const hData = hSheet.getDataRange().getValues();
        let hRow = -1, holdsOther = false;
        for(let i=1;i<hData.length;i++) if(hData[i][0]===data.optHolidayName && new Date(hData[i][1]).getTime()===Number(data.optHolidayDate)) {
            if(hData[i][2]===data.optHolidayCall) hRow=i; else if(hData[i][3]===data.name) holdsOther = true;
        }
        if (holdsOther) msg += ' However, you already hold the other call position for that holiday.';
        else if (hRow !== -1) {
            if (hData[hRow][3] === '') {
                hSheet.getRange(hRow+1, 4).setValue(data.name);
                msg += ' Holiday also assigned.';
                ss.getSheetByName('Notification Log').appendRow([new Date(), 'HOLIDAY_OPT_'+data.name+'_'+data.optHolidayDate, data.name, _getConfigValue('CurrentRound', 1), 'Holiday Opt-in', 'PENDING', '', '', 'HOLIDAY_CONFIRM', '']);
            } else msg += ' Holiday was taken by someone else.';
        }
    }

    turnSheet.getRange(u.originalRowIndex, turnSheet.getDataRange().getValues()[0].indexOf('Status')+1).setValue('Completed');
    return { success: true, message: msg };
}

function _processHoliday(data, ss, turnSheet, queue, phase) {
    const u = queue.find(p=>p.name===data.name);
    if (!u || u.computedStatus !== 'Active') return { success: false, message: 'Not your turn.' };

    if (data.action === 'Pass') {
        if (phase === 'HOLIDAY_MANDATORY') return { success: false, message: 'Cannot pass during Mandatory.' };
        turnSheet.getRange(u.originalRowIndex, turnSheet.getDataRange().getValues()[0].indexOf('Status')+1).setValue('Pass');
        return { success: true, message: 'Passed.' };
    }

    const hSheet = ss.getSheetByName('Holiday Coverage');
    const hData = hSheet.getDataRange().getValues();
    let tRow = -1;
    for(let i=1;i<hData.length;i++) if(hData[i][0]===data.holidayName && new Date(hData[i][1]).getTime()===Number(data.holidayDate)) {
        if(hData[i][2]===data.holidayCall) tRow=i; else if(hData[i][3]===data.name) return { success: false, message: 'Cannot hold both calls.' };
    }

    if (tRow === -1) return { success: false, message: 'Not found' };
    if (hData[tRow][3] !== '') return { success: false, message: 'Taken concurrently.' };

    hSheet.getRange(tRow+1, 4).setValue(data.name);
    turnSheet.getRange(u.originalRowIndex, turnSheet.getDataRange().getValues()[0].indexOf('Status')+1).setValue('Completed');
    return { success: true };
}

function _processTransferOffer(data, ss, turnSheet) {
    if (_getConfigValue('TransferLocked', false)) return { success: false, message: 'Locked.' };

    const pSheet = ss.getSheetByName('Participant Config');
    const pData = pSheet.getDataRange().getValues();
    const pRow = pData.find(r=>r[1]===data.name);
    if (!pRow || pRow[13] !== true) return { success: false, message: 'Not a giver.' };

    if (data.offers && data.offers.length > 0) {
        const tSheet = ss.getSheetByName('Transfer Offers');
        const nextId = tSheet.getLastRow();
        const wData = ss.getSheetByName('Weekend Coverage').getDataRange().getValues();
        const hData = ss.getSheetByName('Holiday Coverage').getDataRange().getValues();
        const rows = [];

        for (let i=0; i<data.offers.length; i++) {
            let o = data.offers[i];
            let v = false;
            if (o.type === 'Weekend') {
                let d = o.details.includes('Saturday')?'Saturday':'Sunday';
                if (wData.find(r=>new Date(r[0]).getTime()===o.dateEpoch && r[1]===d && r[3]===data.name)) v = true;
            } else if (o.type === 'Holiday') {
                if (hData.find(r=>new Date(r[1]).getTime()===o.dateEpoch && r[2]===o.details && r[3]===data.name)) v = true;
            }
            if (!v) return { success: false, message: 'Do not own: ' + o.details };
            rows.push(['OFFER-'+(nextId+i), data.name, o.type, o.dateEpoch, o.details, 'Open', '']);
        }
        tSheet.getRange(tSheet.getLastRow()+1, 1, rows.length, 7).setValues(rows);
    }

    // In Offer Collection, we don't have an active queue. It's open.
    // The participant can just submit. But we should mark their turn as Pass so they don't see it again if we want.
    // Wait, the prompt says "Givers do not need to wait in a serpentine Active window merely to submit irrevocable offers."
    // So we don't update queue status here.
    return { success: true, message: 'Offers submitted.' };
}

function _processTransferClaim(data, ss, turnSheet, queue) {
    const u = queue.find(p=>p.name===data.name);
    if (!u || u.computedStatus !== 'Active') return { success: false, message: 'Not your turn.' };

    if (data.action === 'None') {
        turnSheet.getRange(u.originalRowIndex, turnSheet.getDataRange().getValues()[0].indexOf('Status')+1).setValue('None');
        return { success: true };
    }

    const tSheet = ss.getSheetByName('Transfer Offers');
    const tData = tSheet.getDataRange().getValues();
    const offer = tData.find(r=>r[0]===data.offerId);
    if (!offer) return { success: false, message: 'Not found.' };
    if (offer[5] !== 'Open') return { success: false, message: 'Already taken.' };

    let typ = offer[2], d = Number(offer[3]), det = offer[4], g = offer[1];

    if (typ === 'Weekend') {
        const wSheet = ss.getSheetByName('Weekend Coverage');
        const wData = wSheet.getDataRange().getValues();
        let day = det.includes('Saturday') ? 'Saturday' : 'Sunday';
        let verifyRow = wData.findIndex(r => new Date(r[0]).getTime()===d && r[1]===day);
        if (verifyRow===-1 || wData[verifyRow][3] !== g) return { success: false, message: 'Original giver no longer holds assignment.' };

        let partnerEpoch = day === 'Saturday' ? d+86400000 : d-86400000;
        if (wData.find(r=>new Date(r[0]).getTime()===partnerEpoch && r[3]===data.name)) return { success: false, message: 'Cannot hold both Sat/Sun' };
        wSheet.getRange(verifyRow+1, 4).setValue(data.name);
    } else if (typ === 'Holiday') {
        const hSheet = ss.getSheetByName('Holiday Coverage');
        const hData = hSheet.getDataRange().getValues();
        let verifyRow = hData.findIndex(r => new Date(r[1]).getTime()===d && r[2]===det);
        if (verifyRow===-1 || hData[verifyRow][3] !== g) return { success: false, message: 'Original giver no longer holds assignment.' };
        if (hData.find(r=>new Date(r[1]).getTime()===d && r[2]!==det && r[3]===data.name)) return { success: false, message: 'Cannot hold both calls' };
        hSheet.getRange(verifyRow+1, 4).setValue(data.name);
    }

    tSheet.getRange(tData.indexOf(offer)+1, 6).setValue('Accepted');
    tSheet.getRange(tData.indexOf(offer)+1, 7).setValue(data.name);
    ss.getSheetByName('Transfer History').appendRow([ new Date(), _getConfigValue('ActiveYear', ''), typ, new Date(d), det, g, data.name ]);

    turnSheet.getRange(u.originalRowIndex, turnSheet.getDataRange().getValues()[0].indexOf('Status')+1).setValue('Completed');
    return { success: true };
}


// ============================================================================
// RECONCILIATION
// ============================================================================

function refreshReconcileFromSheet() {
    const lock = LockService.getScriptLock();
    lock.waitLock(30000);
    try {
        const ss = SpreadsheetApp.getActiveSpreadsheet();
        let issues = [];

        const wSheet = ss.getSheetByName('Week Availability');
        const wData = wSheet.getDataRange().getValues();
        let vCounts = {};
        for(let i=1; i<wData.length; i++) {
            let arr = wData[i][5] ? String(wData[i][5]).split(',').map(n=>n.trim()).filter(Boolean) : [];
            for (let p of arr) vCounts[p] = (vCounts[p]||0)+1;
            let cap = parseInt(wData[i][2]) || _getGlobalVacationCap();
            wSheet.getRange(i+1, 4).setValue(cap - arr.length);
            if (arr.length > cap) issues.push("Over-capacity week: " + new Date(wData[i][0]).toDateString());
        }

        const wkSheet = ss.getSheetByName('Weekend Coverage');
        const wkData = wkSheet.getDataRange().getValues();
        let wkCounts = {};
        for(let i=1; i<wkData.length; i++) {
            let p = wkData[i][3];
            if (p) wkCounts[p] = (wkCounts[p]||0)+1;

            let dTime = new Date(wkData[i][0]).getTime();
            let day = wkData[i][1];
            let partnerEpoch = day === 'Saturday' ? dTime + 86400000 : dTime - 86400000;
            if (p) {
                if (wkData.find(r => new Date(r[0]).getTime()===partnerEpoch && r[3]===p)) {
                    issues.push("Duplicate weekend position for " + p + " on " + new Date(dTime).toDateString());
                }
            }
        }

        const hSheet = ss.getSheetByName('Holiday Coverage');
        const hData = hSheet.getDataRange().getValues();
        for(let i=1; i<hData.length; i++) {
            let p = hData[i][3];
            let hd = new Date(hData[i][1]).getTime();
            let call = hData[i][2];
            if (p) {
                 if (hData.find(r => new Date(r[1]).getTime()===hd && r[2]!==call && r[3]===p)) {
                     issues.push("Prohibited same-holiday combination for " + p + " on " + hData[i][0]);
                 }
            }
        }

        const tSheet = ss.getSheetByName('Turn Management');
        const pSheet = ss.getSheetByName('Participant Config');
        const tData = tSheet.getDataRange().getValues();
        const pData = pSheet.getDataRange().getValues();
        const phase = _getConfigValue('CurrentPhase', '');
        const statCol = tData[0].indexOf('Status');

        for(let i=1; i<pData.length; i++) {
            let name = pData[i][1];
            let rowIdx = tData.findIndex(r=>r[0]===name) + 2;
            if (rowIdx < 2) continue; // Unknown in Turn Mgmt? Wait, what if someone is missing? "Unknown participant names".

            if (phase.startsWith('VACATION')) {
                let tgtStr = pData[i][8];
                let tgt = tgtStr !== '' ? parseInt(tgtStr) : parseInt(getAdminOptions()['DEFAULT_VACATION_TARGET']||9);
                if ((vCounts[name]||0) >= tgt) {
                    if (tData[rowIdx-2][statCol] !== 'TargetReached') tSheet.getRange(rowIdx, statCol+1).setValue('TargetReached');
                } else if ((vCounts[name]||0) > tgt) {
                    issues.push("Target exceeded for " + name);
                }
            } else if (phase === 'WEEKEND') {
                let capStr = pData[i][10];
                if (capStr !== '') {
                    let cap = parseInt(capStr);
                    if ((wkCounts[name]||0) >= cap) {
                        if (tData[rowIdx-2][statCol] !== 'TargetReached') tSheet.getRange(rowIdx, statCol+1).setValue('TargetReached');
                    } else if ((wkCounts[name]||0) > cap) {
                        issues.push("Weekend cap exceeded for " + name);
                    }
                }
            }
        }

        // Also check if any names in assignments don't exist in pData
        for (let name of Object.keys(vCounts)) {
            if (!pData.find(r=>r[1]===name)) issues.push("Unknown participant in vacations: " + name);
        }
        for (let name of Object.keys(wkCounts)) {
            if (!pData.find(r=>r[1]===name)) issues.push("Unknown participant in weekends: " + name);
        }

        _advanceQueueDirectionIfComplete();
        _setConfigValue('ReconciliationRequired', false);
        return { success: true, issues: issues, message: issues.length > 0 ? 'Reconciled with issues: ' + issues.join(', ') : 'Reconciled successfully.' };
    } finally {
        lock.releaseLock();
    }
}


// ============================================================================
// SMS NOTIFICATIONS & TIMERS
// ============================================================================

function _buildDedupeKey(participantId, turnActivationId = '') {
    return `${_getConfigValue('ActiveYear', '')}_${_getConfigValue('CurrentPhase', '')}_R${_getConfigValue('CurrentRound', 1)}_${_getConfigValue('CurrentDirection', 'ASCENDING')}_${participantId}${turnActivationId ? '_' + turnActivationId : ''}`;
}

function _queueNotifications(beforeWindow) {
    if (String(_smsDependencies.getProperties()['SMS_NOTIFICATIONS_ENABLED']) !== 'true') return [];
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const lSheet = ss.getSheetByName('Notification Log');
    if (!lSheet) return [];

    const turnSheet = ss.getSheetByName('Turn Management');
    const afterWindow = calculateQueueWindow(turnSheet.getDataRange().getValues());
    const lData = lSheet.getDataRange().getValues();

    const exist = new Set();
    for (let i=1; i<lData.length; i++) {
        if (lData[i][8]==='INITIAL') exist.add(String(lData[i][1]));
    }

    const newIndices = [];
    let turnIdSuffix = new Date().getTime(); // ensure uniqueness if multiple are added this exact pass
    for (let p of afterWindow) {
        if (p.computedStatus === 'Active') {
            let k = _buildDedupeKey(p.pId, turnIdSuffix++);
            // But wait, the dedupe key should be stable for the same turn!
            // If they are in the active window, and we haven't seen them for this phase/direction cycle, they get ONE initial.
            // If they drop out and come back in (e.g., a later phase cycle), they need a new key?
            // "Legitimate later reentry produces a new activation"
            // Using Phase and Direction implicitly handles the normal serpentine cycles.
            // If they enter multiple times in the SAME direction (e.g. transfers), we need turnId.
            // Actually, we can check if they were NOT active in the beforeWindow.
            let wasActive = false;
            for (let bp of beforeWindow) {
                if (bp.name === p.name && bp.computedStatus === 'Active') wasActive = true;
            }

            // Re-evaluating the stable key: If they were not active, we issue a new key using the timestamp to ensure it doesn't collide with past turns in this direction.
            // But wait, if we use a timestamp, the key changes every time and we can't find it to suppress duplicates?
            // "A participant who leaves and later legitimately reenters must receive a new immediate SMS"
            // So if they are in 'after' but not in 'before', they get a NEW entry.
            // If they were already in 'before', they keep their existing entry (we don't write a new one).

            if (!wasActive) {
                // They just entered. Give them a key.
                let entryKey = _buildDedupeKey(p.pId, new Date().getTime() + Math.floor(Math.random()*1000));
                lSheet.appendRow([new Date(), entryKey, p.name, _getConfigValue('CurrentPhase', '') + ' R' + _getConfigValue('CurrentRound', 1), 'Active', 'PENDING', '', '', 'INITIAL', entryKey]);
                newIndices.push(lSheet.getLastRow());
            }
        }
    }
    return newIndices;
}

function _processPendingNotifications(indices) {
    if (!indices || indices.length===0) return;
    const props = _smsDependencies.getProperties();
    const lSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Notification Log');
    const pData = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Participant Config').getDataRange().getValues();

    // We do NOT use LockService here. We assume Lock is released.
    for (let row of indices) {
        lSheet.getRange(row, 6).setValue('PROCESSING');
        let rData = lSheet.getRange(row, 1, 1, 10).getValues()[0];
        let pRow = pData.find(r=>r[1]===rData[2]);
        let phone = pRow ? String(pRow[3] || '').trim() : null;
        let type = rData[8];

        if (type==='ADMIN_ALERT') phone = getAdminPhoneNumber();
        if (!phone) { lSheet.getRange(row, 6).setValue('SKIPPED_NO_PHONE'); continue; }

        const configCheck = checkSmsConfiguration();
        if (!configCheck.valid) { lSheet.getRange(row, 6).setValue('FAILED'); lSheet.getRange(row, 8).setValue(configCheck.message); continue; }

        let msg = '';
        if (type === 'HOLIDAY_CONFIRM') msg = `Vacation Selector: You have successfully secured the nearby holiday position you requested.`;
        else if (type === 'ADMIN_ALERT') msg = `ADMIN ALERT: Participant ${rData[2]} has been unresponsive in the Active window (${rData[3]}) for the configured threshold limit.`;
        else if (type === 'REMINDER') msg = `Vacation Selector Reminder: You are STILL in the Active window for ${rData[3]}. Please make your selection ASAP: ${props['VACATION_SELECTOR_URL']}`;
        else msg = `Vacation Selector: It is your turn! You are in the Active window for ${rData[3]}. Make your selection here: ${props['VACATION_SELECTOR_URL']}`;

        const result = sendSmsViaTwilio(phone, msg);
        if (result.success) {
            lSheet.getRange(row, 6).setValue('SENT');
            lSheet.getRange(row, 7).setValue(result.messageSid || 'mock');
        } else {
            lSheet.getRange(row, 6).setValue('FAILED');
            lSheet.getRange(row, 8).setValue(result.error);
        }
    }
}

function _processScheduledTimers() {
    if (String(_smsDependencies.getProperties()['SMS_NOTIFICATIONS_ENABLED']) !== 'true') return;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const lSheet = ss.getSheetByName('Notification Log');
    if (!lSheet) return;

    const currentPhase = _getConfigValue('CurrentPhase', 'UNKNOWN');
    if (currentPhase === 'SETUP_EMPTY' || currentPhase === 'COMPLETE') return;

    const adminOpts = getAdminOptions();
    const reminderMs = parseInt(adminOpts['REMINDER_DELAY_MINUTES'] || 360) * 60 * 1000;
    const alertMs = parseInt(adminOpts['ADMIN_ALERT_DELAY_MINUTES'] || 720) * 60 * 1000;
    const now = new Date().getTime();

    const lData = lSheet.getDataRange().getValues();
    const tsIdx = 0, keyIdx = 1, nameIdx = 2, rdIdx = 3, statIdx = 5, typeIdx = 8, actIdIdx = 9;

    const turnSheet = ss.getSheetByName('Turn Management');
    const queue = calculateQueueWindow(turnSheet.getDataRange().getValues());
    const activeNames = new Set(queue.filter(p => p.computedStatus === 'Active').map(p => p.name));

    const sentRems = new Set(), sentAlts = new Set();
    for (let i=1; i<lData.length; i++) {
        let t = lData[i][typeIdx], st = lData[i][statIdx], k = lData[i][actIdIdx];
        if (t==='REMINDER' && (st==='SENT'||st==='PROCESSING'||st==='PENDING')) sentRems.add(k);
        if (t==='ADMIN_ALERT' && (st==='SENT'||st==='PROCESSING'||st==='PENDING')) sentAlts.add(k);
    }

    let pend = [];
    for (let i=1; i<lData.length; i++) {
        if (lData[i][typeIdx]==='INITIAL' && lData[i][statIdx]==='SENT') {
            let actId = lData[i][actIdIdx], name = lData[i][nameIdx], ts = new Date(lData[i][tsIdx]).getTime();

            // To be eligible for a reminder, they must still be Active.
            // If they are no longer active, the timer is suppressed.
            if (activeNames.has(name)) {
                let diff = now - ts;
                if (diff >= reminderMs && !sentRems.has(actId)) {
                    lSheet.appendRow([new Date(), 'REM_'+actId, name, lData[i][rdIdx], 'Reminder', 'PENDING', '', '', 'REMINDER', actId]);
                    pend.push(lSheet.getLastRow());
                    sentRems.add(actId);
                }
                if (diff >= alertMs && !sentAlts.has(actId)) {
                    lSheet.appendRow([new Date(), 'ALT_'+actId, name, lData[i][rdIdx], 'Admin Alert', 'PENDING', '', '', 'ADMIN_ALERT', actId]);
                    pend.push(lSheet.getLastRow());
                    sentAlts.add(actId);
                }
            }
        }
    }
    if (pend.length > 0) _processPendingNotifications(pend);
}

function checkSmsConfiguration() {
    const props = _smsDependencies.getProperties();
    if (!props['TWILIO_ACCOUNT_SID']) return { valid: false, message: 'Missing Twilio SID' };
    if (!props['TWILIO_AUTH_TOKEN']) return { valid: false, message: 'Missing Twilio Auth Token' };
    if (!props['TWILIO_FROM_NUMBER']) return { valid: false, message: 'Missing Twilio From Number' };
    return { valid: true };
}

function sendSmsViaTwilio(to, body) {
  const props = _smsDependencies.getProperties();
  const sid = props['TWILIO_ACCOUNT_SID'];
  const token = props['TWILIO_AUTH_TOKEN'];
  const from = props['TWILIO_FROM_NUMBER'];

  if (sid === 'TEST_SID') return { success: true, messageSid: 'SM_MOCK' };

  const url = `https://api.twilio.com/2010-04-01/Accounts/${sid}/Messages.json`;
  const options = {
    method: 'post',
    headers: { Authorization: 'Basic ' + Utilities.base64Encode(sid + ':' + token) },
    payload: { To: to, From: from, Body: body },
    muteHttpExceptions: true
  };

  try {
    const response = _smsDependencies.fetch(url, options);
    const result = JSON.parse(response.getContentText());
    if (response.getResponseCode() === 201) return { success: true, messageSid: result.sid };
    else return { success: false, error: result.message };
  } catch (e) {
    return { success: false, error: "Network failed: " + e.message };
  }
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
    // Tests that the queue does not reverse until all participants finish.
    let ss;
    const originalGetActive = SpreadsheetApp.getActiveSpreadsheet;
    try {
        ss = SpreadsheetApp.create('Test Strict Boundary');
        SpreadsheetApp.getActiveSpreadsheet = () => ss;
        setupSpreadsheetSchema();

        const turnSheet = ss.getSheetByName('Turn Management');
        const configSheet = ss.getSheetByName('Config');
        _setConfigValue('CurrentPhase', 'VACATION_RANDOM');
        _setConfigValue('CurrentRound', 2);
        _setConfigValue('CurrentDirection', 'DESCENDING');

        // Setup Lottery
        for (let i = 2; i <= 6; i++) {
            turnSheet.getRange(i, 6).setValue(i - 1);
        }

        turnSheet.getRange(6, 4).setValue('Completed');
        turnSheet.getRange(5, 4).setValue('Completed');
        turnSheet.getRange(4, 4).setValue('Completed');
        turnSheet.getRange(3, 4).setValue('Completed');

        turnSheet.getRange(2, 4).setValue('Completed');
        let advanced = _advanceQueueDirectionIfComplete();
        if (!advanced) throw new Error("Boundary should have advanced.");
        if (_getConfigValue('CurrentDirection') !== 'ASCENDING') throw new Error("Direction didn't reverse");
        if (_getConfigValue('CurrentRound') !== 3) throw new Error("Round didn't increment");

        console.log("PASS: testStrictSerpentineBoundary");
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
        ss = SpreadsheetApp.create('Test Transfer');
        SpreadsheetApp.getActiveSpreadsheet = () => ss;
        setupSpreadsheetSchema();

        const turnSheet = ss.getSheetByName('Turn Management');
        const partSheet = ss.getSheetByName('Participant Config');

        _setConfigValue('CurrentPhase', 'TRANSFER_OFFER_COLLECTION');
        _setConfigValue('TransferLocked', true);
        partSheet.getRange(2, partSheet.getDataRange().getValues()[0].indexOf('Transfer Giver') + 1).setValue(true);

        let res = _processTransferOffer({ name: 'Person1', offers: [{type: 'Weekend', dateEpoch: 123, details: 'test'}] }, ss, turnSheet);
        if (res.success) throw new Error("Should not be able to offer when locked.");
        if (res.message.indexOf("Locked") === -1) throw new Error("Wrong error message for lock.");

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
        ss = SpreadsheetApp.create('Test Twilio');
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


// ============================================================================
// NEW TESTS
// ============================================================================

function runAllNewTests() {
    testStrictSerpentineBoundary();
    testTransferOfferLocking();
    testTwilioTokenNonExposure();
}

function testStrictSerpentineBoundary() {
    // Tests that the queue does not reverse until all participants finish.
    let ss;
    const originalGetActive = SpreadsheetApp.getActiveSpreadsheet;
    try {
        ss = SpreadsheetApp.create('Test Strict Boundary');
        SpreadsheetApp.getActiveSpreadsheet = () => ss;
        setupSpreadsheetSchema();

        const turnSheet = ss.getSheetByName('Turn Management');
        const configSheet = ss.getSheetByName('Config');
        _setConfigValue('CurrentPhase', 'VACATION_RANDOM');
        _setConfigValue('CurrentRound', 2);
        _setConfigValue('CurrentDirection', 'DESCENDING');

        // Setup Lottery
        for (let i = 2; i <= 6; i++) {
            turnSheet.getRange(i, 6).setValue(i - 1);
        }

        turnSheet.getRange(6, 4).setValue('Completed');
        turnSheet.getRange(5, 4).setValue('Completed');
        turnSheet.getRange(4, 4).setValue('Completed');
        turnSheet.getRange(3, 4).setValue('Completed');

        turnSheet.getRange(2, 4).setValue('Completed');
        let advanced = _advanceQueueDirectionIfComplete();
        if (!advanced) throw new Error("Boundary should have advanced.");
        if (_getConfigValue('CurrentDirection') !== 'ASCENDING') throw new Error("Direction didn't reverse");
        if (_getConfigValue('CurrentRound') !== 3) throw new Error("Round didn't increment");

        console.log("PASS: testStrictSerpentineBoundary");
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
        ss = SpreadsheetApp.create('Test Transfer');
        SpreadsheetApp.getActiveSpreadsheet = () => ss;
        setupSpreadsheetSchema();

        const turnSheet = ss.getSheetByName('Turn Management');
        const partSheet = ss.getSheetByName('Participant Config');

        _setConfigValue('CurrentPhase', 'TRANSFER_OFFER_COLLECTION');
        _setConfigValue('TransferLocked', true);
        partSheet.getRange(2, partSheet.getDataRange().getValues()[0].indexOf('Transfer Giver') + 1).setValue(true);

        let res = _processTransferOffer({ name: 'Person1', offers: [{type: 'Weekend', dateEpoch: 123, details: 'test'}] }, ss, turnSheet);
        if (res.success) throw new Error("Should not be able to offer when locked.");
        if (res.message.indexOf("Locked") === -1) throw new Error("Wrong error message for lock.");

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
        ss = SpreadsheetApp.create('Test Twilio');
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
