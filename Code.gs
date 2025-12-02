/**
 * Unified Calendar for Multi-Department Event Sheets
 *
 * This Google Apps Script creates a calendar view that aggregates events
 * from multiple department sheets into a single unified calendar.
 *
 * Usage:
 * 1. Open your Google Sheet with department event sheets
 * 2. Go to Extensions > Apps Script
 * 3. Paste this code and save
 * 4. Run installCalendar() to set up the calendar
 * 5. Use Calendar menu to refresh manually, or wait for auto-refresh
 */

// Configuration
const CONFIG = {
  calendarSheetName: 'Calendrier Unifié',
  // Supports both old (Date) and new (Début/Fin) column formats
  requiredColumns: ['Date', 'Service', 'Évènement'], // Legacy check
  startColumn: 'Début',  // New: start datetime column
  endColumn: 'Fin',      // New: end datetime column (optional)
  dateColumn: 'Date',    // Legacy: fallback if Début not found
  maxEventsPerDay: 8,
  refreshIntervalMinutes: 5,
  defaultEventDurationHours: 1, // Default duration for timed events without end
  locale: 'fr-FR',

  // Academic year runs August to July
  academicYearStartMonth: 7, // 0-indexed, so 7 = August

  // French month names (August to July order for academic year)
  monthNames: [
    'août', 'septembre', 'octobre', 'novembre', 'décembre',
    'janvier', 'février', 'mars', 'avril', 'mai', 'juin', 'juillet'
  ],

  // French weekday abbreviations (Monday first)
  weekdayNames: ['Lun', 'Mar', 'Mer', 'Jeu', 'Ven', 'Sam', 'Dim'],

  // Colored circle emojis for departments (19 distinct colors, assigned alphabetically)
  departmentEmojis: [
    '🔴', '🟡', '🟢', '🔵', '🟣',
    '🟠', '🩵', '💚', '💛', '🌕',
    '🧡', '💜', '💙', '🩷', '🩶',
    '💚', '🟠', '❤️', '🩷'
  ],

  // Google Calendar Sync settings
  googleCalendarId: 'c_45b0a97a534d195cca5fde4c9bc29f0d73dd80413b7937243f29166680031f15@group.calendar.google.com',
  syncTrackingSheetName: '_CalendarSync',
  syncDebounceSeconds: 30,

  // Error tracking
  errorSheetName: 'Erreurs',

  // School holidays (Hong Kong FIS/LFI calendar)
  schoolHolidays: {
    '2024-2025': [
      {name: 'Vacances été', start: '2024-06-28', end: '2024-08-27'},
      {name: "Vacances d'octobre", start: '2024-10-25', end: '2024-11-01'},
      {name: "Vacances d'hiver", start: '2024-12-23', end: '2025-01-05'},
      {name: 'Vacances Nouvel An chinois', start: '2025-01-29', end: '2025-02-02'},
      {name: 'Vacances de Pâques', start: '2025-04-14', end: '2025-04-27'},
      {name: 'Vacances de printemps', start: '2025-05-26', end: '2025-05-29'},
      {name: 'Vacances été', start: '2025-06-27', end: '2025-08-27'}
    ],
    '2025-2026': [
      {name: 'Vacances été', start: '2025-06-27', end: '2025-08-27'},
      {name: "Vacances d'octobre", start: '2025-10-24', end: '2025-10-31'},
      {name: "Vacances d'hiver", start: '2025-12-22', end: '2026-01-02'},
      {name: 'Vacances Nouvel An chinois', start: '2026-02-16', end: '2026-02-20'},
      {name: 'Vacances de Pâques', start: '2026-03-30', end: '2026-04-10'},
      {name: 'Vacances de printemps', start: '2026-05-26', end: '2026-05-29'},
      {name: 'Vacances été', start: '2026-06-27', end: '2026-08-27'}
    ]
  },

  // Hong Kong public holidays
  publicHolidays: {
    '2024-2025': [
      {date: '2024-09-18', name: 'Après Mi-automne'},
      {date: '2024-10-01', name: 'Fête nationale'},
      {date: '2024-10-11', name: 'Chung Yeung'},
      {date: '2024-12-25', name: 'Noël'},
      {date: '2024-12-26', name: 'Boxing Day'},
      {date: '2025-01-01', name: 'Jour de l\'an'},
      {date: '2025-01-29', name: 'Nouvel An chinois'},
      {date: '2025-01-30', name: '2e jour Nouvel An chinois'},
      {date: '2025-01-31', name: '3e jour Nouvel An chinois'},
      {date: '2025-04-04', name: 'Ching Ming'},
      {date: '2025-04-18', name: 'Vendredi saint'},
      {date: '2025-04-19', name: 'Après Vendredi saint'},
      {date: '2025-04-21', name: 'Lundi Pâques'},
      {date: '2025-05-01', name: 'Fête travail'},
      {date: '2025-05-05', name: 'Bouddha'},
      {date: '2025-05-31', name: 'Tuen Ng'},
      {date: '2025-07-01', name: 'Fête SAR HK'}
    ],
    '2025-2026': [
      {date: '2025-07-01', name: 'Fête SAR HK'},
      {date: '2025-10-01', name: 'Fête nationale'},
      {date: '2025-10-07', name: 'Après Mi-automne'},
      {date: '2025-10-29', name: 'Chung Yeung'},
      {date: '2025-12-25', name: 'Noël'},
      {date: '2025-12-26', name: 'Boxing Day'},
      {date: '2026-01-01', name: 'Jour de l\'an'},
      {date: '2026-02-17', name: 'Nouvel An chinois'},
      {date: '2026-02-18', name: '2e jour Nouvel An chinois'},
      {date: '2026-02-19', name: '3e jour Nouvel An chinois'},
      {date: '2026-04-03', name: 'Vendredi saint'},
      {date: '2026-04-04', name: 'Après Vendredi saint'},
      {date: '2026-04-06', name: 'Après Ching Ming'},
      {date: '2026-04-07', name: 'Après Lundi Pâques'},
      {date: '2026-05-01', name: 'Fête travail'},
      {date: '2026-05-25', name: 'Bouddha'},
      {date: '2026-06-19', name: 'Tuen Ng'},
      {date: '2026-07-01', name: 'Fête SAR HK'}
    ]
  },

  // Background colors for calendar cells
  colors: {
    saturday: '#E3F2FD',        // Light blue
    sunday: '#FFE0B2',          // Light orange
    schoolHoliday: '#E1BEE7',   // Light purple
    publicHoliday: '#FFCDD2'    // Light red
  }
};

// ============================================================================
// UTILITY FUNCTIONS
// ============================================================================

/**
 * Normalizes a string by removing accents for comparison
 * @param {string} str - String to normalize
 * @returns {string} Lowercase string without accents
 */
function normalizeForComparison(str) {
  return String(str).toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g, '');
}

/**
 * Checks if a header matches "Evenement" (with any accent variation)
 * Matches: Évènement, Événement, Évenement, evenement, etc.
 * @param {string} header - Header to check
 * @returns {boolean}
 */
function isEventColumn(header) {
  return normalizeForComparison(header) === 'evenement';
}

/**
 * Checks if a header matches "Sur site CCFHK?" (flexible matching)
 * @param {string} header - Header to check
 * @returns {boolean}
 */
function isSurSiteColumn(header) {
  const normalized = normalizeForComparison(header);
  return normalized.includes('sur site ccfhk') || normalized === 'sur site ccfhk?';
}

/**
 * Checks if a header matches "Sur calendrier excel?" (flexible matching)
 * @param {string} header - Header to check
 * @returns {boolean}
 */
function isSurCalendrierColumn(header) {
  const normalized = normalizeForComparison(header);
  return normalized.includes('sur calendrier excel') || normalized === 'sur calendrier excel?';
}

/**
 * Converts a cell value to a boolean (empty = false/No)
 * @param {*} value - Cell value
 * @returns {boolean}
 */
function toBooleanFilter(value) {
  if (!value) return false;
  const str = String(value).toLowerCase().trim();
  return str === 'oui' || str === 'yes' || str === 'true' || str === '1' || str === 'x';
}

/**
 * Gets the emoji prefix for an event based on Sur calendrier excel status
 * @param {Object} event - Event object with surCalendrierExcel and surCalendrierExcelRaw
 * @param {string} deptEmoji - The department emoji
 * @returns {string} Emoji prefix (dept emoji, or ❓+dept for blank, or 🙈+dept for Non)
 */
function getCalendarStatusPrefix(event, deptEmoji) {
  if (event.surCalendrierExcel) {
    // Marked as Oui - just show department emoji
    return deptEmoji;
  }
  // Not on calendar - check if blank or explicit Non
  if (event.surCalendrierExcelRaw === '' || event.surCalendrierExcelRaw === undefined) {
    // Blank - show ❓ + dept emoji
    return '❓' + deptEmoji;
  }
  // Explicit Non - show 🙈 + dept emoji
  return '🙈' + deptEmoji;
}

/**
 * Gets the current academic year from the calendar filter
 * @returns {string} Academic year like "2025-2026"
 */
function getCurrentAcademicYear() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const calSheet = ss.getSheetByName(CONFIG.calendarSheetName);
  if (!calSheet) return '2025-2026'; // default

  const yearFilter = calSheet.getRange('B1').getValue();
  return yearFilter || '2025-2026';
}

/**
 * Checks if a date is a school holiday
 * @param {Date} date - Date to check
 * @param {string} academicYear - Academic year like "2025-2026"
 * @returns {string|null} Holiday name if date is a school holiday, null otherwise
 */
function getSchoolHoliday(date, academicYear) {
  const holidays = CONFIG.schoolHolidays[academicYear];
  if (!holidays) return null;

  const dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  for (const period of holidays) {
    if (dateStr >= period.start && dateStr <= period.end) {
      return period.name;
    }
  }
  return null;
}

/**
 * Checks if a date is a public holiday
 * @param {Date} date - Date to check
 * @param {string} academicYear - Academic year like "2025-2026"
 * @returns {string|null} Holiday name if date is a public holiday, null otherwise
 */
function getPublicHoliday(date, academicYear) {
  const holidays = CONFIG.publicHolidays[academicYear];
  if (!holidays) return null;

  const dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  for (const holiday of holidays) {
    if (holiday.date === dateStr) {
      return holiday.name;
    }
  }
  return null;
}

/**
 * Gets the colored circle emoji for a department
 * @param {string} department - Department name
 * @param {Object} emojiMap - Map of department name to emoji (optional, will build if not provided)
 * @returns {string} Colored circle emoji
 */
function getDepartmentEmoji(department, emojiMap) {
  if (emojiMap && emojiMap[department]) {
    return emojiMap[department];
  }
  // Fallback: build map on the fly
  const departments = getDepartmentNames().sort();
  const index = departments.indexOf(department);
  if (index >= 0) {
    return CONFIG.departmentEmojis[index % CONFIG.departmentEmojis.length];
  }
  return '⚪'; // Default white circle for unknown departments
}

/**
 * Builds emoji map for departments
 * @returns {Object} Department name -> emoji
 */
function buildEmojiMap() {
  const departments = getDepartmentNames().sort();
  const emojiMap = {};

  departments.forEach((dept, index) => {
    emojiMap[dept] = CONFIG.departmentEmojis[index % CONFIG.departmentEmojis.length];
  });

  return emojiMap;
}

// ============================================================================
// ERROR TRACKING
// ============================================================================

/**
 * Gets or creates the error tracking sheet
 * @returns {Sheet} The error sheet
 */
function getErrorSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let errorSheet = ss.getSheetByName(CONFIG.errorSheetName);

  if (!errorSheet) {
    errorSheet = ss.insertSheet(CONFIG.errorSheetName);
    // Set up headers
    errorSheet.getRange('A1:F1').setValues([['Horodatage', 'Type', 'Département', 'Ligne', 'Description', 'Détails']]);
    errorSheet.getRange('A1:F1').setFontWeight('bold').setBackground('#F3F4F6');
    // Set column widths
    errorSheet.setColumnWidth(1, 150); // Horodatage
    errorSheet.setColumnWidth(2, 100); // Type
    errorSheet.setColumnWidth(3, 120); // Département
    errorSheet.setColumnWidth(4, 60);  // Ligne
    errorSheet.setColumnWidth(5, 250); // Description
    errorSheet.setColumnWidth(6, 300); // Détails
    errorSheet.setFrozenRows(1);
  }

  return errorSheet;
}

/**
 * Logs an error to the error sheet
 * @param {string} type - Error type (e.g., 'DONNÉES', 'SYNC_GCAL', 'PARSING')
 * @param {string} department - Department name (or empty)
 * @param {number|string} row - Row number (or empty)
 * @param {string} description - Brief error description
 * @param {string} details - Additional details
 */
function logError(type, department, row, description, details) {
  try {
    const errorSheet = getErrorSheet();
    const timestamp = new Date().toLocaleString(CONFIG.locale);
    errorSheet.appendRow([timestamp, type, department || '', row || '', description, details || '']);
  } catch (e) {
    // Fallback to Logger if error sheet fails
    Logger.log(`ERROR [${type}]: ${description} - ${details}`);
  }
}

/**
 * Clears all errors from the error sheet (keeps header)
 */
function clearErrors() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const errorSheet = ss.getSheetByName(CONFIG.errorSheetName);

  if (errorSheet) {
    const lastRow = errorSheet.getLastRow();
    if (lastRow > 1) {
      errorSheet.getRange(2, 1, lastRow - 1, 6).clear();
    }
  }

  // Show confirmation
  try {
    SpreadsheetApp.getUi().alert('Erreurs effacées', 'Toutes les erreurs ont été supprimées.', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (e) {
    Logger.log('Erreurs effacées');
  }
}

/**
 * Shows the error sheet
 */
function showErrors() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let errorSheet = ss.getSheetByName(CONFIG.errorSheetName);

  if (!errorSheet) {
    errorSheet = getErrorSheet();
  }

  ss.setActiveSheet(errorSheet);
}

/**
 * Gets error count
 * @returns {number} Number of errors logged
 */
function getErrorCount() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const errorSheet = ss.getSheetByName(CONFIG.errorSheetName);

  if (!errorSheet) {
    return 0;
  }

  return Math.max(0, errorSheet.getLastRow() - 1);
}

/**
 * Clears errors of a specific type from the error sheet
 * Used at the start of operations to clear stale errors of that type
 * @param {string} type - Error type to clear (e.g., 'DONNÉES', 'SYNC_GCAL')
 */
function clearErrorsByType(type) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const errorSheet = ss.getSheetByName(CONFIG.errorSheetName);

  if (!errorSheet) {
    return;
  }

  const lastRow = errorSheet.getLastRow();
  if (lastRow <= 1) {
    return; // Only header, nothing to clear
  }

  // Get all data (skip header)
  const data = errorSheet.getRange(2, 1, lastRow - 1, 6).getValues();

  // Filter out rows matching the type
  const rowsToKeep = data.filter(row => row[1] !== type);

  // Clear the data area
  errorSheet.getRange(2, 1, lastRow - 1, 6).clear();

  // Write back rows to keep
  if (rowsToKeep.length > 0) {
    errorSheet.getRange(2, 1, rowsToKeep.length, 6).setValues(rowsToKeep);
  }
}

// ============================================================================
// DEBUG FUNCTIONS
// ============================================================================

/**
 * Debug function to check which sheets are detected as departments
 * Run this from Apps Script editor and check View > Logs
 */
function debugSheetDetection() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();

  Logger.log('=== Sheet Detection Debug ===');

  sheets.forEach(sheet => {
    const name = sheet.getName();

    // Skip system sheets
    if (name === CONFIG.calendarSheetName ||
        name === CONFIG.syncTrackingSheetName ||
        name === CONFIG.errorSheetName) {
      Logger.log(name + ': [SYSTEM SHEET - SKIPPED]');
      return;
    }

    const lastCol = sheet.getLastColumn();
    if (lastCol < 1) {
      Logger.log(name + ': [EMPTY SHEET]');
      return;
    }

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());

    const hasDate = headers.includes(CONFIG.dateColumn);
    const hasDebut = headers.includes(CONFIG.startColumn);
    const hasService = headers.includes('Service');
    const hasEvent = headers.some(h => isEventColumn(h));

    const eventColName = headers.find(h => isEventColumn(h)) || 'NOT FOUND';

    Logger.log(name + ':');
    Logger.log('  - Has Date: ' + hasDate + ' | Has Début: ' + hasDebut);
    Logger.log('  - Has Service: ' + hasService);
    Logger.log('  - Has Event column: ' + hasEvent + ' (' + eventColName + ')');
    Logger.log('  - VALID DEPARTMENT: ' + ((hasDate || hasDebut) && hasService && hasEvent));
  });

  Logger.log('=== Detected Departments ===');
  const depts = getDepartmentNames(true);
  Logger.log('Count: ' + depts.length);
  Logger.log('Names: ' + JSON.stringify(depts));
}

// ============================================================================
// MAIN ENTRY POINTS
// ============================================================================

/**
 * Creates the Calendar menu when the spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('Calendrier')
    .addItem('Mettre à jour le calendrier', 'renderCalendar')
    .addItem('Filtres par défaut', 'resetFilters')
    .addItem('Installer le calendrier', 'installCalendar')
    .addSeparator()
    .addItem('Synchroniser Google Calendar', 'syncToGoogleCalendar')
    .addSeparator()
    .addItem('Voir les erreurs', 'showErrors')
    .addItem('Effacer les erreurs', 'clearErrors')
    .addSeparator()
    .addItem('Configurer rafraîchissement auto', 'setupAutoRefresh')
    .addItem('Désactiver rafraîchissement auto', 'removeAutoRefresh')
    .addToUi();
}

/**
 * Resets all filters to default values and refreshes the calendar
 * Defaults: 2025-2026, Actuel + À venir, Tous services, Sur site CCFHK: Tous, Sur calendrier excel: Oui
 */
function resetFilters() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const calSheet = ss.getSheetByName(CONFIG.calendarSheetName);

  if (!calSheet) {
    throw new Error(`Feuille "${CONFIG.calendarSheetName}" introuvable.`);
  }

  // Reset to defaults
  // Col A-B: Year, Period, Service
  calSheet.getRange('B1').setValue('2025-2026');
  calSheet.getRange('B2').setValue('Actuel + À venir');
  calSheet.getRange('B3').setValue('Tous');
  // Col C-D: Sur calendrier, Sur site
  calSheet.getRange('D1').setValue('Oui');
  calSheet.getRange('D2').setValue('Tous');

  // Refresh calendar with new filter values
  renderCalendar();
}

/**
 * Handles edit events - triggers refresh when filters/checkbox change or department data changes
 */
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  const range = e.range;
  const sheetName = sheet.getName();

  // --- Calendar sheet edits (filters, checkboxes) ---
  if (sheetName === CONFIG.calendarSheetName) {
    const cell = range.getA1Notation();

    // Filter cells:
    // Col A-B: B1 (Year), B2 (Period), B3 (Service)
    // Col C-D: D1 (Sur calendrier), D2 (Sur site)
    const filterCells = ['B1', 'B2', 'B3', 'D1', 'D2'];

    // Reset checkbox (F1)
    if (cell === 'F1' && range.getValue() === true) {
      range.setValue(false);
      resetFilters();
      return;
    }

    // Refresh checkbox (F2)
    if (cell === 'F2' && range.getValue() === true) {
      range.setValue(false);
      renderCalendar();
      return;
    }

    // Filter dropdowns changed
    if (filterCells.includes(cell)) {
      renderCalendar();
    }
    return;
  }

  // --- Department sheet edits (event data changes) ---
  // Check if this is a valid department sheet (has required columns)
  const departments = getDepartmentNames();
  if (departments.includes(sheetName)) {
    // Trigger calendar refresh when department data changes
    renderCalendar();
    // Schedule debounced Google Calendar sync
    scheduleGoogleCalendarSync();
  }
}

/**
 * One-click installation: creates calendar sheet, dropdowns, and triggers
 */
function installCalendar() {
  // Setup the calendar sheet with filters
  setupCalendarSheet();

  // Setup auto-refresh trigger
  setupAutoRefresh();

  // Setup installable onEdit trigger (required for scheduleGoogleCalendarSync to work)
  setupOnEditTrigger();

  // Initial render
  renderCalendar();

  // Try to show alert (only works when run from spreadsheet UI)
  try {
    const ui = SpreadsheetApp.getUi();
    ui.alert(
      'Installation terminée',
      'Le calendrier a été installé avec succès!\n\n' +
      '• Utilisez les filtres pour personnaliser l\'affichage\n' +
      '• Le calendrier se rafraîchit automatiquement toutes les 5 minutes\n' +
      '• Utilisez Calendrier > Rafraîchir pour une mise à jour immédiate',
      ui.ButtonSet.OK
    );
  } catch (e) {
    // Running from script editor - just log success
    Logger.log('Installation terminée! Retournez à votre feuille de calcul pour voir le calendrier.');
  }
}

/**
 * Main render function - orchestrates the calendar generation
 */
function renderCalendar() {
  // Clear caches at start of render
  clearDepartmentCache();

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const calSheet = ss.getSheetByName(CONFIG.calendarSheetName);

  if (!calSheet) {
    throw new Error(`Feuille "${CONFIG.calendarSheetName}" introuvable. Exécutez d'abord installCalendar().`);
  }

  // Show loading message immediately (F3 = Sheets calendar status)
  calSheet.getRange('F3').setValue('En cours...');
  SpreadsheetApp.flush(); // Force immediate UI update

  // Read filter values
  const filters = readFilters(calSheet);

  // Get all events from department sheets
  const allEvents = getAllEvents();

  // Apply filters
  const filteredEvents = getFilteredEvents(allEvents, filters);

  // Generate calendar grid data
  const { gridData, formatInfo } = generateCalendarGrid(filteredEvents, filters);

  // Clear previous calendar content (keep filter rows)
  clearCalendarContent(calSheet);

  // Write calendar grid
  if (gridData.length > 0) {
    const startRow = 5;
    calSheet.getRange(startRow, 1, gridData.length, 7).setValues(gridData);

    // Apply formatting
    applyFormatting(calSheet, startRow, formatInfo, filteredEvents);
  }

  // Update timestamp (F3 = Sheets calendar status)
  calSheet.getRange('F3').setValue(new Date().toLocaleString(CONFIG.locale));
}

// ============================================================================
// DATA COLLECTION
// ============================================================================

/**
 * Checks if a datetime value has a time component (not just a date)
 * In Google Sheets, dates are stored as numbers where the fractional part is the time
 * @param {Date|number} value - Date object or sheet numeric value
 * @returns {boolean} True if the value has a time component
 */
function hasTimeComponent(value) {
  if (value instanceof Date) {
    // Check if hours, minutes, or seconds are non-zero
    return value.getHours() !== 0 || value.getMinutes() !== 0 || value.getSeconds() !== 0;
  }
  if (typeof value === 'number') {
    // Fractional part indicates time
    return (value % 1) !== 0;
  }
  return false;
}

/**
 * Parses a date/datetime value from a sheet cell
 * @param {*} value - Cell value (Date, number, or string)
 * @returns {Object|null} { date: Date, hasTime: boolean } or null if invalid
 */
function parseDateTimeValue(value) {
  if (!value) return null;

  let date;
  let hasTime = false;

  if (value instanceof Date) {
    date = value;
    hasTime = hasTimeComponent(value);
  } else if (typeof value === 'number') {
    // Sheet serial date
    hasTime = hasTimeComponent(value);
    date = new Date(value);
  } else {
    date = new Date(value);
  }

  if (isNaN(date.getTime())) return null;

  return { date, hasTime };
}

/**
 * Collects all events from all department sheets
 * Supports both legacy Date column and new Début/Fin columns
 * @returns {Array} Array of event objects
 */
function getAllEvents() {
  // Clear previous data errors before collecting new ones
  clearErrorsByType('DONNÉES');

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const events = [];

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    const sheetId = sheet.getSheetId();

    // Skip system sheets (calendar, sync tracking, errors)
    if (sheetName === CONFIG.calendarSheetName) return;
    if (sheetName === CONFIG.syncTrackingSheetName) return;
    if (sheetName === CONFIG.errorSheetName) return;

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return; // Need at least header + 1 row

    const headers = data[0].map(h => String(h).trim());

    // Find column indices - try new format first, fall back to legacy
    const startCol = headers.indexOf(CONFIG.startColumn);  // Début
    const endCol = headers.indexOf(CONFIG.endColumn);      // Fin
    const legacyDateCol = headers.indexOf(CONFIG.dateColumn); // Date (legacy)

    const serviceCol = headers.indexOf('Service');
    const eventCol = headers.findIndex(h => isEventColumn(h));

    // Find optional filter columns
    const surSiteCol = headers.findIndex(h => isSurSiteColumn(h));
    const surCalendrierCol = headers.findIndex(h => isSurCalendrierColumn(h));

    // Determine which date column to use
    const useLegacy = startCol === -1 && legacyDateCol !== -1;
    const dateColIndex = useLegacy ? legacyDateCol : startCol;

    // Skip sheets without required columns (need either Date or Début, plus Service and Event)
    if (dateColIndex === -1 || serviceCol === -1 || eventCol === -1) {
      return;
    }

    // Process data rows
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const rowNum = i + 1; // 1-indexed for display

      const service = String(row[serviceCol] || '').trim();
      const eventName = String(row[eventCol] || '').trim();
      const rawDateValue = row[dateColIndex];

      // Parse start date/time
      const startParsed = parseDateTimeValue(rawDateValue);

      // Log error if event name exists but date is invalid/missing
      if (!startParsed && eventName) {
        logError(
          'DONNÉES',
          sheetName,
          rowNum,
          'Date manquante ou invalide',
          `Événement: "${eventName}", Valeur date: "${rawDateValue || '(vide)'}"`
        );
        continue;
      }

      // Skip empty rows (no date)
      if (!startParsed) continue;

      // Parse end date/time (if column exists and has value)
      let endParsed = null;
      if (!useLegacy && endCol !== -1 && row[endCol]) {
        endParsed = parseDateTimeValue(row[endCol]);
        // Log if end date is invalid but present
        if (!endParsed) {
          logError(
            'DONNÉES',
            sheetName,
            rowNum,
            'Date de fin invalide',
            `Événement: "${eventName}", Valeur fin: "${row[endCol]}"`
          );
        }
      }

      // Read optional filter columns (empty = false/No)
      const surSiteCCFHK = surSiteCol !== -1 ? toBooleanFilter(row[surSiteCol]) : false;
      const surCalendrierExcel = surCalendrierCol !== -1 ? toBooleanFilter(row[surCalendrierCol]) : false;
      // Track raw value for Sur calendrier excel (to distinguish blank vs explicit Non)
      const surCalendrierExcelRaw = surCalendrierCol !== -1 ? String(row[surCalendrierCol] || '').trim().toLowerCase() : '';

      // Log error if date exists but event name is missing
      if (!eventName) {
        logError(
          'DONNÉES',
          sheetName,
          rowNum,
          'Nom d\'événement manquant',
          `Date: ${startParsed.date.toLocaleDateString(CONFIG.locale)}, Service: "${service || '(vide)'}"`
        );
        continue;
      }

      events.push({
        // Legacy compatibility: keep 'date' for existing code that uses it
        date: startParsed.date,
        // New fields for datetime support
        startDate: startParsed.date,
        endDate: endParsed ? endParsed.date : null,
        hasTime: startParsed.hasTime || (endParsed && endParsed.hasTime),
        // Other fields
        service: service,
        event: eventName,
        department: sheetName,
        sheetId: sheetId,
        rowIndex: i + 1, // 1-indexed for user display
        surSiteCCFHK: surSiteCCFHK,
        surCalendrierExcel: surCalendrierExcel,
        surCalendrierExcelRaw: surCalendrierExcelRaw // 'oui', 'non', or '' (blank)
      });
    }
  });

  // Sort by start date
  events.sort((a, b) => a.startDate - b.startDate);

  return events;
}

// ============================================================================
// FILTERING
// ============================================================================

/**
 * Reads current filter values from the calendar sheet
 * @param {Sheet} calSheet - The calendar sheet
 * @returns {Object} Filter settings
 */
function readFilters(calSheet) {
  // Layout:
  // Col A-B: A1=Année B1=[year] | A2=Période B2=[period] | A3=Service B3=[dept]
  // Col C-D: C1=Sur calendrier D1=[oui/non] | C2=Sur site D2=[oui/non]
  return {
    year: calSheet.getRange('B1').getValue() || '2025-2026',
    timeRange: calSheet.getRange('B2').getValue() || 'Tout',
    department: calSheet.getRange('B3').getValue() || 'Tous',
    surCalendrierExcel: calSheet.getRange('D1').getValue() || 'Oui',
    surSiteCCFHK: calSheet.getRange('D2').getValue() || 'Tous',
    search: '' // Search removed from UI
  };
}

/**
 * Filters events based on user-selected criteria
 * @param {Array} allEvents - All events
 * @param {Object} filters - Filter settings
 * @returns {Array} Filtered events
 */
function getFilteredEvents(allEvents, filters) {
  // Parse academic year (e.g., "2025-2026" -> Aug 2025 to Jul 2026)
  const yearRange = parseAcademicYear(filters.year);

  return allEvents.filter(event => {
    // Year filter (academic year: Aug to Jul)
    if (event.date < yearRange.start || event.date > yearRange.end) {
      return false;
    }

    // Time range filter
    if (!passesTimeRangeFilter(event, filters.timeRange, yearRange)) {
      return false;
    }

    // Department filter
    if (filters.department !== 'Tous' && event.department !== filters.department) {
      return false;
    }

    // Search filter
    if (filters.search) {
      const searchText = (event.service + ' ' + event.event).toLowerCase();
      if (!searchText.includes(filters.search)) {
        return false;
      }
    }

    // Sur site CCFHK filter (Tous/Oui/Non)
    if (filters.surSiteCCFHK && filters.surSiteCCFHK !== 'Tous') {
      const wantSurSite = filters.surSiteCCFHK === 'Oui';
      if (event.surSiteCCFHK !== wantSurSite) {
        return false;
      }
    }

    // Sur calendrier excel filter (Tous/Oui/Non)
    if (filters.surCalendrierExcel && filters.surCalendrierExcel !== 'Tous') {
      const wantSurCalendrier = filters.surCalendrierExcel === 'Oui';
      if (event.surCalendrierExcel !== wantSurCalendrier) {
        return false;
      }
    }

    return true;
  });
}

/**
 * Parses academic year string to date range
 * @param {string} yearStr - e.g., "2025-2026"
 * @returns {Object} { start: Date, end: Date }
 */
function parseAcademicYear(yearStr) {
  const match = yearStr.match(/(\d{4})-(\d{4})/);
  if (!match) {
    // Default to current academic year
    const now = new Date();
    const year = now.getMonth() >= 7 ? now.getFullYear() : now.getFullYear() - 1;
    return {
      start: new Date(year, 7, 1), // Aug 1
      end: new Date(year + 1, 6, 31, 23, 59, 59) // Jul 31
    };
  }

  const startYear = parseInt(match[1]);
  return {
    start: new Date(startYear, 7, 1), // Aug 1
    end: new Date(startYear + 1, 6, 31, 23, 59, 59) // Jul 31
  };
}

/**
 * Checks if event passes time range filter
 * @param {Object} event - Event object
 * @param {string} timeRange - Filter value
 * @param {Object} yearRange - Academic year range
 * @returns {boolean}
 */
function passesTimeRangeFilter(event, timeRange, yearRange) {
  if (timeRange === 'Tout') {
    return true;
  }

  if (timeRange === 'Actuel + À venir') {
    // Show all events from the START of current month (not just today)
    const now = new Date();
    const startOfMonth = new Date(now.getFullYear(), now.getMonth(), 1);
    return event.date >= startOfMonth;
  }

  // Specific month filter (e.g., "septembre")
  const monthIndex = CONFIG.monthNames.indexOf(timeRange.toLowerCase());
  if (monthIndex !== -1) {
    // Convert academic year month index to calendar month
    // Academic: 0=Aug, 1=Sep, ..., 11=Jul
    // Calendar: 0=Jan, ..., 7=Aug, ..., 11=Dec
    const calendarMonth = (monthIndex + 7) % 12;

    // Determine the year for this month
    const eventMonth = event.date.getMonth();
    const eventYear = event.date.getFullYear();

    // Check if month matches
    if (eventMonth !== calendarMonth) {
      return false;
    }

    // Check if within academic year
    return event.date >= yearRange.start && event.date <= yearRange.end;
  }

  return true;
}

// ============================================================================
// CALENDAR GRID GENERATION
// ============================================================================

/**
 * Generates the calendar grid data
 * @param {Array} filteredEvents - Filtered events
 * @param {Object} filters - Current filters
 * @returns {Object} { gridData: Array, formatInfo: Array }
 */
function generateCalendarGrid(filteredEvents, filters) {
  const gridData = [];
  const formatInfo = []; // Tracks row types for formatting

  // Determine which months to render
  const months = getMonthsToRender(filteredEvents, filters);

  if (months.length === 0) {
    gridData.push(['Aucun événement trouvé', '', '', '', '', '', '']);
    formatInfo.push({ type: 'message', row: 0 });
    return { gridData, formatInfo };
  }

  // Group events by day for quick lookup
  const eventsByDay = groupEventsByDay(filteredEvents);

  // Build emoji map once for all months
  const emojiMap = buildEmojiMap();

  months.forEach((monthInfo, monthIndex) => {
    const baseRow = gridData.length;

    // Add separator between months (except first)
    if (monthIndex > 0) {
      gridData.push(['', '', '', '', '', '', '']);
      formatInfo.push({ type: 'separator', row: baseRow });
    }

    // Month header
    const monthHeader = formatMonthHeader(monthInfo.year, monthInfo.month);
    gridData.push([monthHeader, '', '', '', '', '', '']);
    formatInfo.push({ type: 'monthHeader', row: gridData.length - 1 });

    // Weekday headers
    gridData.push([...CONFIG.weekdayNames]);
    formatInfo.push({ type: 'weekdayHeader', row: gridData.length - 1 });

    // Generate day grid
    const monthGrid = generateMonthGrid(monthInfo.year, monthInfo.month, eventsByDay, emojiMap);
    monthGrid.forEach((weekRow, weekIndex) => {
      gridData.push(weekRow.map(cell => cell.display));
      formatInfo.push({
        type: 'dayRow',
        row: gridData.length - 1,
        cells: weekRow
      });
    });
  });

  return { gridData, formatInfo };
}

/**
 * Determines which months to render based on filters
 * @param {Array} filteredEvents - Filtered events
 * @param {Object} filters - Current filters
 * @returns {Array} Array of { year, month } objects
 */
function getMonthsToRender(filteredEvents, filters) {
  const yearRange = parseAcademicYear(filters.year);

  // If specific month selected, show only that month
  const monthIndex = CONFIG.monthNames.indexOf(filters.timeRange.toLowerCase());
  if (monthIndex !== -1) {
    const calendarMonth = (monthIndex + 7) % 12;
    const year = calendarMonth >= 7 ? yearRange.start.getFullYear() : yearRange.end.getFullYear();
    return [{ year, month: calendarMonth }];
  }

  // For "Tout" or "Actuel + À venir", show all months with events
  if (filteredEvents.length === 0) {
    return [];
  }

  // Collect unique months from events
  const monthSet = new Set();
  filteredEvents.forEach(event => {
    const key = `${event.date.getFullYear()}-${event.date.getMonth()}`;
    monthSet.add(key);
  });

  // Convert to array and sort
  const months = Array.from(monthSet).map(key => {
    const [year, month] = key.split('-').map(Number);
    return { year, month };
  });

  // Sort by academic year order (Aug first)
  months.sort((a, b) => {
    // Convert to academic year order for sorting
    const aOrder = a.month >= 7 ? a.month - 7 : a.month + 5;
    const bOrder = b.month >= 7 ? b.month - 7 : b.month + 5;

    if (a.year !== b.year) {
      return a.year - b.year;
    }
    return aOrder - bOrder;
  });

  return months;
}

/**
 * Groups events by day for quick lookup
 * @param {Array} events - Events array
 * @returns {Object} Map of "YYYY-MM-DD" -> events array
 */
function groupEventsByDay(events) {
  const grouped = {};

  events.forEach(event => {
    const key = formatDateKey(event.date);
    if (!grouped[key]) {
      grouped[key] = [];
    }
    grouped[key].push(event);
  });

  return grouped;
}

/**
 * Formats date as key string
 * @param {Date} date
 * @returns {string} "YYYY-MM-DD"
 */
function formatDateKey(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

/**
 * Formats month header text
 * @param {number} year
 * @param {number} month - 0-indexed calendar month
 * @returns {string}
 */
function formatMonthHeader(year, month) {
  const date = new Date(year, month, 1);
  // Capitalize first letter
  const monthName = date.toLocaleString(CONFIG.locale, { month: 'long' });
  return monthName.charAt(0).toUpperCase() + monthName.slice(1) + ' ' + year;
}

/**
 * Generates grid rows for a single month
 * @param {number} year
 * @param {number} month - 0-indexed
 * @param {Object} eventsByDay - Events grouped by day
 * @param {Object} emojiMap - Map of department name to emoji
 * @returns {Array} Array of week rows, each containing 7 cell objects
 */
function generateMonthGrid(year, month, eventsByDay, emojiMap) {
  const firstDay = new Date(year, month, 1);
  const lastDay = new Date(year, month + 1, 0);
  // Convert Sunday=0 to Monday-first format (Mon=0, Tue=1, ..., Sun=6)
  const jsDayOfWeek = firstDay.getDay(); // 0 = Sunday in JS
  const startingDayOfWeek = jsDayOfWeek === 0 ? 6 : jsDayOfWeek - 1; // Convert to Mon=0
  const daysInMonth = lastDay.getDate();

  // Check if this month contains today
  const today = new Date();
  const isCurrentMonth = (today.getFullYear() === year && today.getMonth() === month);
  const todayDate = today.getDate();

  const weeks = [];
  let currentDay = 1;

  // Generate up to 6 weeks
  for (let week = 0; week < 6; week++) {
    const weekRow = [];

    for (let dayOfWeek = 0; dayOfWeek < 7; dayOfWeek++) {
      // Check if we should place a day
      const shouldPlaceDay = (week === 0 && dayOfWeek >= startingDayOfWeek) ||
                             (week > 0 && currentDay <= daysInMonth);

      if (shouldPlaceDay && currentDay <= daysInMonth) {
        const date = new Date(year, month, currentDay);
        const dateKey = formatDateKey(date);
        const dayEvents = eventsByDay[dateKey] || [];
        const isToday = isCurrentMonth && currentDay === todayDate;

        weekRow.push(formatDayCell(currentDay, dayEvents, isToday, emojiMap, year, month));
        currentDay++;
      } else {
        weekRow.push({ display: '', events: [], dayNumber: null, isToday: false });
      }
    }

    weeks.push(weekRow);

    // Stop if we've placed all days
    if (currentDay > daysInMonth) break;
  }

  return weeks;
}

/**
 * Formats time for display (HH:MM format)
 * @param {Date} date - Date object
 * @returns {string} Time string like "14:30"
 */
function formatTimeForDisplay(date) {
  const hours = String(date.getHours()).padStart(2, '0');
  const minutes = String(date.getMinutes()).padStart(2, '0');
  return `${hours}:${minutes}`;
}

/**
 * Formats a single day cell
 * @param {number} dayNumber
 * @param {Array} events - Events for this day
 * @param {boolean} isToday - Whether this is today's date
 * @param {Object} emojiMap - Map of department name to emoji
 * @param {number} year - Year of the date
 * @param {number} month - Month of the date (0-indexed)
 * @returns {Object} { display: string, events: Array, dayNumber: number, isToday: boolean, date: Date, schoolHoliday: string|null, publicHoliday: string|null }
 */
function formatDayCell(dayNumber, events, isToday, emojiMap, year, month) {
  let display = String(dayNumber);

  // Check for holidays
  const date = new Date(year, month, dayNumber);
  const currentAcademicYear = getCurrentAcademicYear();
  const schoolHoliday = getSchoolHoliday(date, currentAcademicYear);
  const publicHoliday = getPublicHoliday(date, currentAcademicYear);

  // Add holiday labels after day number
  if (publicHoliday) {
    display += '\nHK PH - ' + publicHoliday;
  } else if (schoolHoliday) {
    display += '\nVacances LFI';
  }

  if (events.length > 0) {
    const displayEvents = events.slice(0, CONFIG.maxEventsPerDay);

    displayEvents.forEach(event => {
      const deptEmoji = getDepartmentEmoji(event.department, emojiMap);
      // Get prefix based on Sur calendrier excel status: 🙈 for Non, ❓ for blank, dept emoji for Oui
      const prefix = getCalendarStatusPrefix(event, deptEmoji);

      // Build event text with optional time prefix
      let eventText = '';
      if (event.hasTime) {
        eventText = formatTimeForDisplay(event.startDate) + ' ';
      }
      eventText += event.service ? `${event.service} | ${event.event}` : event.event;

      display += '\n' + prefix + ' ' + eventText;
    });

    // Indicate overflow
    if (events.length > CONFIG.maxEventsPerDay) {
      display += `\n+ ${events.length - CONFIG.maxEventsPerDay} autres...`;
    }
  }

  return {
    display: display,
    events: events,
    dayNumber: dayNumber,
    isToday: isToday || false,
    date: date,
    schoolHoliday: schoolHoliday,
    publicHoliday: publicHoliday
  };
}

// ============================================================================
// SHEET SETUP
// ============================================================================

/**
 * Creates or resets the calendar sheet with filter controls
 */
function setupCalendarSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let calSheet = ss.getSheetByName(CONFIG.calendarSheetName);

  // Create sheet if doesn't exist
  if (!calSheet) {
    calSheet = ss.insertSheet(CONFIG.calendarSheetName);
  } else {
    // Clear existing content
    calSheet.clear();
  }

  // Get department names for dropdown
  const departments = getDepartmentNames();

  // ============================================================================
  // FILTER LAYOUT
  // ============================================================================
  // Col A-B: Année, Période, Service (time filters)
  // Col C-D: Sur calendrier, Sur site (content filters)
  // Col E-G: Actions + Timestamp
  // ============================================================================

  // --- COLUMN A-B: Time filters ---
  calSheet.getRange('A1').setValue('Année:');
  calSheet.getRange('A2').setValue('Période:');
  calSheet.getRange('A3').setValue('Service:');

  // --- COLUMN C-D: Content filters ---
  calSheet.getRange('C1').setValue('Sur calendrier:');
  calSheet.getRange('C2').setValue('Sur site:');

  // --- COLUMN E-F: Actions ---
  calSheet.getRange('E1').setValue('Filtres par défaut:');
  calSheet.getRange('F1').insertCheckboxes();
  calSheet.getRange('F1').setValue(false);
  calSheet.getRange('F1').setHorizontalAlignment('left');

  calSheet.getRange('E2').setValue('Mettre à jour:');
  calSheet.getRange('F2').insertCheckboxes();
  calSheet.getRange('F2').setValue(false);
  calSheet.getRange('F2').setHorizontalAlignment('left');

  // --- ROW 3-4: Status labels ---
  calSheet.getRange('E3').setValue('MAJ Calendrier Sheets:');
  calSheet.getRange('E4').setValue('MAJ Google Calendar:');

  // Year dropdown (B1)
  const currentYear = new Date().getFullYear();
  const yearOptions = [
    `${currentYear - 1}-${currentYear}`,
    `${currentYear}-${currentYear + 1}`,
    `${currentYear + 1}-${currentYear + 2}`
  ];
  const defaultYear = new Date().getMonth() >= 7
    ? `${currentYear}-${currentYear + 1}`
    : `${currentYear - 1}-${currentYear}`;

  const yearRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(yearOptions, true)
    .build();
  calSheet.getRange('B1').setDataValidation(yearRule).setValue(defaultYear);

  // Time range dropdown (B2)
  const timeOptions = ['Tout', 'Actuel + À venir', ...CONFIG.monthNames];
  const timeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(timeOptions, true)
    .build();
  calSheet.getRange('B2').setDataValidation(timeRule).setValue('Actuel + À venir');

  // Department dropdown (B3)
  const deptOptions = ['Tous', ...departments.sort()];
  const deptRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(deptOptions, true)
    .build();
  calSheet.getRange('B3').setDataValidation(deptRule).setValue('Tous');

  // Sur calendrier excel dropdown (D1) - default Oui
  const ouiNonOptions = ['Tous', 'Oui', 'Non'];
  const surCalendrierRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(ouiNonOptions, true)
    .build();
  calSheet.getRange('D1').setDataValidation(surCalendrierRule).setValue('Oui');

  // Sur site CCFHK dropdown (D2) - default Tous
  const surSiteRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(ouiNonOptions, true)
    .build();
  calSheet.getRange('D2').setDataValidation(surSiteRule).setValue('Tous');

  // Format filter area - labels bold
  calSheet.getRange('A1:A3').setFontWeight('bold');
  calSheet.getRange('C1:C2').setFontWeight('bold');
  calSheet.getRange('E1:E2').setFontWeight('bold');

  // Right-align action labels
  calSheet.getRange('E1:E2').setHorizontalAlignment('right');

  // Set column widths for better layout (calendar uses 7 columns A-G)
  calSheet.setColumnWidth(1, 140);  // A: Labels + Mon
  calSheet.setColumnWidth(2, 140);  // B: Dropdowns + Tue
  calSheet.setColumnWidth(3, 140);  // C: Labels + Wed
  calSheet.setColumnWidth(4, 140);  // D: Dropdowns + Thu
  calSheet.setColumnWidth(5, 140);  // E: Action labels + Fri
  calSheet.setColumnWidth(6, 140);  // F: Checkboxes + Sat
  calSheet.setColumnWidth(7, 140);  // G: Sun

  // Freeze filter rows
  calSheet.setFrozenRows(4);
}

// Cache for department names (cleared on each render)
let cachedDepartmentNames = null;

/**
 * Gets list of department sheet names (cached within a render cycle)
 * Supports both legacy (Date) and new (Début) column formats
 * @param {boolean} forceRefresh - Force refresh the cache
 * @returns {Array} Department names
 */
function getDepartmentNames(forceRefresh = false) {
  if (cachedDepartmentNames && !forceRefresh) {
    return cachedDepartmentNames;
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const departments = [];

  sheets.forEach(sheet => {
    const name = sheet.getName();
    // Skip system sheets
    if (name === CONFIG.calendarSheetName) return;
    if (name === CONFIG.syncTrackingSheetName) return;
    if (name === CONFIG.errorSheetName) return;

    // Only read first row for headers (optimization)
    const lastCol = sheet.getLastColumn();
    if (lastCol < 1) return;

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());

    // Check for required columns (flexible with accents for Evenement)
    // Accept either Date (legacy) or Début (new) column
    const hasDateOrDebut = headers.includes(CONFIG.dateColumn) || headers.includes(CONFIG.startColumn);
    const hasService = headers.includes('Service');
    const hasEvent = headers.some(h => isEventColumn(h));

    if (hasDateOrDebut && hasService && hasEvent) {
      departments.push(name);
    }
  });

  cachedDepartmentNames = departments;
  return departments;
}

/**
 * Clears the department names cache
 */
function clearDepartmentCache() {
  cachedDepartmentNames = null;
}

/**
 * Clears calendar content while preserving filter rows
 * @param {Sheet} calSheet
 */
function clearCalendarContent(calSheet) {
  const lastRow = calSheet.getMaxRows();
  if (lastRow > 4) {
    calSheet.getRange(5, 1, lastRow - 4, 7).clear();
  }
}

// ============================================================================
// FORMATTING
// ============================================================================

/**
 * Applies formatting to the calendar grid (optimized for performance)
 * @param {Sheet} calSheet
 * @param {number} startRow
 * @param {Array} formatInfo
 * @param {Array} filteredEvents
 */
function applyFormatting(calSheet, startRow, formatInfo, filteredEvents) {
  const emojiMap = buildEmojiMap();
  const ss = calSheet.getParent();
  const ssUrl = ss.getUrl();

  // Collect rows by type for batch operations
  const monthHeaderRows = [];
  const weekdayHeaderRows = [];
  const dayRows = [];
  const separatorRows = [];
  const messageRows = [];
  const todayCells = [];
  const richTextUpdates = [];

  // First pass: categorize rows and collect data
  formatInfo.forEach(info => {
    const actualRow = startRow + info.row;

    switch (info.type) {
      case 'monthHeader':
        monthHeaderRows.push(actualRow);
        break;
      case 'weekdayHeader':
        weekdayHeaderRows.push(actualRow);
        break;
      case 'dayRow':
        dayRows.push(actualRow);
        // Process cells for today highlight and rich text
        info.cells.forEach((cell, colIndex) => {
          if (cell.dayNumber !== null) {
            if (cell.isToday) {
              todayCells.push({ row: actualRow, col: colIndex + 1 });
            }
            if (cell.events.length > 0) {
              richTextUpdates.push({
                row: actualRow,
                col: colIndex + 1,
                cell: cell
              });
            }
          }
        });
        break;
      case 'separator':
        separatorRows.push(actualRow);
        break;
      case 'message':
        messageRows.push(actualRow);
        break;
    }
  });

  // Batch format month headers
  monthHeaderRows.forEach(row => {
    calSheet.getRange(row, 1, 1, 7)
      .merge()
      .setFontWeight('bold')
      .setFontSize(14)
      .setBackground('#E8EAED')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    calSheet.setRowHeight(row, 30);
  });

  // Batch format weekday headers
  weekdayHeaderRows.forEach(row => {
    calSheet.getRange(row, 1, 1, 7)
      .setFontWeight('bold')
      .setBackground('#F3F4F6')
      .setHorizontalAlignment('center');
    calSheet.setRowHeight(row, 25);
  });

  // Batch format day rows
  dayRows.forEach(row => {
    calSheet.getRange(row, 1, 1, 7)
      .setWrap(true)
      .setVerticalAlignment('top')
      .setFontSize(9);
    calSheet.setRowHeight(row, 120);
  });

  // Batch format separator rows
  separatorRows.forEach(row => {
    calSheet.setRowHeight(row, 10);
  });

  // Batch format message rows
  messageRows.forEach(row => {
    calSheet.getRange(row, 1, 1, 7)
      .merge()
      .setFontStyle('italic')
      .setFontColor('#666666')
      .setHorizontalAlignment('center');
  });

  // Apply background colors for weekends and holidays
  formatInfo.forEach(info => {
    if (info.type === 'dayRow' && info.cells) {
      const actualRow = startRow + info.row;
      info.cells.forEach((cell, colIndex) => {
        if (cell.dayNumber !== null) {
          const cellRange = calSheet.getRange(actualRow, colIndex + 1);

          // Determine background color based on day type
          const dayOfWeek = colIndex; // 0=Mon, 6=Sun
          const isSaturday = dayOfWeek === 5;
          const isSunday = dayOfWeek === 6;

          let backgroundColor = '#FFFFFF'; // default white

          // Priority: public holiday > school holiday on weekdays > weekends
          if (cell.publicHoliday) {
            backgroundColor = CONFIG.colors.publicHoliday;
          } else if (cell.schoolHoliday && !isSaturday && !isSunday) {
            backgroundColor = CONFIG.colors.schoolHoliday;
          } else if (isSaturday) {
            backgroundColor = CONFIG.colors.saturday;
          } else if (isSunday) {
            backgroundColor = CONFIG.colors.sunday;
          }

          cellRange.setBackground(backgroundColor);
        }
      });
    }
  });

  // Highlight today's cells (override background color with today's color)
  todayCells.forEach(({ row, col }) => {
    const cellRange = calSheet.getRange(row, col);
    cellRange.setBackground('#BBDEFB');
    cellRange.setBorder(true, true, true, true, false, false, '#1976D2', SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
  });

  // Apply rich text with hyperlinks (this is the slowest part, but unavoidable)
  richTextUpdates.forEach(({ row, col, cell }) => {
    const richText = buildRichTextWithLinksOptimized(cell, ssUrl, emojiMap);
    if (richText) {
      calSheet.getRange(row, col).setRichTextValue(richText);
    }
  });
}

/**
 * Builds rich text with hyperlinks (optimized - accepts pre-fetched URL)
 * @param {Object} cell - Cell object with events
 * @param {string} ssUrl - Spreadsheet URL (pre-fetched)
 * @param {Object} emojiMap - Map of department name to emoji
 * @returns {RichTextValue|null}
 */
function buildRichTextWithLinksOptimized(cell, ssUrl, emojiMap) {
  try {
    const builder = SpreadsheetApp.newRichTextValue();
    let text = String(cell.dayNumber);

    // Add holiday labels
    if (cell.publicHoliday) {
      text += '\nHK PH - ' + cell.publicHoliday;
    } else if (cell.schoolHoliday) {
      text += '\nVacances LFI';
    }

    const displayEvents = cell.events.slice(0, CONFIG.maxEventsPerDay);
    const linkPositions = [];

    displayEvents.forEach(event => {
      const deptEmoji = emojiMap[event.department] || '⚪';
      // Get prefix based on Sur calendrier excel status: 🙈 for Non, ❓ for blank, dept emoji for Oui
      const prefix = getCalendarStatusPrefix(event, deptEmoji);

      // Build event text with optional time prefix
      let eventText = '';
      if (event.hasTime) {
        eventText = formatTimeForDisplay(event.startDate) + ' ';
      }
      eventText += event.service ? `${event.service} | ${event.event}` : event.event;

      const fullEventText = prefix + ' ' + eventText;
      const startPos = text.length + 1;
      text += '\n' + fullEventText;
      const endPos = text.length;

      const linkStart = startPos + prefix.length + 1;
      linkPositions.push({
        start: linkStart,
        end: endPos,
        sheetId: event.sheetId
      });
    });

    if (cell.events.length > CONFIG.maxEventsPerDay) {
      text += `\n+ ${cell.events.length - CONFIG.maxEventsPerDay} autres...`;
    }

    builder.setText(text);

    // Add hyperlinks
    const linkStyle = SpreadsheetApp.newTextStyle()
      .setUnderline(true)
      .setForegroundColor('#1155CC')
      .build();

    linkPositions.forEach(link => {
      const url = `${ssUrl}#gid=${link.sheetId}`;
      builder.setTextStyle(link.start, link.end, linkStyle);
      builder.setLinkUrl(link.start, link.end, url);
    });

    return builder.build();
  } catch (e) {
    return null;
  }
}

// ============================================================================
// AUTO-REFRESH TRIGGERS
// ============================================================================

/**
 * Sets up automatic refresh trigger
 */
function setupAutoRefresh() {
  // Remove existing triggers first
  removeAutoRefresh();

  // Create new time-based trigger
  ScriptApp.newTrigger('renderCalendar')
    .timeBased()
    .everyMinutes(CONFIG.refreshIntervalMinutes)
    .create();
}

/**
 * Removes automatic refresh trigger
 */
function removeAutoRefresh() {
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'renderCalendar') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}

/**
 * Sets up installable onEdit trigger
 * Required because simple onEdit triggers cannot create other triggers
 */
function setupOnEditTrigger() {
  // Remove existing onEdit triggers first
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'onEdit') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Create new installable onEdit trigger
  ScriptApp.newTrigger('onEdit')
    .forSpreadsheet(SpreadsheetApp.getActive())
    .onEdit()
    .create();
}

// ============================================================================
// GOOGLE CALENDAR SYNC
// ============================================================================

/**
 * Gets the Google Calendar by ID from CONFIG
 * @returns {Calendar} Google Calendar object
 */
function getCalendar() {
  return CalendarApp.getCalendarById(CONFIG.googleCalendarId);
}

/**
 * Sets up the sync tracking sheet if it doesn't exist
 * @returns {Sheet} The sync tracking sheet
 */
function setupSyncSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let syncSheet = ss.getSheetByName(CONFIG.syncTrackingSheetName);

  if (!syncSheet) {
    syncSheet = ss.insertSheet(CONFIG.syncTrackingSheetName);
    // Set up headers
    syncSheet.getRange('A1:E1').setValues([['eventHash', 'gcalEventId', 'department', 'eventDate', 'lastSynced']]);
    syncSheet.getRange('A1:E1').setFontWeight('bold');
    // Hide the sheet (it's for internal tracking)
    syncSheet.hideSheet();
  }

  return syncSheet;
}

/**
 * Computes MD5 hash for an event (deterministic ID)
 * Includes start/end datetime for proper change detection
 * @param {Object} event - Event object with department, startDate, endDate, service, event
 * @returns {string} MD5 hash string
 */
function computeEventHash(event) {
  // Use ISO format for full datetime precision
  const startStr = event.startDate.toISOString();
  const endStr = event.endDate ? event.endDate.toISOString() : '';
  // Include surSiteCCFHK in hash so toggling it triggers resync
  const surSiteStr = event.surSiteCCFHK ? '1' : '0';
  const input = event.department + '|' + startStr + '|' + endStr + '|' + event.service + '|' + event.event + '|' + surSiteStr;
  const rawHash = Utilities.computeDigest(Utilities.DigestAlgorithm.MD5, input);
  // Convert to hex string
  return rawHash.map(byte => {
    const hex = (byte < 0 ? byte + 256 : byte).toString(16);
    return hex.length === 1 ? '0' + hex : hex;
  }).join('');
}

/**
 * Reads sync tracking data from the tracking sheet
 * @returns {Map} Map of eventHash -> { gcalEventId, department, eventDate, lastSynced }
 */
function getSyncTrackingData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const syncSheet = ss.getSheetByName(CONFIG.syncTrackingSheetName);

  if (!syncSheet) {
    return new Map();
  }

  const data = syncSheet.getDataRange().getValues();
  const trackingMap = new Map();

  // Skip header row
  for (let i = 1; i < data.length; i++) {
    const [hash, gcalId, dept, date, lastSynced] = data[i];
    if (hash) {
      trackingMap.set(hash, {
        gcalEventId: gcalId,
        department: dept,
        eventDate: date,
        lastSynced: lastSynced
      });
    }
  }

  return trackingMap;
}

/**
 * Writes sync tracking data to the tracking sheet
 * @param {Map} trackingMap - Map of eventHash -> tracking data
 */
function writeSyncTracking(trackingMap) {
  const syncSheet = setupSyncSheet();

  // Clear existing data (keep header)
  const lastRow = syncSheet.getLastRow();
  if (lastRow > 1) {
    syncSheet.getRange(2, 1, lastRow - 1, 5).clear();
  }

  // Write new data
  if (trackingMap.size > 0) {
    const rows = [];
    trackingMap.forEach((value, hash) => {
      rows.push([hash, value.gcalEventId, value.department, value.eventDate, value.lastSynced]);
    });
    syncSheet.getRange(2, 1, rows.length, 5).setValues(rows);
  }
}

/**
 * Creates a Google Calendar event
 * Supports timed events, all-day events, and multi-day events
 * @param {Calendar} calendar - Google Calendar
 * @param {Object} event - Event object with startDate, endDate, hasTime
 * @returns {string} Google Calendar event ID
 */
function createCalendarEvent(calendar, event) {
  const title = event.service ? `${event.service} | ${event.event}` : event.event;
  const description = `[[[${event.department}]]]\n\nSource: CCFHK Events`;

  let gcalEvent;

  if (event.hasTime) {
    // Timed event
    let endTime;
    if (event.endDate) {
      endTime = event.endDate;
    } else {
      // Default duration if no end time specified
      endTime = new Date(event.startDate.getTime() + CONFIG.defaultEventDurationHours * 60 * 60 * 1000);
    }
    gcalEvent = calendar.createEvent(title, event.startDate, endTime, {
      description: description
    });
  } else {
    // All-day event
    if (event.endDate) {
      // Multi-day all-day event
      // Google Calendar end date is exclusive, so add 1 day
      const exclusiveEnd = new Date(event.endDate);
      exclusiveEnd.setDate(exclusiveEnd.getDate() + 1);
      gcalEvent = calendar.createAllDayEvent(title, event.startDate, exclusiveEnd, {
        description: description
      });
    } else {
      // Single day all-day event
      gcalEvent = calendar.createAllDayEvent(title, event.startDate, {
        description: description
      });
    }
  }

  return gcalEvent.getId();
}

/**
 * Updates a Google Calendar event
 * Supports timed events, all-day events, and multi-day events
 * @param {Calendar} calendar - Google Calendar
 * @param {string} gcalEventId - Google Calendar event ID
 * @param {Object} event - Event object with startDate, endDate, hasTime
 * @returns {boolean} True if updated successfully
 */
function updateCalendarEvent(calendar, gcalEventId, event) {
  try {
    const gcalEvent = calendar.getEventById(gcalEventId);
    if (!gcalEvent) {
      return false;
    }

    const title = event.service ? `${event.service} | ${event.event}` : event.event;
    const description = `[[[${event.department}]]]\n\nSource: CCFHK Events`;

    gcalEvent.setTitle(title);
    gcalEvent.setDescription(description);

    if (event.hasTime) {
      // Timed event
      let endTime;
      if (event.endDate) {
        endTime = event.endDate;
      } else {
        // Default duration if no end time specified
        endTime = new Date(event.startDate.getTime() + CONFIG.defaultEventDurationHours * 60 * 60 * 1000);
      }
      gcalEvent.setTime(event.startDate, endTime);
    } else {
      // All-day event
      if (event.endDate) {
        // Multi-day all-day event
        // Google Calendar end date is exclusive, so add 1 day
        const exclusiveEnd = new Date(event.endDate);
        exclusiveEnd.setDate(exclusiveEnd.getDate() + 1);
        gcalEvent.setAllDayDates(event.startDate, exclusiveEnd);
      } else {
        // Single day all-day event
        gcalEvent.setAllDayDate(event.startDate);
      }
    }

    return true;
  } catch (e) {
    Logger.log('Error updating event: ' + e.message);
    return false;
  }
}

/**
 * Deletes a Google Calendar event
 * @param {Calendar} calendar - Google Calendar
 * @param {string} gcalEventId - Google Calendar event ID
 */
function deleteCalendarEvent(calendar, gcalEventId) {
  try {
    const gcalEvent = calendar.getEventById(gcalEventId);
    if (gcalEvent) {
      gcalEvent.deleteEvent();
    }
  } catch (e) {
    Logger.log('Error deleting event: ' + e.message);
  }
}

/**
 * Main sync function - syncs events with Sur site CCFHK = Oui to Google Calendar
 * Uses incremental sync with create/update/delete operations
 */
function syncToGoogleCalendar() {
  // Clear previous sync errors before starting new sync
  clearErrorsByType('SYNC_GCAL');

  // Show loading state (F4 = Google Calendar status)
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const calSheet = ss.getSheetByName(CONFIG.calendarSheetName);
  if (calSheet) {
    calSheet.getRange('F4').setValue('En cours...');
    SpreadsheetApp.flush();
  }

  const calendar = getCalendar();
  if (!calendar) {
    throw new Error('Calendrier Google introuvable. Vérifiez l\'ID du calendrier dans CONFIG.');
  }

  // Get all events
  const allEvents = getAllEvents();

  // Filter for Sur site CCFHK = Oui only
  const eventsToSync = allEvents.filter(event => event.surSiteCCFHK === true);

  // Get current tracking data
  const trackingMap = getSyncTrackingData();
  const seenHashes = new Set();
  const newTrackingMap = new Map();

  let created = 0;
  let unchanged = 0;
  let deleted = 0;

  let syncErrors = 0;

  // Process each event to sync
  eventsToSync.forEach(event => {
    const hash = computeEventHash(event);
    seenHashes.add(hash);

    const existingEntry = trackingMap.get(hash);
    const eventTitle = event.service ? `${event.service} | ${event.event}` : event.event;

    try {
      if (existingEntry) {
        // Hash matches = event data unchanged, skip API call
        // Just keep the existing tracking entry
        newTrackingMap.set(hash, existingEntry);
        unchanged++;
      } else {
        // New event or event data changed - create in calendar
        const gcalEventId = createCalendarEvent(calendar, event);
        newTrackingMap.set(hash, {
          gcalEventId: gcalEventId,
          department: event.department,
          eventDate: Utilities.formatDate(event.date, 'GMT', 'yyyy-MM-dd'),
          lastSynced: new Date()
        });
        created++;
      }
    } catch (e) {
      // Log sync error
      logError(
        'SYNC_GCAL',
        event.department,
        event.rowIndex,
        'Échec synchronisation Google Calendar',
        `Événement: "${eventTitle}", Erreur: ${e.message}`
      );
      syncErrors++;
    }
  });

  // Delete events that are no longer in the sync list
  trackingMap.forEach((value, hash) => {
    if (!seenHashes.has(hash)) {
      try {
        deleteCalendarEvent(calendar, value.gcalEventId);
        deleted++;
      } catch (e) {
        logError(
          'SYNC_GCAL',
          value.department,
          '',
          'Échec suppression événement',
          `Date: ${value.eventDate}, Erreur: ${e.message}`
        );
        syncErrors++;
      }
    }
  });

  // Write updated tracking data
  writeSyncTracking(newTrackingMap);

  // Update Google Calendar sync timestamp on calendar sheet
  updateGoogleCalendarSyncTimestamp();

  Logger.log(`Sync complete: ${created} created, ${unchanged} unchanged, ${deleted} deleted, ${syncErrors} errors`);
  return { created, unchanged, deleted, errors: syncErrors };
}

/**
 * Updates the Google Calendar sync timestamp on the calendar sheet
 */
function updateGoogleCalendarSyncTimestamp() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const calSheet = ss.getSheetByName(CONFIG.calendarSheetName);
  if (calSheet) {
    calSheet.getRange('F4').setValue(new Date().toLocaleString(CONFIG.locale));
  }
}

/**
 * Schedules a debounced Google Calendar sync
 * Removes any pending sync trigger and schedules a new one
 */
function scheduleGoogleCalendarSync() {
  // Remove existing pending sync triggers
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'executeDebouncedSync') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Schedule new sync after debounce period
  ScriptApp.newTrigger('executeDebouncedSync')
    .timeBased()
    .after(CONFIG.syncDebounceSeconds * 1000)
    .create();
}

/**
 * Executes the debounced sync (called by trigger)
 */
function executeDebouncedSync() {
  // Remove the trigger that called us
  const triggers = ScriptApp.getProjectTriggers();
  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'executeDebouncedSync') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  // Execute sync
  syncToGoogleCalendar();
}
