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
  requiredColumns: ['Date', 'Service', 'Évènement'],
  maxEventsPerDay: 8,
  refreshIntervalMinutes: 5,
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
  ]
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

  // Show loading message immediately (E3 below Refresh checkbox)
  calSheet.getRange('E3').setValue('Mise à jour en cours...');
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

  // Update timestamp (E3 below Refresh checkbox)
  calSheet.getRange('E3').setValue('Mis à jour: ' + new Date().toLocaleString(CONFIG.locale));
}

// ============================================================================
// DATA COLLECTION
// ============================================================================

/**
 * Collects all events from all department sheets
 * @returns {Array} Array of event objects
 */
function getAllEvents() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const events = [];

  sheets.forEach(sheet => {
    const sheetName = sheet.getName();
    const sheetId = sheet.getSheetId();

    // Skip the calendar sheet itself
    if (sheetName === CONFIG.calendarSheetName) return;

    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return; // Need at least header + 1 row

    const headers = data[0].map(h => String(h).trim());

    // Find required column indices (flexible with accents for Evenement)
    const dateCol = headers.indexOf('Date');
    const serviceCol = headers.indexOf('Service');
    const eventCol = headers.findIndex(h => isEventColumn(h));

    // Find optional filter columns
    const surSiteCol = headers.findIndex(h => isSurSiteColumn(h));
    const surCalendrierCol = headers.findIndex(h => isSurCalendrierColumn(h));

    // Skip sheets without required columns
    if (dateCol === -1 || serviceCol === -1 || eventCol === -1) {
      return;
    }

    // Process data rows
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const dateValue = row[dateCol];

      // Skip empty dates
      if (!dateValue) continue;

      // Parse date
      let date;
      if (dateValue instanceof Date) {
        date = dateValue;
      } else {
        date = new Date(dateValue);
      }

      // Skip invalid dates
      if (isNaN(date.getTime())) continue;

      const service = String(row[serviceCol] || '').trim();
      const eventName = String(row[eventCol] || '').trim();

      // Read optional filter columns (empty = false/No)
      const surSiteCCFHK = surSiteCol !== -1 ? toBooleanFilter(row[surSiteCol]) : false;
      const surCalendrierExcel = surCalendrierCol !== -1 ? toBooleanFilter(row[surCalendrierCol]) : false;
      // Track raw value for Sur calendrier excel (to distinguish blank vs explicit Non)
      const surCalendrierExcelRaw = surCalendrierCol !== -1 ? String(row[surCalendrierCol] || '').trim().toLowerCase() : '';

      // Skip if no event name
      if (!eventName) continue;

      events.push({
        date: date,
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

  // Sort by date
  events.sort((a, b) => a.date - b.date);

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

        weekRow.push(formatDayCell(currentDay, dayEvents, isToday, emojiMap));
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
 * Formats a single day cell
 * @param {number} dayNumber
 * @param {Array} events - Events for this day
 * @param {boolean} isToday - Whether this is today's date
 * @param {Object} emojiMap - Map of department name to emoji
 * @returns {Object} { display: string, events: Array, dayNumber: number, isToday: boolean }
 */
function formatDayCell(dayNumber, events, isToday, emojiMap) {
  let display = String(dayNumber);

  if (events.length > 0) {
    const displayEvents = events.slice(0, CONFIG.maxEventsPerDay);

    displayEvents.forEach(event => {
      const deptEmoji = getDepartmentEmoji(event.department, emojiMap);
      // Get prefix based on Sur calendrier excel status: 🙈 for Non, ❓ for blank, dept emoji for Oui
      const prefix = getCalendarStatusPrefix(event, deptEmoji);
      const eventText = event.service ? `${event.service} | ${event.event}` : event.event;
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
    isToday: isToday || false
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
  // E3 will show timestamp

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
    if (name === CONFIG.calendarSheetName) return;

    // Only read first row for headers (optimization)
    const lastCol = sheet.getLastColumn();
    if (lastCol < 1) return;

    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(h => String(h).trim());

    // Check for required columns (flexible with accents for Evenement)
    const hasDate = headers.includes('Date');
    const hasService = headers.includes('Service');
    const hasEvent = headers.some(h => isEventColumn(h));

    if (hasDate && hasService && hasEvent) {
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

  // Highlight today's cells
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

    const displayEvents = cell.events.slice(0, CONFIG.maxEventsPerDay);
    const linkPositions = [];

    displayEvents.forEach(event => {
      const deptEmoji = emojiMap[event.department] || '⚪';
      // Get prefix based on Sur calendrier excel status: 🙈 for Non, ❓ for blank, dept emoji for Oui
      const prefix = getCalendarStatusPrefix(event, deptEmoji);
      const eventText = event.service ? `${event.service} | ${event.event}` : event.event;
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
