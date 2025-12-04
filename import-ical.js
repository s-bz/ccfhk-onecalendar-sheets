#!/usr/bin/env node
/**
 * ICAL to CSV Import Script
 *
 * Converts an ICAL file to CSV format compatible with the CCFHK calendar sheets.
 *
 * Usage:
 *   node import-ical.js <input.ics> [output.csv]
 *   bun import-ical.js <input.ics> [output.csv]
 *
 * Output columns: Début, Fin, Service, Évènement, Lieu, Description, UID
 */

const fs = require('fs');
const path = require('path');

// Hong Kong timezone offset (UTC+8)
const HK_OFFSET_MS = 8 * 60 * 60 * 1000;

/**
 * Parse command line arguments
 */
function parseArgs() {
  const args = process.argv.slice(2);
  if (args.length < 1) {
    console.error('Usage: node import-ical.js <input.ics> [output.csv]');
    process.exit(1);
  }

  const inputFile = args[0];
  const outputFile = args[1] || inputFile.replace(/\.ics$/i, '.csv');

  return { inputFile, outputFile };
}

/**
 * Unfold ICAL lines (lines starting with space/tab are continuations)
 */
function unfoldLines(text) {
  return text.replace(/\r\n[ \t]/g, '').replace(/\r\n/g, '\n').replace(/\n[ \t]/g, '');
}

/**
 * Unescape ICAL text values
 */
function unescapeValue(text) {
  if (!text) return '';
  return text
    .replace(/\\n/g, '\n')
    .replace(/\\,/g, ',')
    .replace(/\\;/g, ';')
    .replace(/\\\\/g, '\\')
    .replace(/rn/g, '\n'); // WordPress exports sometimes use 'rn' instead of \n
}

/**
 * Parse ICAL date/time string to Date object
 * Formats:
 *   - 20250510T100000Z (UTC)
 *   - 20250510T100000 (local, with TZID parameter)
 *   - 20250510 (all-day)
 */
function parseICALDate(dateStr, tzid = null) {
  if (!dateStr) return null;

  // All-day date: YYYYMMDD
  if (dateStr.length === 8) {
    const year = parseInt(dateStr.slice(0, 4));
    const month = parseInt(dateStr.slice(4, 6)) - 1;
    const day = parseInt(dateStr.slice(6, 8));
    return { date: new Date(year, month, day), isAllDay: true };
  }

  // DateTime: YYYYMMDDTHHMMSS or YYYYMMDDTHHMMSSZ
  const isUTC = dateStr.endsWith('Z');
  const cleanStr = dateStr.replace('Z', '').replace('T', '');

  const year = parseInt(cleanStr.slice(0, 4));
  const month = parseInt(cleanStr.slice(4, 6)) - 1;
  const day = parseInt(cleanStr.slice(6, 8));
  const hour = parseInt(cleanStr.slice(8, 10)) || 0;
  const minute = parseInt(cleanStr.slice(10, 12)) || 0;
  const second = parseInt(cleanStr.slice(12, 14)) || 0;

  let date;
  if (isUTC) {
    // UTC time - create as UTC then it will display correctly
    date = new Date(Date.UTC(year, month, day, hour, minute, second));
  } else if (tzid === 'Asia/Hong_Kong') {
    // Already in HK time - treat as local
    date = new Date(year, month, day, hour, minute, second);
  } else {
    // Assume local time
    date = new Date(year, month, day, hour, minute, second);
  }

  return { date, isAllDay: false };
}

/**
 * Format date for CSV output (HK timezone)
 */
function formatDateForCSV(dateObj, isAllDay) {
  if (!dateObj) return '';

  // For all-day events, just return the date
  if (isAllDay) {
    const year = dateObj.getFullYear();
    const month = String(dateObj.getMonth() + 1).padStart(2, '0');
    const day = String(dateObj.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
  }

  // For timed events, convert to HK time and format as datetime
  // Add HK offset to get HK local time display
  const hkDate = new Date(dateObj.getTime() + HK_OFFSET_MS);

  const year = hkDate.getUTCFullYear();
  const month = String(hkDate.getUTCMonth() + 1).padStart(2, '0');
  const day = String(hkDate.getUTCDate()).padStart(2, '0');
  const hour = String(hkDate.getUTCHours()).padStart(2, '0');
  const minute = String(hkDate.getUTCMinutes()).padStart(2, '0');

  return `${year}-${month}-${day} ${hour}:${minute}`;
}

/**
 * Extract department tag from description [[[DEPT-NAME]]]
 */
function extractDepartment(description) {
  if (!description) return { dept: '', cleanDesc: '' };

  const match = description.match(/\[\[\[([A-Z0-9-]+)\]\]\]/);
  const dept = match ? match[1] : '';
  const cleanDesc = description.replace(/\[\[\[[A-Z0-9-]+\]\]\]/, '').trim();

  return { dept, cleanDesc };
}

/**
 * Parse RRULE to get recurrence info
 * Only handles WEEKLY with UNTIL (which is what the file has)
 */
function parseRRule(rrule) {
  if (!rrule) return null;

  const parts = {};
  rrule.split(';').forEach(part => {
    const [key, value] = part.split('=');
    parts[key] = value;
  });

  if (parts.FREQ !== 'WEEKLY') return null;
  if (!parts.UNTIL) return null;

  const untilParsed = parseICALDate(parts.UNTIL);
  return {
    freq: 'WEEKLY',
    until: untilParsed ? untilParsed.date : null,
    interval: parseInt(parts.INTERVAL) || 1
  };
}

/**
 * Expand recurring event into individual occurrences
 */
function expandRecurring(event, rrule) {
  const recurrence = parseRRule(rrule);
  if (!recurrence || !recurrence.until) return [event];

  const events = [];
  const startDate = new Date(event.startDate);
  const endDate = event.endDate ? new Date(event.endDate) : null;
  const duration = endDate ? endDate.getTime() - startDate.getTime() : 0;

  let currentStart = new Date(startDate);
  const weekMs = 7 * 24 * 60 * 60 * 1000 * recurrence.interval;

  while (currentStart <= recurrence.until) {
    const occurrence = { ...event };
    occurrence.startDate = new Date(currentStart);
    if (endDate) {
      occurrence.endDate = new Date(currentStart.getTime() + duration);
    }
    events.push(occurrence);

    currentStart = new Date(currentStart.getTime() + weekMs);
  }

  return events;
}

/**
 * Parse a single VEVENT block
 */
function parseEvent(lines) {
  const event = {
    uid: '',
    summary: '',
    description: '',
    location: '',
    startDate: null,
    endDate: null,
    isAllDay: false,
    rrule: null
  };

  let startTzid = null;

  for (const line of lines) {
    if (line.startsWith('UID:')) {
      event.uid = line.slice(4);
    } else if (line.startsWith('SUMMARY:')) {
      event.summary = unescapeValue(line.slice(8));
    } else if (line.startsWith('DESCRIPTION:')) {
      event.description = unescapeValue(line.slice(12));
    } else if (line.startsWith('LOCATION:')) {
      event.location = unescapeValue(line.slice(9));
    } else if (line.startsWith('DTSTART')) {
      // Parse DTSTART with possible parameters
      const colonIdx = line.indexOf(':');
      const params = line.slice(7, colonIdx);
      const value = line.slice(colonIdx + 1);

      // Check for TZID
      const tzidMatch = params.match(/TZID=([^;:]+)/);
      startTzid = tzidMatch ? tzidMatch[1] : null;

      // Check for VALUE=DATE (all-day)
      const isAllDay = params.includes('VALUE=DATE');

      const parsed = parseICALDate(value, startTzid);
      if (parsed) {
        event.startDate = parsed.date;
        event.isAllDay = isAllDay || parsed.isAllDay;
      }
    } else if (line.startsWith('DTEND')) {
      const colonIdx = line.indexOf(':');
      const params = line.slice(5, colonIdx);
      const value = line.slice(colonIdx + 1);

      const tzidMatch = params.match(/TZID=([^;:]+)/);
      const tzid = tzidMatch ? tzidMatch[1] : startTzid;

      const parsed = parseICALDate(value, tzid);
      if (parsed) {
        event.endDate = parsed.date;
      }
    } else if (line.startsWith('RRULE:')) {
      event.rrule = line.slice(6);
    }
  }

  return event;
}

/**
 * Parse entire ICAL file
 */
function parseICAL(text) {
  const unfolded = unfoldLines(text);
  const lines = unfolded.split('\n');

  const events = [];
  let inEvent = false;
  let eventLines = [];

  for (const line of lines) {
    if (line === 'BEGIN:VEVENT') {
      inEvent = true;
      eventLines = [];
    } else if (line === 'END:VEVENT') {
      inEvent = false;
      const event = parseEvent(eventLines);

      // Expand recurring events
      if (event.rrule) {
        const expanded = expandRecurring(event, event.rrule);
        events.push(...expanded);
      } else {
        events.push(event);
      }
    } else if (inEvent) {
      eventLines.push(line);
    }
  }

  return events;
}

/**
 * Escape CSV field (quote if contains comma, newline, or quote)
 */
function escapeCSV(value) {
  if (!value) return '';
  const str = String(value);
  if (str.includes(',') || str.includes('\n') || str.includes('"')) {
    return '"' + str.replace(/"/g, '""') + '"';
  }
  return str;
}

/**
 * Convert events to CSV
 */
function eventsToCSV(events) {
  const headers = ['Début', 'Fin', 'Service', 'Évènement', 'Lieu', 'Description', 'UID'];
  const rows = [headers.join(',')];

  for (const event of events) {
    const { dept, cleanDesc } = extractDepartment(event.description);

    const row = [
      escapeCSV(formatDateForCSV(event.startDate, event.isAllDay)),
      escapeCSV(formatDateForCSV(event.endDate, event.isAllDay)),
      escapeCSV(dept),
      escapeCSV(event.summary),
      escapeCSV(event.location),
      escapeCSV(cleanDesc),
      escapeCSV(event.uid)
    ];

    rows.push(row.join(','));
  }

  return rows.join('\n');
}

/**
 * Main function
 */
function main() {
  const { inputFile, outputFile } = parseArgs();

  console.log(`Reading: ${inputFile}`);

  if (!fs.existsSync(inputFile)) {
    console.error(`Error: File not found: ${inputFile}`);
    process.exit(1);
  }

  const icalText = fs.readFileSync(inputFile, 'utf-8');
  console.log(`Parsing ICAL...`);

  const events = parseICAL(icalText);
  console.log(`Found ${events.length} events (including expanded recurring)`);

  // Sort by start date
  events.sort((a, b) => {
    if (!a.startDate) return 1;
    if (!b.startDate) return -1;
    return a.startDate.getTime() - b.startDate.getTime();
  });

  const csv = eventsToCSV(events);

  fs.writeFileSync(outputFile, csv, 'utf-8');
  console.log(`Written: ${outputFile}`);

  // Print summary by department
  const deptCounts = {};
  for (const event of events) {
    const { dept } = extractDepartment(event.description);
    const key = dept || '(no tag)';
    deptCounts[key] = (deptCounts[key] || 0) + 1;
  }

  console.log('\nEvents by department:');
  Object.entries(deptCounts)
    .sort((a, b) => b[1] - a[1])
    .forEach(([dept, count]) => {
      console.log(`  ${dept}: ${count}`);
    });
}

main();
