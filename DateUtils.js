function parseWeekMondayDate(weekString) {
  if (!weekString) return null;
  const match = String(weekString).trim().match(/\((\d{1,2})\/(\d{1,2})\)/);
  if (!match) return null;

  const month = parseInt(match[1], 10) - 1;
  const day = parseInt(match[2], 10);
  const now = new Date();
  const parsedDate = new Date(now.getFullYear(), month, day, 0, 0, 0, 0);

  if (parsedDate > now && (parsedDate - now) > 180 * 24 * 60 * 60 * 1000) {
    parsedDate.setFullYear(now.getFullYear() - 1);
  }

  return parsedDate;
}

function getMostRecentMonday() {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const diffToMonday = today.getDay() === 0 ? -6 : 1 - today.getDay();
  const monday = new Date(today);
  monday.setDate(today.getDate() + diffToMonday);
  return monday;
}

function getWeekDates(weekStart) {
  return Array.from({ length: 7 }, (_, dayOffset) => {
    const date = new Date(weekStart);
    date.setDate(weekStart.getDate() + dayOffset);
    return date;
  });
}

function getWeekDate(value) {
  if (value instanceof Date) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate(), 0, 0, 0, 0);
  }

  const text = String(value || '').trim();
  return text.toUpperCase().startsWith('WEEK ') ? parseWeekMondayDate(text) : null;
}

function areSameDate(left, right) {
  return left && right && left.getTime() === right.getTime();
}

function getStartOfDay(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate(), 0, 0, 0, 0);
}

function getDayOffset(date, weekStart) {
  const millisecondsPerDay = 24 * 60 * 60 * 1000;
  const normalizedDate = getStartOfDay(date);
  const normalizedWeekStart = getStartOfDay(weekStart);
  return Math.round((normalizedDate.getTime() - normalizedWeekStart.getTime()) / millisecondsPerDay);
}

function parseActivityDate(activity) {
  const [year, month, day] = activity.start_date_local.split('T')[0].split('-').map(Number);
  return new Date(year, month - 1, day, 0, 0, 0, 0);
}

function isLocalDevelopmentEnvironment() {
  return typeof PropertiesService === 'undefined';
}

function normalizeSheetDateColumn(sheet, sheetData, startRowIndex = 1) {
  for (let rowIndex = startRowIndex; rowIndex < sheetData.length; rowIndex++) {
    const parsedDate = parseSheetDateValue(sheetData[rowIndex]?.[0]);
    if (!parsedDate) continue;

    sheet.getRange(rowIndex + 1, 1)
      .setNumberFormat('@')
      .setValue(formatSheetDateText(parsedDate));
  }
}

function formatSheetDateText(date) {
  return date.toLocaleDateString('en-US', { month: 'long', day: 'numeric' });
}

// Sheet dates are stored as plain "Month Day" text, so the year comes from the current year.
function parseSheetDateText(value) {
  if (value instanceof Date) return getStartOfDay(value);

  const text = String(value || '').trim();
  if (!text) return null;

  const monthDayMatch = text.match(/^([A-Za-z]+)\s+(\d{1,2})$/);
  const parsedDate = monthDayMatch
    ? new Date(`${monthDayMatch[1]} ${monthDayMatch[2]}, ${new Date().getFullYear()}`)
    : new Date(text);
  return Number.isNaN(parsedDate.getTime()) ? null : getStartOfDay(parsedDate);
}

function parseSheetDateValue(value) {
  if (value instanceof Date) return getStartOfDay(value);
  const text = String(value || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}T/.test(text)) return null;

  const parsedDate = new Date(text);
  return Number.isNaN(parsedDate.getTime()) ? null : getStartOfDay(parsedDate);
}