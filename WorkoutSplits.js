const WORKOUT_SPLITS_SHEET_NAME = 'Workout Splits';
const WORKOUT_SPLITS_HEADERS = ['Date', 'Workout', 'Splits', 'Splits (Pace)'];

function parseSplitSeconds(value) {
  const match = String(value || '').trim().match(/(\d+):([0-5]\d)(?:\.(\d+))?\s*$/);
  if (!match) return null;

  return Number(match[1]) * 60 + Number(match[2]) + Number(`0.${match[3] || 0}`);
}

function parseWorkoutDistanceMiles(workout) {
  const matches = [...String(workout || '').matchAll(/(\d+(?:\.\d+)?)\s*(km|kilometers?|m|meters?|mi|miles?)(?:\b|$)/gi)];
  if (matches.length === 0) return null;

  const match = matches[matches.length - 1];
  const distance = Number(match[1]);
  const unit = match[2].toLowerCase();
  if (unit === 'km' || unit.startsWith('kilometer')) return distance * 0.621371;
  if (unit === 'm' || unit.startsWith('meter')) return distance / 1609.344;
  return distance;
}

function formatSplitPace(seconds) {
  if (seconds === null || !isFinite(seconds)) return '';
  const roundedSeconds = Math.round(seconds);
  return `${Math.floor(roundedSeconds / 60)}:${String(roundedSeconds % 60).padStart(2, '0')}`;
}

function parseSplitLine(line, distanceMiles) {
  const durations = String(line).match(/\d+:[0-5]\d(?:\.\d+)?/g) || [];
  if (durations.length === 0) return { value: line, pace: '' };

  const split = durations[0];
  const splitSeconds = parseSplitSeconds(split);
  const explicitPace = durations[1] || '';
  return {
    value: split,
    pace: explicitPace || (distanceMiles && splitSeconds !== null
      ? formatSplitPace(splitSeconds / distanceMiles)
      : '')
  };
}

function parseWorkoutSplits(description, workout) {
  if (!workout) return [];
  const match = String(description || '').match(/splits:\s*\r?\n([\s\S]*)/i);
  if (!match) return [];

  const distanceMiles = parseWorkoutDistanceMiles(workout);
  const splits = [];
  for (const line of match[1].split(/\r?\n/)) {
    const split = line.trim();
    if (!split) break;

    splits.push(parseSplitLine(split, distanceMiles));
  }
  return splits;
}

function ensureWorkoutSplitsHeader(sheet, sheetData) {
  if (sheetData.length === 0 || !sheetData[0].some(value => String(value || '').trim())) {
    sheet.getRange(1, 1, 1, WORKOUT_SPLITS_HEADERS.length).setValues([WORKOUT_SPLITS_HEADERS]);
    return [WORKOUT_SPLITS_HEADERS];
  }
  return sheetData;
}

function writeWorkoutSplitRows(sheet, sheetData, activityDate, workout, splits) {
  const dateKey = getDateKey(activityDate);
  const firstRowIndex = sheetData.findIndex((row, index) => {
    if (index === 0) return false;
    const rowDate = parseSheetDate(row[0]);
    return rowDate && getDateKey(rowDate) === dateKey && String(row[1] || '').trim() === workout;
  });
  const startRowIndex = firstRowIndex === -1 ? sheetData.length : firstRowIndex;
  const activityRange = sheet.getRange(startRowIndex + 1, 1, splits.length, 2);
  if (activityRange.breakApart) activityRange.breakApart();

  splits.forEach((split, splitIndex) => {
    const rowIndex = startRowIndex + splitIndex;
    sheet.getRange(rowIndex + 1, 1, 1, WORKOUT_SPLITS_HEADERS.length).setValues([[
      splitIndex === 0 ? activityDate : '',
      splitIndex === 0 ? workout : '',
      split.value,
      split.pace
    ]]);
    if (splitIndex === 0) sheet.getRange(rowIndex + 1, 1).setNumberFormat('mmmm d, yyyy');
  });

  for (let columnIndex = 1; columnIndex <= 2; columnIndex++) {
    sheet.getRange(startRowIndex + 1, columnIndex, splits.length, 1)
      .merge()
      .setVerticalAlignment('top');
  }
}

function updateWorkoutSplitsSheet(spreadsheet, activities) {
  const sheet = spreadsheet.getSheetByName(WORKOUT_SPLITS_SHEET_NAME)
    || spreadsheet.insertSheet(WORKOUT_SPLITS_SHEET_NAME);
  let sheetData = ensureWorkoutSplitsHeader(sheet, sheet.getDataRange().getValues());

  activities.filter(isRunningActivity).forEach(activity => {
    const workout = parseWorkout(activity.description);
    const splits = parseWorkoutSplits(activity.description, workout);
    if (!workout || splits.length === 0) return;

    const activityDate = parseActivityDate(activity);
    writeWorkoutSplitRows(sheet, sheetData, activityDate, workout, splits);
    sheetData = sheet.getDataRange().getValues();
  });

  sheet.getRange(1, 1, Math.max(sheetData.length, 1), WORKOUT_SPLITS_HEADERS.length)
    .setWrap(true)
    .setVerticalAlignment('middle');
  sheet.getRange(1, 1).setFontWeight('bold');
  sheet.setColumnWidth(1, 140);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 130);
}
