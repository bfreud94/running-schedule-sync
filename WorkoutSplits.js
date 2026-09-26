const WORKOUT_SPLITS_SHEET_NAME = 'Workout Splits';
const WORKOUT_SPLITS_HEADERS = ['Date', 'Workout', 'Splits', 'Splits (Pace)'];
const WORKOUT_SPLITS_SEPARATOR_BACKGROUND = '#000000';
const WORKOUT_SPLITS_SEPARATOR_HEIGHT = 10;

function isLocalWorkoutSplitsEnvironment() {
  return typeof PropertiesService === 'undefined';
}

function formatWorkoutSplitsDateText(date) {
  return date.toLocaleDateString('en-US', { month: 'long', day: 'numeric' });
}

// Plain-text "Month Day" cells carry no year, so matching must compare month/day only.
function parseWorkoutSplitsRowDate(value) {
  if (value instanceof Date) return value;
  const match = String(value || '').trim().match(/^([A-Za-z]+)\s+(\d{1,2})$/);
  if (!match) return null;

  const parsedDate = new Date(`${match[1]} ${match[2]}, 2000`);
  return Number.isNaN(parsedDate.getTime()) ? null : parsedDate;
}

function isSameWorkoutSplitsDay(left, right) {
  return left.getMonth() === right.getMonth() && left.getDate() === right.getDate();
}

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
  const calculatedPace = distanceMiles && splitSeconds !== null
    ? formatSplitPace(splitSeconds / distanceMiles)
    : '';
  return {
    value: split,
    pace: explicitPace || calculatedPace || split
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

function writeWorkoutSplitsSeparatorRow(sheet, rowIndex) {
  for (let column = 1; column <= WORKOUT_SPLITS_HEADERS.length; column++) {
    sheet.getRange(rowIndex + 1, column).setBackground(WORKOUT_SPLITS_SEPARATOR_BACKGROUND);
  }
  sheet.setRowHeight(rowIndex + 1, WORKOUT_SPLITS_SEPARATOR_HEIGHT);
}

function isWorkoutSplitsBlankRow(row) {
  return !row || row.every(value => String(value || '').trim() === '');
}

function getAllWorkoutSplitsBlocks(sheetData) {
  const blocks = [];
  let blockStartIndex = null;

  for (let rowIndex = 1; rowIndex <= sheetData.length; rowIndex++) {
    const row = sheetData[rowIndex];
    const isBoundary = rowIndex === sheetData.length || isWorkoutSplitsBlankRow(row);

    if (isBoundary) {
      if (blockStartIndex !== null) {
        blocks.push({ startRowIndex: blockStartIndex, rowCount: rowIndex - blockStartIndex });
        blockStartIndex = null;
      }
      continue;
    }

    if (blockStartIndex === null) blockStartIndex = rowIndex;
  }

  return blocks;
}

function removeDuplicateWorkoutSplitBlocks(sheet, sheetData) {
  const blocks = getAllWorkoutSplitsBlocks(sheetData);
  const lastBlockBySignature = new Map();
  const duplicateBlocks = [];

  blocks.forEach(block => {
    const anchorRow = sheetData[block.startRowIndex];
    const dateText = String(anchorRow?.[0] || '').trim();
    const workoutText = String(anchorRow?.[1] || '').trim();
    if (!dateText || !workoutText) return;

    const signature = `${dateText}|${workoutText}`;
    const previousBlock = lastBlockBySignature.get(signature);
    if (previousBlock) duplicateBlocks.push(previousBlock);
    lastBlockBySignature.set(signature, block);
  });

  duplicateBlocks
    .sort((left, right) => right.startRowIndex - left.startRowIndex)
    .forEach(block => {
      sheet.deleteRows(block.startRowIndex + 1, block.rowCount + 1);
    });

  return duplicateBlocks.length > 0;
}

function ensureWorkoutSplitSeparatorRows(sheet, sheetData) {
  const headerSeparatorIsMissing = !isWorkoutSplitsBlankRow(sheetData[1]);
  if (headerSeparatorIsMissing || sheetData.length < 2) sheet.insertRows(2, 1);
  writeWorkoutSplitsSeparatorRow(sheet, 1);
  sheetData = sheet.getDataRange().getValues();

  const dayStartRowIndexes = sheetData
    .map((row, rowIndex) => ({ rowIndex, date: parseSheetDate(row?.[0]) }))
    .filter(({ date }) => date)
    .map(({ rowIndex }) => rowIndex);

  dayStartRowIndexes.slice(1).reverse().forEach(rowIndex => {
    if (isWorkoutSplitsBlankRow(sheetData[rowIndex - 1])) {
      writeWorkoutSplitsSeparatorRow(sheet, rowIndex - 1);
      return;
    }

    sheet.insertRows(rowIndex + 1, 1);
    writeWorkoutSplitsSeparatorRow(sheet, rowIndex);
  });

  const lastRow = sheet.getLastRow();
  const finalSeparatorRow = lastRow + 1;
  if (sheet.getDataRange().getValues().length < finalSeparatorRow) {
    sheet.insertRows(finalSeparatorRow, 1);
  }
  writeWorkoutSplitsSeparatorRow(sheet, finalSeparatorRow - 1);

  return sheet.getDataRange().getValues();
}

function clearSeparatorFormattingFromWorkoutRow(sheet, rowIndex) {
  if (!isLocalWorkoutSplitsEnvironment()) return;

  const sheetRow = rowIndex + 1;
  sheet.getRange(sheetRow, 1, 1, WORKOUT_SPLITS_HEADERS.length).setBackground(null);

  let column = WORKOUT_SPLITS_HEADERS.length + 1;
  while (sheet.getRange(sheetRow, column).getBackground() === WORKOUT_SPLITS_SEPARATOR_BACKGROUND) {
    sheet.getRange(sheetRow, column).setBackground(null);
    column++;
  }
  sheet.autoResizeRows?.(sheetRow, 1);
}

function findAppendRowIndex(sheet, sheetData) {
  // Separator rows hold no values, so the data range can stop short of them.
  let rowIndex = sheetData.length;
  while (sheet.getRange(rowIndex + 1, 1).getBackground() === WORKOUT_SPLITS_SEPARATOR_BACKGROUND) {
    rowIndex++;
  }
  return rowIndex;
}

function writeWorkoutSplitRows(sheet, sheetData, activityDate, workout, splits) {
  const firstRowIndex = sheetData.findIndex((row, index) => {
    if (index === 0) return false;
    const rowDate = parseWorkoutSplitsRowDate(row[0]);
    return rowDate && isSameWorkoutSplitsDay(rowDate, activityDate) && String(row[1] || '').trim() === workout;
  });
  const startRowIndex = firstRowIndex === -1 ? findAppendRowIndex(sheet, sheetData) : firstRowIndex;
  const activityRange = sheet.getRange(startRowIndex + 1, 1, splits.length, 2);
  if (activityRange.breakApart) activityRange.breakApart();

  splits.forEach((split, splitIndex) => {
    const rowIndex = startRowIndex + splitIndex;
    clearSeparatorFormattingFromWorkoutRow(sheet, rowIndex);
    if (splitIndex === 0) sheet.getRange(rowIndex + 1, 1).setNumberFormat('@');
    sheet.getRange(rowIndex + 1, 1, 1, WORKOUT_SPLITS_HEADERS.length).setValues([[
      splitIndex === 0 ? formatWorkoutSplitsDateText(activityDate) : '',
      splitIndex === 0 ? workout : '',
      split.value,
      split.pace
    ]]);
    if (splitIndex === 0) {
      sheet.getRange(rowIndex + 1, 1).setFontWeight('bold');
    }
  });

  for (let columnIndex = 1; columnIndex <= 2; columnIndex++) {
    sheet.getRange(startRowIndex + 1, columnIndex, splits.length, 1)
      .merge()
      .setVerticalAlignment('top');
  }

  writeWorkoutSplitsSeparatorRow(sheet, startRowIndex + splits.length);
}

function updateWorkoutSplitsSheet(spreadsheet, activities) {
  const sheet = spreadsheet.getSheetByName(WORKOUT_SPLITS_SHEET_NAME)
    || spreadsheet.insertSheet(WORKOUT_SPLITS_SHEET_NAME);
  let sheetData = ensureWorkoutSplitsHeader(sheet, sheet.getDataRange().getValues());
  if (removeDuplicateWorkoutSplitBlocks(sheet, sheetData)) sheetData = sheet.getDataRange().getValues();

  activities.filter(isRunningActivity).forEach(activity => {
    const workout = parseWorkout(activity.description);
    const splits = parseWorkoutSplits(activity.description, workout);
    if (!workout || splits.length === 0) return;

    const activityDate = parseActivityDate(activity);
    writeWorkoutSplitRows(sheet, sheetData, activityDate, workout, splits);
    sheetData = sheet.getDataRange().getValues();
  });

  if (isLocalWorkoutSplitsEnvironment()) {
    sheetData = ensureWorkoutSplitSeparatorRows(sheet, sheetData);
  }

  const contentRange = sheet.getRange(1, 1, Math.max(sheetData.length, 1), WORKOUT_SPLITS_HEADERS.length)
    .setWrap(true)
    .setVerticalAlignment('middle');
  if (isLocalWorkoutSplitsEnvironment()) {
    contentRange.setHorizontalAlignment('left');
    for (let rowIndex = 1; rowIndex < sheetData.length; rowIndex++) {
      const cellValue = sheetData[rowIndex]?.[0];
      if (!cellValue) continue;

      // Legacy Date-typed cells (from before dates became plain text) get migrated on the next local run.
      const dateCell = sheet.getRange(rowIndex + 1, 1);
      dateCell.setNumberFormat('@');
      if (cellValue instanceof Date) dateCell.setValue(formatWorkoutSplitsDateText(cellValue));
      dateCell.setFontWeight('bold');
    }
  }
  if (sheetData.length > 1) {
    sheet.getRange(2, 1, sheetData.length - 1, 2)
      .setVerticalAlignment('top')
      .setHorizontalAlignment('left');
    if (isLocalWorkoutSplitsEnvironment()) {
      // Column 1 holds bolded dates; only body columns 2+ reset to normal weight.
      sheet.getRange(2, 2, sheetData.length - 1, WORKOUT_SPLITS_HEADERS.length - 1)
        .setFontWeight('normal');
    }
  }
  if (!isLocalWorkoutSplitsEnvironment()) {
    sheet.getRange(1, 3, Math.max(sheetData.length, 1), 2).setHorizontalAlignment('left');
  }
  sheet.getRange(1, 1, 1, WORKOUT_SPLITS_HEADERS.length).setFontWeight('bold');
  sheet.setColumnWidth(1, 140);
  sheet.setColumnWidth(2, 140);
  sheet.setColumnWidth(3, 100);
  sheet.setColumnWidth(4, 130);
}
