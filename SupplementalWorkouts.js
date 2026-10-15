const SUPPLEMENTAL_WORKOUTS_SHEET_NAME = 'Supplemental Workouts';
const SUPPLEMENTAL_WORKOUTS_HEADERS = ['Date', 'Workout', 'Exercise', 'Sets', 'Reps/Hold Time', 'Weight', 'Notes'];
const SUPPLEMENTAL_WORKOUTS_SEPARATOR_BACKGROUND = '#000000';
const SUPPLEMENTAL_WORKOUTS_SEPARATOR_HEIGHT = 10;

function isLocalSupplementalWorkoutsEnvironment() {
  return typeof PropertiesService === 'undefined';
}

function ensureSupplementalWorkoutsHeader(sheet, sheetData) {
  if (sheetData.length === 0 || !sheetData[0].some(value => String(value || '').trim())) {
    sheet.getRange(1, 1, 1, SUPPLEMENTAL_WORKOUTS_HEADERS.length).setValues([SUPPLEMENTAL_WORKOUTS_HEADERS]);
    return [SUPPLEMENTAL_WORKOUTS_HEADERS];
  }
  return sheetData;
}

function groupExercisesByCategory(exercises) {
  const groups = [];

  exercises.forEach(exercise => {
    let group = groups.find(existingGroup => existingGroup.category === exercise.category);
    if (!group) {
      group = { category: exercise.category, exercises: [] };
      groups.push(group);
    }
    group.exercises.push(exercise);
  });

  return groups;
}

function buildDaySignature(dateKey, groups) {
  const groupSignatures = groups.map(group => {
    const exerciseSignatures = group.exercises
      .map(exercise => [exercise.workout, exercise.sets, exercise.repsHoldTime, exercise.weight, exercise.notes].join('|'))
      .join(';');
    return `${group.category}:${exerciseSignatures}`;
  });
  return `${dateKey}|${groupSignatures.join('|')}`;
}

function isBlankRow(row) {
  return !row || row.every(value => String(value || '').trim() === '');
}

// Manually-entered rest-day blocks opt out of all script rewrites via this note.
function isNonRunningDayNote(value) {
  return /non[\s\-_]*running[\s\-_]*day/i.test(String(value || ''));
}

function isNonRunningDayBlock(sheetData, block) {
  for (let rowOffset = 0; rowOffset < block.rowCount; rowOffset++) {
    if (isNonRunningDayNote(sheetData[block.startRowIndex + rowOffset]?.[6])) return true;
  }
  return false;
}

function getBlockGroups(sheetData, block) {
  const groups = [];
  let currentGroup = null;

  for (let rowOffset = 0; rowOffset < block.rowCount; rowOffset++) {
    const row = sheetData[block.startRowIndex + rowOffset];
    const category = String(row[1] || '').trim();
    if (category) {
      currentGroup = { category, exercises: [] };
      groups.push(currentGroup);
    }

    if (currentGroup) {
      currentGroup.exercises.push({
        workout: row[2] || '',
        sets: row[3] || '',
        repsHoldTime: row[4] || '',
        weight: row[5] || '',
        notes: row[6] || ''
      });
    }
  }

  return groups;
}

function getAllDayBlocks(sheetData) {
  const blocks = [];
  let blockStartIndex = null;

  for (let rowIndex = 1; rowIndex <= sheetData.length; rowIndex++) {
    const row = sheetData[rowIndex];
    const isBoundary = rowIndex === sheetData.length || isBlankRow(row);

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

function getBlockDateKey(sheetData, block) {
  const blockDate = parseSheetDate(sheetData[block.startRowIndex]?.[0]);
  return blockDate ? getDateKey(blockDate) : null;
}

function findDayBlock(sheetData, dateKey) {
  return getAllDayBlocks(sheetData).find(block => getBlockDateKey(sheetData, block) === dateKey) || null;
}

function removeSupplementalWorkoutsBefore(sheet, sheetData, cutoffDate) {
  const cutoffKey = getDateKey(cutoffDate);
  const blocksToRemove = getAllDayBlocks(sheetData)
    .filter(block => {
      if (isNonRunningDayBlock(sheetData, block)) return false;
      const blockDateKey = getBlockDateKey(sheetData, block);
      return blockDateKey && blockDateKey < cutoffKey;
    })
    .sort((left, right) => right.startRowIndex - left.startRowIndex);

  blocksToRemove.forEach(block => {
    sheet.deleteRows(block.startRowIndex + 1, block.rowCount + 1);
  });
}

function findChronologicalInsertionRowIndex(sheetData, dateKey) {
  const nextBlock = getAllDayBlocks(sheetData).find(block => {
    const blockDateKey = getBlockDateKey(sheetData, block);
    return blockDateKey && blockDateKey > dateKey;
  });
  return nextBlock ? nextBlock.startRowIndex : null;
}

// A day with no Strava run shows as "Rest" (or is still blank) in Actual Runs. Supplemental rows
// existing for such a day were therefore entered by hand, which is the only way to tell a manual
// entry apart from a generated one — the sheet itself carries no manual/automatic marker.
function getManualEntryDateKeys(actualSheet, rowIndex, targetMonday, todayOffset) {
  const manualDateKeys = {};
  const weekDates = getWeekDates(targetMonday);

  for (let dayIndex = 0; dayIndex <= todayOffset; dayIndex++) {
    const cellValue = actualSheet.getRange(rowIndex + 1, dayIndex + 2).getValue();
    const hasNoRun = isRestValue(cellValue) || splitCellLines(cellValue).length === 0;
    if (hasNoRun) manualDateKeys[getDateKey(weekDates[dayIndex])] = true;
  }

  return manualDateKeys;
}

// Re-sorting must never reshape rows it did not generate, so a relocated block is carried over
// verbatim (all seven columns) instead of being rebuilt from its parsed groups.
function getBlockRawRows(sheetData, block) {
  const rawRows = [];
  for (let rowOffset = 0; rowOffset < block.rowCount; rowOffset++) {
    const row = sheetData[block.startRowIndex + rowOffset] || [];
    rawRows.push(SUPPLEMENTAL_WORKOUTS_HEADERS.map((_, columnIndex) => (
      row[columnIndex] === null || row[columnIndex] === undefined ? '' : row[columnIndex]
    )));
  }
  return rawRows;
}

function writeRawBlockRows(sheet, startRowIndex, rawRows) {
  rawRows.forEach((rawRow, rowOffset) => {
    const rowNumber = startRowIndex + rowOffset + 1;
    // Only the Date cell is normalized to the sheet's "Month Day" text; every other column moves
    // across exactly as the user left it.
    const blockDate = parseSheetDate(rawRow[0]);
    const values = rawRow.slice();
    if (blockDate) values[0] = formatSheetDateText(blockDate);

    if (String(values[0] || '').trim()) {
      sheet.getRange(rowNumber, 1).setNumberFormat('@');
    }
    sheet.getRange(rowNumber, 1, 1, SUPPLEMENTAL_WORKOUTS_HEADERS.length).setValues([values]);
    if (String(values[0] || '').trim()) {
      sheet.getRange(rowNumber, 1).setFontWeight('bold');
    }
  });

  // Each category span (a row with a Workout value plus the blank continuation rows under it)
  // regains the merged Date/Workout cells the block had before it moved.
  let spanStart = 0;
  for (let rowOffset = 1; rowOffset <= rawRows.length; rowOffset++) {
    const startsNewSpan = rowOffset === rawRows.length || String(rawRows[rowOffset][1] || '').trim() !== '';
    if (!startsNewSpan) continue;

    const spanLength = rowOffset - spanStart;
    if (spanLength > 1) {
      for (let columnIndex = 1; columnIndex <= 2; columnIndex++) {
        sheet.getRange(startRowIndex + spanStart + 1, columnIndex, spanLength, 1)
          .merge()
          .setVerticalAlignment('top');
      }
    }
    spanStart = rowOffset;
  }
}

// Moves only the blocks that sit in the wrong place, and only within the window the script is
// allowed to touch. Blocks already in order — and the entire log before the window — keep their
// existing rows, formatting, and merges untouched.
function sortDayBlocksIfNeeded(sheet, sheetData, earliestEditableDateKey = null) {
  const movedDateKeys = {};

  for (let pass = 0; pass < 100; pass++) {
    const blocks = getAllDayBlocks(sheetData);
    const blockDateKeys = blocks.map(block => getBlockDateKey(sheetData, block));

    // A block whose date cannot be read has no defined position, so ordering is left alone.
    if (blockDateKeys.some(dateKey => !dateKey)) return;

    const misplacedIndex = blockDateKeys.findIndex((dateKey, index) => {
      if (index === 0 || dateKey >= blockDateKeys[index - 1]) return false;
      if (earliestEditableDateKey && dateKey < earliestEditableDateKey) return false;
      return !movedDateKeys[dateKey];
    });
    if (misplacedIndex === -1) return;

    const block = blocks[misplacedIndex];
    const dateKey = blockDateKeys[misplacedIndex];
    const rawRows = getBlockRawRows(sheetData, block);
    movedDateKeys[dateKey] = true;

    sheet.deleteRows(block.startRowIndex + 1, block.rowCount + 1);
    sheetData = sheet.getDataRange().getValues();

    const insertionRowIndex = findChronologicalInsertionRowIndex(sheetData, dateKey);
    const targetRowIndex = insertionRowIndex === null ? findAppendRowIndex(sheet, sheetData) : insertionRowIndex;
    sheet.insertRows(targetRowIndex + 1, rawRows.length + 1);
    writeRawBlockRows(sheet, targetRowIndex, rawRows);
    writeSeparatorRow(sheet, targetRowIndex + rawRows.length);
    sheetData = sheet.getDataRange().getValues();
  }
}

function findAppendRowIndex(sheet, sheetData) {
  // A trailing separator can be omitted from getDataRange(); reuse that row when appending.
  return sheetData.length;
}

function writeSupplementalWorkoutGroup(sheet, startRowIndex, activityDate, group) {
  // Re-sorts pass the existing cell value through, which may still be a legacy Date/ISO value.
  const parsedDate = activityDate instanceof Date ? activityDate : parseSheetDateText(activityDate);
  const dateText = parsedDate ? formatSheetDateText(parsedDate) : activityDate;

  group.exercises.forEach((exercise, exerciseIndex) => {
    const rowIndex = startRowIndex + exerciseIndex;
    if (exerciseIndex === 0) sheet.getRange(rowIndex + 1, 1).setNumberFormat('@');
    sheet.getRange(rowIndex + 1, 1, 1, SUPPLEMENTAL_WORKOUTS_HEADERS.length).setValues([[
      exerciseIndex === 0 ? dateText : '',
      exerciseIndex === 0 ? group.category : '',
      exercise.workout,
      exercise.sets,
      exercise.repsHoldTime,
      exercise.weight,
      exercise.notes
    ]]);
    if (exerciseIndex === 0) {
      sheet.getRange(rowIndex + 1, 1).setFontWeight('bold');
    }
  });

  for (let columnIndex = 1; columnIndex <= 2; columnIndex++) {
    sheet.getRange(startRowIndex + 1, columnIndex, group.exercises.length, 1)
      .merge()
      .setVerticalAlignment('top');
  }

  return startRowIndex + group.exercises.length;
}

function writeSeparatorRow(sheet, rowIndex) {
  sheet.getRange(rowIndex + 1, 1, 1, SUPPLEMENTAL_WORKOUTS_HEADERS.length)
    .setBackground(SUPPLEMENTAL_WORKOUTS_SEPARATOR_BACKGROUND);
  sheet.setRowHeight(rowIndex + 1, SUPPLEMENTAL_WORKOUTS_SEPARATOR_HEIGHT);
}

// Row existence/content (not paint color) decides whether a row needs inserting: a blank row that
// lost its black background (e.g. from an external style wipe) is still a valid separator slot.
function ensureHeaderSeparator(sheet, sheetData) {
  const separatorRowIndex = 1;
  const rowIsMissingOrOccupied = sheetData.length <= separatorRowIndex || !isBlankRow(sheetData[separatorRowIndex]);
  if (rowIsMissingOrOccupied) {
    sheet.insertRows(separatorRowIndex + 1, 1);
  }

  writeSeparatorRow(sheet, separatorRowIndex);
  return sheet.getDataRange().getValues();
}

function fillMissingSeparatorColumns(sheet, sheetData) {
  const repaintIfSeparator = rowIndex => {
    const sheetRow = rowIndex + 1;
    const hasBlackSeparatorCell = [1, 2, 3, 4].some(column =>
      sheet.getRange(sheetRow, column).getBackground() === SUPPLEMENTAL_WORKOUTS_SEPARATOR_BACKGROUND
    );
    if (hasBlackSeparatorCell) {
      sheet.getRange(sheetRow, 5, 1, 3)
        .setBackground(SUPPLEMENTAL_WORKOUTS_SEPARATOR_BACKGROUND);
    }
  };

  for (let rowIndex = 1; rowIndex < sheetData.length; rowIndex++) {
    repaintIfSeparator(rowIndex);
  }

  // A trailing formatting-only separator may be omitted from getDataRange().
  repaintIfSeparator(sheetData.length);
}

function ensureDaySeparatorRows(sheet, sheetData) {
  getAllDayBlocks(sheetData).forEach(block => {
    const separatorRowIndex = block.startRowIndex + block.rowCount;
    const separatorAlreadyExists = sheet.getRange(separatorRowIndex + 1, 1).getBackground()
      === SUPPLEMENTAL_WORKOUTS_SEPARATOR_BACKGROUND;
    if (separatorRowIndex >= sheetData.length && !separatorAlreadyExists && sheet.insertRows) {
      sheet.insertRows(separatorRowIndex + 1, 1);
    }
    writeSeparatorRow(sheet, separatorRowIndex);
  });
}

// Collapses already-duplicated blank rows left over from the ensureHeaderSeparator/
// ensureFinalSeparatorRow bug above, keeping just one blank row per gap.
function removeConsecutiveBlankRows(sheet, sheetData) {
  for (let rowIndex = sheetData.length - 1; rowIndex >= 2; rowIndex--) {
    if (isBlankRow(sheetData[rowIndex]) && isBlankRow(sheetData[rowIndex - 1])) {
      sheet.deleteRows(rowIndex + 1, 1);
    }
  }
}

// Undoes stray black rows left over from prior bugs (e.g. a blank row before the first day block).
function clearUnexpectedSeparatorRows(sheet, sheetData) {
  const expectedSeparatorRowIndexes = new Set([1, sheetData.length]);
  getAllDayBlocks(sheetData).forEach(block => {
    expectedSeparatorRowIndexes.add(block.startRowIndex + block.rowCount);
  });

  for (let rowIndex = 1; rowIndex < sheetData.length; rowIndex++) {
    if (!isBlankRow(sheetData[rowIndex]) || expectedSeparatorRowIndexes.has(rowIndex)) continue;
    const isPainted = sheet.getRange(rowIndex + 1, 1).getBackground() === SUPPLEMENTAL_WORKOUTS_SEPARATOR_BACKGROUND;
    if (!isPainted) continue;
    sheet.getRange(rowIndex + 1, 1, 1, SUPPLEMENTAL_WORKOUTS_HEADERS.length).setBackground(null);
  }
}

function ensureFinalSeparatorRow(sheet) {
  if (!sheet.getLastRow) return;

  const lastRow = sheet.getLastRow();
  const separatorRow = lastRow + 1;
  const currentValues = sheet.getDataRange().getValues();
  const rowHasContent = currentValues.length >= separatorRow && !isBlankRow(currentValues[separatorRow - 1]);
  const localSeparatorIsMissing = isLocalSupplementalWorkoutsEnvironment()
    && currentValues.length < separatorRow;
  if ((rowHasContent || localSeparatorIsMissing) && sheet.insertRows) {
    sheet.insertRows(separatorRow, 1);
  }

  writeSeparatorRow(sheet, separatorRow - 1);
}

function forceKnownSeparatorRow(sheet, rowNumber) {
  const separatorRange = sheet.getRange(rowNumber, 1, 1, SUPPLEMENTAL_WORKOUTS_HEADERS.length);
  const hasContent = Array.from({ length: SUPPLEMENTAL_WORKOUTS_HEADERS.length }, (_, columnIndex) =>
    String(sheet.getRange(rowNumber, columnIndex + 1).getValue?.() || '').trim()
  ).some(Boolean);
  if (hasContent) return;

  if (separatorRange.breakApart) separatorRange.breakApart();
  if (separatorRange.clearContent) separatorRange.clearContent();
  separatorRange.setBackground(SUPPLEMENTAL_WORKOUTS_SEPARATOR_BACKGROUND);
  sheet.setRowHeight(rowNumber, SUPPLEMENTAL_WORKOUTS_SEPARATOR_HEIGHT);
}

function clearBlackFromWorkoutRows(sheet, sheetData) {
  sheetData.forEach((row, rowIndex) => {
    if (rowIndex === 0 || isBlankRow(row)) return;
    sheet.autoResizeRows?.(rowIndex + 1, 1);
    sheet.getRange(rowIndex + 1, 1, 1, SUPPLEMENTAL_WORKOUTS_HEADERS.length)
      .setBackground(null);
  });
}

function matchSupplementalFontToRowThree(sheet, sheetData) {
  if (sheetData.length < 3) return;

  const referenceCell = sheet.getRange(3, 1);
  const fontFamily = referenceCell.getFontFamily?.();
  const fontSize = referenceCell.getFontSize?.();
  const range = sheet.getRange(1, 1, sheetData.length, SUPPLEMENTAL_WORKOUTS_HEADERS.length);
  if (fontFamily !== undefined) range.setFontFamily(fontFamily);
  if (fontSize !== undefined) range.setFontSize(fontSize);
}

// Re-paints separators/fonts for the local view; shared by the main sync and by tools that pull
// raw values from production (which carry no formatting of their own).
function repaintLocalSupplementalWorkoutsSeparators(sheet, sheetData) {
  removeConsecutiveBlankRows(sheet, sheetData);
  sheetData = sheet.getDataRange().getValues();
  ensureDaySeparatorRows(sheet, sheetData);
  clearUnexpectedSeparatorRows(sheet, sheetData);
  clearBlackFromWorkoutRows(sheet, sheetData);
  matchSupplementalFontToRowThree(sheet, sheetData);
  return sheet.getDataRange().getValues();
}

function updateSupplementalWorkoutsSheet(spreadsheet, activities, options = {}) {
  const manualDateKeys = options.manualDateKeys || {};
  const earliestEditableDateKey = options.earliestEditableDateKey || null;
  const sheet = spreadsheet.getSheetByName(SUPPLEMENTAL_WORKOUTS_SHEET_NAME)
    || spreadsheet.insertSheet(SUPPLEMENTAL_WORKOUTS_SHEET_NAME);
  let sheetData = ensureSupplementalWorkoutsHeader(sheet, sheet.getDataRange().getValues());
  sheetData = ensureHeaderSeparator(sheet, sheetData);
  normalizeSheetDateColumn(sheet, sheet.getDataRange().getValues());
  if (isLocalSupplementalWorkoutsEnvironment()) {
    removeSupplementalWorkoutsBefore(sheet, sheet.getDataRange().getValues(), new Date(new Date().getFullYear(), 8, 15));
    sheetData = sheet.getDataRange().getValues();
  }

  activities.forEach(activity => {
    const exercises = parseSupplementalWorkoutDetails(activity.description).exercises;
    if (exercises.length === 0) return;

    const activityDate = parseActivityDate(activity);
    const dateKey = getDateKey(activityDate);

    // Days before the sync window belong to the historical log and are never rewritten.
    if (earliestEditableDateKey && dateKey < earliestEditableDateKey) return;

    const groups = groupExercisesByCategory(exercises);
    const daySignature = buildDaySignature(dateKey, groups);

    const existingBlock = findDayBlock(sheetData, dateKey);
    if (existingBlock) {
      if (isNonRunningDayBlock(sheetData, existingBlock)) return;
      // No run that day means these rows were typed in by hand, so leave them alone.
      if (manualDateKeys[dateKey]) return;

      const existingSignature = buildDaySignature(dateKey, getBlockGroups(sheetData, existingBlock));
      if (existingSignature === daySignature) return;

      // Overwrite: drop the day's old rows (plus its trailing separator) before re-inserting fresh ones.
      sheet.deleteRows(existingBlock.startRowIndex + 1, existingBlock.rowCount + 1);
      sheetData = sheet.getDataRange().getValues();
    }

    const rowsNeeded = groups.reduce((sum, group) => sum + group.exercises.length, 0) + 1;
    const insertionRowIndex = findChronologicalInsertionRowIndex(sheetData, dateKey);

    let rowIndex;
    if (insertionRowIndex === null) {
      rowIndex = findAppendRowIndex(sheet, sheetData);
    } else {
      sheet.insertRows(insertionRowIndex + 1, rowsNeeded);
      rowIndex = insertionRowIndex;
    }

    groups.forEach(group => {
      rowIndex = writeSupplementalWorkoutGroup(sheet, rowIndex, activityDate, group);
    });
    writeSeparatorRow(sheet, rowIndex);

    sheetData = sheet.getDataRange().getValues();
  });

  sortDayBlocksIfNeeded(sheet, sheetData, earliestEditableDateKey);
  sheetData = sheet.getDataRange().getValues();
  if (isLocalSupplementalWorkoutsEnvironment()) {
    sheetData = repaintLocalSupplementalWorkoutsSeparators(sheet, sheetData);
  } else {
    ensureDaySeparatorRows(sheet, sheetData);
    ensureFinalSeparatorRow(sheet);
    sheetData = sheet.getDataRange().getValues();
    fillMissingSeparatorColumns(sheet, sheetData);
    forceKnownSeparatorRow(sheet, 12);
  }

  const contentRange = sheet.getRange(1, 1, Math.max(sheetData.length, 1), SUPPLEMENTAL_WORKOUTS_HEADERS.length)
    .setWrap(true)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('left');
  if (isLocalSupplementalWorkoutsEnvironment()) {
    contentRange.setFontFamily(null).setFontSize(null);
    for (let rowIndex = 1; rowIndex < sheetData.length; rowIndex++) {
      if (sheetData[rowIndex]?.[0]) {
        sheet.getRange(rowIndex + 1, 1).setNumberFormat('@');
        sheet.getRange(rowIndex + 1, 1).setFontWeight('bold');
      }
    }
  }
  if (sheetData.length > 1) {
    sheet.getRange(2, 1, sheetData.length - 1, 2)
      .setVerticalAlignment('top')
      .setHorizontalAlignment('left');
  }
  sheet.getRange(1, 1, 1, SUPPLEMENTAL_WORKOUTS_HEADERS.length).setFontWeight('bold');
  sheet.setColumnWidth(1, 150);
  sheet.setColumnWidth(2, 150);
  sheet.setColumnWidth(3, 200);
  sheet.setColumnWidth(4, 70);
  sheet.setColumnWidth(5, 140);
  sheet.setColumnWidth(6, 150);
  sheet.setColumnWidth(7, 220);
  if (isLocalSupplementalWorkoutsEnvironment()) ensureFinalSeparatorRow(sheet);
}

// Collects the Workout column categories logged for each day, keyed by date. Rows that only
// continue the day/category above them (blank Date/Workout cells, or cells hidden by a merge)
// inherit the most recent values seen above.
function getSupplementalWorkoutCategoriesByDate(sheetData) {
  const categoriesByDate = {};
  let currentDateKey = null;

  for (let rowIndex = 1; rowIndex < sheetData.length; rowIndex++) {
    const row = sheetData[rowIndex];
    if (isBlankRow(row)) continue;

    const rowDate = parseSheetDate(row[0]);
    if (rowDate) currentDateKey = getDateKey(rowDate);
    if (!currentDateKey) continue;

    const category = String(row[1] || '').replace(/\u00A0/g, ' ').trim();
    if (!category) continue;

    if (!categoriesByDate[currentDateKey]) categoriesByDate[currentDateKey] = [];
    if (!categoriesByDate[currentDateKey].some(existing => existing.toLowerCase() === category.toLowerCase())) {
      categoriesByDate[currentDateKey].push(category);
    }
  }

  return categoriesByDate;
}

function appendSupplementalWorkoutLines(existingValue, categories) {
  const originalValue = splitCellLines(existingValue).join('\n');
  const updatedValue = appendMissingCellLines(originalValue === '' ? 'Rest' : originalValue, categories);

  return updatedValue === originalValue ? null : updatedValue;
}

// Second pass over the week: anything logged on the Supplemental Workouts sheet (including rows
// entered by hand on rest days, which no Strava activity covers) is appended to its Actual Runs cell.
function applySupplementalWorkoutsToActualRuns(spreadsheet, actualSheet, rowIndex, targetMonday, todayOffset) {
  const sheet = spreadsheet.getSheetByName(SUPPLEMENTAL_WORKOUTS_SHEET_NAME);
  if (!sheet) return [];

  const categoriesByDate = getSupplementalWorkoutCategoriesByDate(sheet.getDataRange().getValues());
  const weekDates = getWeekDates(targetMonday);
  const updatedDayIndexes = [];

  for (let dayIndex = 0; dayIndex <= todayOffset; dayIndex++) {
    const categories = categoriesByDate[getDateKey(weekDates[dayIndex])];
    if (!categories || categories.length === 0) continue;

    const targetCell = actualSheet.getRange(rowIndex + 1, dayIndex + 2);
    const updatedValue = appendSupplementalWorkoutLines(targetCell.getValue(), categories);
    if (updatedValue === null) continue;

    targetCell.setValue(updatedValue);
    if (targetCell.setFontWeight) targetCell.setFontWeight('normal');
    updatedDayIndexes.push(dayIndex);
  }

  return updatedDayIndexes;
}

// Reapplies separator formatting to a sheet whose values came from elsewhere (e.g. pulled
// straight from production), without touching data or pruning old rows.
function repairSupplementalWorkoutsSeparators(sheet) {
  let sheetData = ensureSupplementalWorkoutsHeader(sheet, sheet.getDataRange().getValues());
  sheetData = ensureHeaderSeparator(sheet, sheetData);
  normalizeSheetDateColumn(sheet, sheet.getDataRange().getValues());
  sheetData = sheet.getDataRange().getValues();

  sortDayBlocksIfNeeded(sheet, sheetData);
  sheetData = sheet.getDataRange().getValues();
  repaintLocalSupplementalWorkoutsSeparators(sheet, sheetData);
  ensureFinalSeparatorRow(sheet);
}