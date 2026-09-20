const SUPPLEMENTAL_WORKOUTS_SHEET_NAME = 'Supplemental Workouts';
const SUPPLEMENTAL_WORKOUTS_HEADERS = ['Date', 'Workout', 'Exercise', 'Sets', 'Reps/Hold Time', 'Weight', 'Notes'];
const SUPPLEMENTAL_WORKOUTS_SEPARATOR_BACKGROUND = '#000000';
const SUPPLEMENTAL_WORKOUTS_SEPARATOR_HEIGHT = 10;

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
  return row.every(value => String(value || '').trim() === '');
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
  const blockDate = parseSheetDate(sheetData[block.startRowIndex][0]);
  return blockDate ? getDateKey(blockDate) : null;
}

function findDayBlock(sheetData, dateKey) {
  return getAllDayBlocks(sheetData).find(block => getBlockDateKey(sheetData, block) === dateKey) || null;
}

function findChronologicalInsertionRowIndex(sheetData, dateKey) {
  const nextBlock = getAllDayBlocks(sheetData).find(block => {
    const blockDateKey = getBlockDateKey(sheetData, block);
    return blockDateKey && blockDateKey > dateKey;
  });
  return nextBlock ? nextBlock.startRowIndex : null;
}

function sortDayBlocksIfNeeded(sheet, sheetData) {
  const blocks = getAllDayBlocks(sheetData);
  const blockDateKeys = blocks.map(block => getBlockDateKey(sheetData, block));
  const isAlreadySorted = blockDateKeys.every((dateKey, index) => index === 0 || blockDateKeys[index - 1] <= dateKey);
  if (isAlreadySorted) return;

  const orderedBlocks = blocks
    .map(block => ({
      activityDate: sheetData[block.startRowIndex][0],
      dateKey: getBlockDateKey(sheetData, block),
      groups: getBlockGroups(sheetData, block)
    }))
    .sort((left, right) => (left.dateKey < right.dateKey ? -1 : left.dateKey > right.dateKey ? 1 : 0));

  const firstBlockRowIndex = Math.min(...blocks.map(block => block.startRowIndex));
  sheet.deleteRows(firstBlockRowIndex + 1, sheetData.length - firstBlockRowIndex);

  let rowIndex = firstBlockRowIndex;
  orderedBlocks.forEach(({ activityDate, groups }) => {
    groups.forEach(group => {
      rowIndex = writeSupplementalWorkoutGroup(sheet, rowIndex, activityDate, group);
    });
    writeSeparatorRow(sheet, rowIndex);
    rowIndex++;
  });
}

function findAppendRowIndex(sheet, sheetData) {
  // Separator rows hold no values, so the data range can stop short of them.
  let rowIndex = sheetData.length;
  while (sheet.getRange(rowIndex + 1, 1).getBackground() === SUPPLEMENTAL_WORKOUTS_SEPARATOR_BACKGROUND) {
    rowIndex++;
  }
  return rowIndex;
}

function writeSupplementalWorkoutGroup(sheet, startRowIndex, activityDate, group) {
  group.exercises.forEach((exercise, exerciseIndex) => {
    const rowIndex = startRowIndex + exerciseIndex;
    sheet.getRange(rowIndex + 1, 1, 1, SUPPLEMENTAL_WORKOUTS_HEADERS.length).setValues([[
      exerciseIndex === 0 ? activityDate : '',
      exerciseIndex === 0 ? group.category : '',
      exercise.workout,
      exercise.sets,
      exercise.repsHoldTime,
      exercise.weight,
      exercise.notes
    ]]);
    if (exerciseIndex === 0) sheet.getRange(rowIndex + 1, 1).setNumberFormat('mmmm d, yyyy');
  });

  for (let columnIndex = 1; columnIndex <= 2; columnIndex++) {
    sheet.getRange(startRowIndex + 1, columnIndex, group.exercises.length, 1)
      .merge()
      .setVerticalAlignment('top');
  }

  return startRowIndex + group.exercises.length;
}

function writeSeparatorRow(sheet, rowIndex) {
  for (let column = 1; column <= SUPPLEMENTAL_WORKOUTS_HEADERS.length; column++) {
    sheet.getRange(rowIndex + 1, column).setBackground(SUPPLEMENTAL_WORKOUTS_SEPARATOR_BACKGROUND);
  }
  sheet.setRowHeight(rowIndex + 1, SUPPLEMENTAL_WORKOUTS_SEPARATOR_HEIGHT);
}

function repairSeparatorRows(sheet, sheetData) {
  // Separator rows written before a column (e.g. Exercise, Weight) existed only got painted
  // up to the column count at that time; repaint every one across all current columns.
  // A stray merge from before a column existed can also leave a "blank" row reading as
  // non-blank, so also treat an already-black column A as a separator signal.
  sheetData.slice(1).forEach((row, index) => {
    const rowIndex = index + 1;
    const looksBlank = isBlankRow(row);
    const alreadyMarkedAsSeparator = sheet.getRange(rowIndex + 1, 1).getBackground() === SUPPLEMENTAL_WORKOUTS_SEPARATOR_BACKGROUND;
    if (!looksBlank && !alreadyMarkedAsSeparator) return;

    writeSeparatorRow(sheet, rowIndex);
  });
}

function updateSupplementalWorkoutsSheet(spreadsheet, activities) {
  const sheet = spreadsheet.getSheetByName(SUPPLEMENTAL_WORKOUTS_SHEET_NAME)
    || spreadsheet.insertSheet(SUPPLEMENTAL_WORKOUTS_SHEET_NAME);
  let sheetData = ensureSupplementalWorkoutsHeader(sheet, sheet.getDataRange().getValues());

  activities.forEach(activity => {
    const exercises = parseSupplementalWorkoutDetails(activity.description).exercises;
    if (exercises.length === 0) return;

    const activityDate = parseActivityDate(activity);
    const dateKey = getDateKey(activityDate);
    const groups = groupExercisesByCategory(exercises);
    const daySignature = buildDaySignature(dateKey, groups);

    const existingBlock = findDayBlock(sheetData, dateKey);
    if (existingBlock) {
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

  sortDayBlocksIfNeeded(sheet, sheetData);
  sheetData = sheet.getDataRange().getValues();
  repairSeparatorRows(sheet, sheetData);

  sheet.getRange(1, 1, Math.max(sheetData.length, 1), SUPPLEMENTAL_WORKOUTS_HEADERS.length)
    .setWrap(true)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('left');
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
}