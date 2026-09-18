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

function getExistingDaySignatures(sheetData) {
  const signatures = new Set();
  let dateKey = null;
  let groups = [];
  let currentGroup = null;

  const flushDay = () => {
    if (dateKey && groups.length > 0) signatures.add(buildDaySignature(dateKey, groups));
    dateKey = null;
    groups = [];
    currentGroup = null;
  };

  sheetData.slice(1).forEach(row => {
    if (isBlankRow(row)) {
      flushDay();
      return;
    }

    const rowDate = parseSheetDate(row[0]);
    if (rowDate) dateKey = getDateKey(rowDate);

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
  });
  flushDay();

  return signatures;
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

function updateSupplementalWorkoutsSheet(spreadsheet, activities) {
  const sheet = spreadsheet.getSheetByName(SUPPLEMENTAL_WORKOUTS_SHEET_NAME)
    || spreadsheet.insertSheet(SUPPLEMENTAL_WORKOUTS_SHEET_NAME);
  let sheetData = ensureSupplementalWorkoutsHeader(sheet, sheet.getDataRange().getValues());
  const existingDaySignatures = getExistingDaySignatures(sheetData);

  activities.forEach(activity => {
    const exercises = parseSupplementalWorkoutDetails(activity.description).exercises;
    if (exercises.length === 0) return;

    const activityDate = parseActivityDate(activity);
    const groups = groupExercisesByCategory(exercises);
    const daySignature = buildDaySignature(getDateKey(activityDate), groups);
    if (existingDaySignatures.has(daySignature)) return;

    let rowIndex = findAppendRowIndex(sheet, sheetData);
    groups.forEach(group => {
      rowIndex = writeSupplementalWorkoutGroup(sheet, rowIndex, activityDate, group);
    });
    writeSeparatorRow(sheet, rowIndex);

    existingDaySignatures.add(daySignature);
    sheetData = sheet.getDataRange().getValues();
  });

  sheet.getRange(1, 1, Math.max(sheetData.length, 1), SUPPLEMENTAL_WORKOUTS_HEADERS.length)
    .setWrap(true)
    .setVerticalAlignment('middle');
  if (sheetData.length > 1) {
    sheet.getRange(2, 1, sheetData.length - 1, 2)
      .setVerticalAlignment('top')
      .setHorizontalAlignment('left');
  }
  sheet.getRange(1, 1, 1, SUPPLEMENTAL_WORKOUTS_HEADERS.length).setFontWeight('bold');
  sheet.setColumnWidth(1, 120);
  sheet.setColumnWidth(2, 130);
  sheet.setColumnWidth(3, 170);
  sheet.setColumnWidth(4, 70);
  sheet.setColumnWidth(5, 140);
  sheet.setColumnWidth(6, 110);
  sheet.setColumnWidth(7, 220);
}