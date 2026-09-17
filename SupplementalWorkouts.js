const SUPPLEMENTAL_WORKOUTS_SHEET_NAME = 'Supplemental Workouts';
const SUPPLEMENTAL_WORKOUTS_HEADERS = ['Date', 'Workout', 'Sets', 'Reps/Hold Time', 'Notes'];

function ensureSupplementalWorkoutsHeader(sheet, sheetData) {
  if (sheetData.length === 0 || !sheetData[0].some(value => String(value || '').trim())) {
    sheet.getRange(1, 1, 1, SUPPLEMENTAL_WORKOUTS_HEADERS.length).setValues([SUPPLEMENTAL_WORKOUTS_HEADERS]);
    return [SUPPLEMENTAL_WORKOUTS_HEADERS];
  }
  return sheetData;
}

function getSupplementalWorkoutRowKey(row) {
  const rowDate = parseSheetDate(row[0]);
  const dateKey = rowDate ? getDateKey(rowDate) : String(row[0] || '').trim();
  return [
    dateKey,
    String(row[1] || '').trim(),
    String(row[2] || '').trim(),
    String(row[3] || '').trim(),
    String(row[4] || '').trim()
  ].join('\u0001');
}

function appendSupplementalWorkoutRows(sheet, sheetData, rows) {
  const existingRows = new Set(sheetData.slice(1).map(getSupplementalWorkoutRowKey));
  let nextRowIndex = sheetData.length;

  rows.forEach(row => {
    const rowKey = getSupplementalWorkoutRowKey(row);
    if (existingRows.has(rowKey)) return;

    sheet.getRange(nextRowIndex + 1, 1, 1, SUPPLEMENTAL_WORKOUTS_HEADERS.length).setValues([row]);
    sheet.getRange(nextRowIndex + 1, 1).setNumberFormat('mmmm d, yyyy');
    existingRows.add(rowKey);
    nextRowIndex++;
  });
}

function formatExerciseValue(exercises, propertyName) {
  return exercises
    .map(exercise => exercise[propertyName] ? `${exercise.workout}: ${exercise[propertyName]}` : '')
    .filter(Boolean)
    .join('\n');
}

function groupSupplementalWorkoutRows(activityDate, exercises) {
  const groups = [];

  exercises.forEach(exercise => {
    let group = groups.find(existingGroup => existingGroup.category === exercise.category);
    if (!group) {
      group = { category: exercise.category, exercises: [] };
      groups.push(group);
    }
    group.exercises.push(exercise);
  });

  return groups.map(group => [
    activityDate,
    group.category,
    formatExerciseValue(group.exercises, 'sets'),
    formatExerciseValue(group.exercises, 'repsHoldTime'),
    formatExerciseValue(group.exercises, 'notes')
  ]);
}

function updateSupplementalWorkoutsSheet(spreadsheet, activities) {
  const sheet = spreadsheet.getSheetByName(SUPPLEMENTAL_WORKOUTS_SHEET_NAME)
    || spreadsheet.insertSheet(SUPPLEMENTAL_WORKOUTS_SHEET_NAME);
  const sheetData = ensureSupplementalWorkoutsHeader(sheet, sheet.getDataRange().getValues());
  const rows = [];

  activities.forEach(activity => {
    const activityDate = parseActivityDate(activity);
    rows.push(...groupSupplementalWorkoutRows(activityDate, parseSupplementalWorkoutDetails(activity.description).exercises));
  });

  appendSupplementalWorkoutRows(sheet, sheetData, rows);

  sheet.getRange(1, 1, Math.max(sheetData.length + rows.length, 1), SUPPLEMENTAL_WORKOUTS_HEADERS.length)
    .setWrap(true)
    .setVerticalAlignment('middle');
  sheet.getRange(1, 1, 1, SUPPLEMENTAL_WORKOUTS_HEADERS.length).setFontWeight('bold');
  sheet.setColumnWidth(1, 140);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(3, 80);
  sheet.setColumnWidth(4, 140);
  sheet.setColumnWidth(5, 240);
}