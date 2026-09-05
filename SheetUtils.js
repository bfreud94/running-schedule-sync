const COLORS = {
  green: '#34a853',
  red: 'red',
  yellow: 'yellow',
  white: '#ffffff',
};

function findWeekRowIndex(sheetData, targetMonday) {
  return sheetData.findIndex(row => areSameDate(getWeekDate(row[0]), targetMonday));
}

function isRestValue(value) {
  return /(^|\n)rest($|\n)/i.test(String(value || '').trim());
}

function getDailyCellStyle(plannedValue, actualMiles) {
  if (isRestValue(plannedValue)) {
    return actualMiles === 0
      ? { background: COLORS.green, fontColor: COLORS.white }
      : { background: COLORS.red, fontColor: COLORS.white };
  }

  const difference = Math.abs(actualMiles - parseMilesFromCell(plannedValue));
  if (difference <= 0.25) return { background: COLORS.green, fontColor: COLORS.white };
  if (difference <= 0.50) return { background: COLORS.yellow, fontColor: null };
  return { background: COLORS.red, fontColor: COLORS.white };
}

function getTotalCellStyle(actualMiles, plannedMiles) {
  const percentageDifference = Math.abs((actualMiles - plannedMiles) / plannedMiles);
  if (percentageDifference <= 0.03) return { background: COLORS.green, fontColor: COLORS.white };
  if (percentageDifference <= 0.10) return { background: COLORS.yellow, fontColor: null };
  return { background: COLORS.red, fontColor: COLORS.white };
}

function applyCellStyle(cell, style) {
  cell.setBackground(style.background);
  cell.setFontColor(style.fontColor);
}

function updateDailyCells(actualSheet, rowIndex, plannedRow, dailyMiles, dailyWorkouts, todayOffset) {
  for (let dayIndex = 0; dayIndex <= todayOffset; dayIndex++) {
    const targetCell = actualSheet.getRange(rowIndex + 1, dayIndex + 2);
    const existingValue = String(targetCell.getValue() || '').replace(/\u00A0/g, ' ').trim();
    const isRest = isRestValue(existingValue);
    const workoutText = dailyWorkouts[dayIndex].join('\n');
    const needsWorkoutBackfill = workoutText && !existingValue.includes(workoutText);
    const canUpdate = dayIndex === todayOffset || existingValue === '' || isRest || existingValue.toLowerCase().startsWith('data [') || needsWorkoutBackfill;

    if (!canUpdate) continue;

    const stravaMiles = Math.floor(dailyMiles[dayIndex] * 100) / 100;
    const milesText = stravaMiles === 0 ? (isRest ? existingValue : 'Rest') : `${stravaMiles} miles`;
    targetCell.setValue(workoutText ? `${milesText}\n${workoutText}` : milesText);

    if (plannedRow) {
      applyCellStyle(targetCell, getDailyCellStyle(plannedRow[dayIndex + 1], stravaMiles));
    }
  }
}

function updateTotalCell(actualSheet, rowIndex, plannedRow) {
  const totalCell = actualSheet.getRange(rowIndex + 1, 9);
  const actualTotalMiles = parseMilesFromCell(totalCell.getValue());

  if (plannedRow) {
    const plannedTotalMiles = parseMilesFromCell(plannedRow[8]);
    if (plannedTotalMiles > 0) {
      applyCellStyle(totalCell, getTotalCellStyle(actualTotalMiles, plannedTotalMiles));
    }
  }

  return actualTotalMiles;
}