const COLORS = {
  green: '#34a853',
  red: 'red',
  yellow: 'yellow',
  white: '#ffffff',
};

function findWeekRowIndex(sheetData, targetMonday) {
  return sheetData.findIndex(row => areSameDate(getWeekDate(row[0]), targetMonday));
}

function formatWeekLabel(weekNumber, monday) {
  return `WEEK ${weekNumber} (${monday.getMonth() + 1}/${monday.getDate()})`;
}

function getNextWeekNumber(sheetData, targetMonday) {
  for (let rowIndex = sheetData.length - 1; rowIndex >= 0; rowIndex--) {
    const label = String(sheetData[rowIndex][0] || '').trim();
    const weekNumberMatch = label.toUpperCase().match(/^WEEK\s+(\d+)/);
    const weekDate = getWeekDate(label);
    if (!weekNumberMatch || !weekDate) continue;

    const weeksElapsed = Math.max(1, Math.round(getDayOffset(targetMonday, weekDate) / 7));
    return parseInt(weekNumberMatch[1], 10) + weeksElapsed;
  }

  return 1;
}

function getLastPopulatedRowIndex(sheetData) {
  for (let rowIndex = sheetData.length - 1; rowIndex >= 0; rowIndex--) {
    if (sheetData[rowIndex].some(value => String(value || '').trim() !== '')) return rowIndex;
  }
  return -1;
}

function appendWeekRow(actualSheet, sheetData, targetMonday) {
  // Trailing blank rows in the data range shouldn't push the new week further down.
  const rowIndex = getLastPopulatedRowIndex(sheetData) + 1;
  const weekLabel = formatWeekLabel(getNextWeekNumber(sheetData, targetMonday), targetMonday);
  actualSheet.getRange(rowIndex + 1, 1).setValue(weekLabel);

  // Carry the previous row's total formula forward; the local mock has no formula support.
  const previousTotalFormula = rowIndex > 0 ? actualSheet.getRange(rowIndex, 9).getFormula?.() : '';
  if (previousTotalFormula) {
    actualSheet.getRange(rowIndex + 1, 9).setFormula(previousTotalFormula);
  }

  return rowIndex;
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

function updateDailyCells(actualSheet, rowIndex, plannedRow, dailyMiles, dailyWorkouts, todayOffset, dailySupplementalWorkouts = []) {
  for (let dayIndex = 0; dayIndex <= todayOffset; dayIndex++) {
    const targetCell = actualSheet.getRange(rowIndex + 1, dayIndex + 2);
    const existingValue = String(targetCell.getValue() || '').replace(/\u00A0/g, ' ').trim();
    const isRest = isRestValue(existingValue);
    const workoutText = [...dailyWorkouts[dayIndex], ...(dailySupplementalWorkouts[dayIndex] || [])].join('\n');
    const needsWorkoutBackfill = workoutText && !existingValue.includes(workoutText);
    const canUpdate = dayIndex === todayOffset || existingValue === '' || isRest || existingValue.toLowerCase().startsWith('data [') || needsWorkoutBackfill;

    if (!canUpdate) continue;

    const stravaMiles = Math.floor(dailyMiles[dayIndex] * 100) / 100;
    const milesText = stravaMiles === 0 ? (isRest ? existingValue : 'Rest') : `${stravaMiles} miles`;
    targetCell.setValue(workoutText ? `${milesText}\n${workoutText}` : milesText);
    if (targetCell.setFontWeight) targetCell.setFontWeight('normal');

    if (plannedRow) {
      applyCellStyle(targetCell, getDailyCellStyle(plannedRow[dayIndex + 1], stravaMiles));
    }
  }
}

function updateTotalCell(actualSheet, rowIndex, plannedRow) {
  const totalCell = actualSheet.getRange(rowIndex + 1, 9);
  let actualTotalMiles = 0;
  for (let column = 2; column <= 8; column++) {
    actualTotalMiles += parseMilesFromCell(actualSheet.getRange(rowIndex + 1, column).getValue());
  }
  actualTotalMiles = Math.floor(actualTotalMiles * 100) / 100;
  totalCell.setValue(`${actualTotalMiles} miles`);

  if (plannedRow) {
    const plannedTotalMiles = parseMilesFromCell(plannedRow[8]);
    if (plannedTotalMiles > 0) {
      applyCellStyle(totalCell, getTotalCellStyle(actualTotalMiles, plannedTotalMiles));
    }
  }

  return actualTotalMiles;
}