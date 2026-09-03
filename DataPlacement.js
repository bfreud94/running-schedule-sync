const ACTUAL_SHEET_NAME = 'Actual Runs';
const PLANNED_SHEET_NAME = 'Planned Schedule';

function syncStravaToActualRuns() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const actualSheet = spreadsheet.getSheetByName(ACTUAL_SHEET_NAME);
  const plannedSheet = spreadsheet.getSheetByName(PLANNED_SHEET_NAME);
  const injuryReportSheet = getOrCreateInjuryReportSheet(spreadsheet);

  if (!actualSheet) {
    Logger.log(`Error: Sheet named "${ACTUAL_SHEET_NAME}" was not found.`);
    return;
  }

  const actualData = actualSheet.getDataRange().getValues();
  if (actualData.length === 0) return;

  const targetMonday = getMostRecentMonday();
  const targetRowIndex = findWeekRowIndex(actualData, targetMonday);

  if (targetRowIndex === -1) {
    Logger.log(`No matching row found for Monday (${targetMonday.toLocaleDateString()}). Skipping execution.`);
    return;
  }

  Logger.log(`Targeting current week row at index ${targetRowIndex + 1}`);

  const plannedData = plannedSheet ? plannedSheet.getDataRange().getValues() : [];
  const plannedRow = targetRowIndex < plannedData.length ? plannedData[targetRowIndex] : null;
  const todayOffset = Math.min(6, Math.max(0, getDayOffset(new Date(), targetMonday)));
  const activities = fetchStravaActivitiesSince(targetMonday);
  const dailyMiles = calculateDailyRunMiles(activities, targetMonday);

  updateDailyCells(actualSheet, targetRowIndex, plannedRow, dailyMiles, todayOffset);
  updateInjuryReportSheet(injuryReportSheet, activities, targetMonday, new Date());

  SpreadsheetApp.flush();
  const actualTotalMiles = updateTotalCell(actualSheet, targetRowIndex, plannedRow);

  Logger.log(`Successfully updated row ${targetRowIndex + 1}. Column I evaluated total: ${actualTotalMiles}`);
}