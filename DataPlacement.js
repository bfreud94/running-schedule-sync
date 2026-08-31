const ACTUAL_SHEET_NAME = 'Actual Runs';
const PLANNED_SHEET_NAME = 'Planned Schedule';

const COLORS = {
  green: "#34a853",
  red: "red",
  yellow: "yellow",
  white: "#ffffff",
};

function parseWeekMondayDate(weekString) {
  if (!weekString) return null;
  const match = String(weekString).trim().match(/\((\d{1,2})\/(\d{1,2})\)/);
  if (!match) return null;

  const month = parseInt(match[1], 10) - 1;
  const day = parseInt(match[2], 10);
  const now = new Date();
  let year = now.getFullYear();

  let parsedDate = new Date(year, month, day, 0, 0, 0, 0);

  if (parsedDate > now && (parsedDate - now) > 180 * 24 * 60 * 60 * 1000) {
    parsedDate.setFullYear(year - 1);
  }

  return parsedDate;
}

function getMostRecentMonday() {
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  const dayOfWeek = today.getDay();
  const diffToMonday = (dayOfWeek === 0 ? -6 : 1 - dayOfWeek);
  
  const monday = new Date(today);
  monday.setDate(today.getDate() + diffToMonday);
  monday.setHours(0, 0, 0, 0);
  return monday;
}

function fetchStravaActivitiesForWeek(startDate) {
  const accessToken = getAccessToken();
  const afterTimestamp = Math.floor(startDate.getTime() / 1000) - 86400;
  const endpoint = `https://www.strava.com/api/v3/athlete/activities?after=${afterTimestamp}&per_page=100`;

  const options = {
    method: 'get',
    headers: { Authorization: 'Bearer ' + accessToken },
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(endpoint, options);
  const statusCode = response.getResponseCode();

  if (statusCode !== 200) {
    Logger.log(`Strava API Request Failed (${statusCode}): ${response.getContentText()}`);
    return [];
  }

  return JSON.parse(response.getContentText());
}

function parseMilesFromCell(val) {
  if (val === null || val === undefined) return 0;
  const strVal = String(val).replace(/\u00A0/g, ' ').trim().toLowerCase();
  if (strVal === "" || strVal === "rest" || strVal.startsWith("data [")) return 0;
  
  const extractedNum = parseFloat(strVal);
  return isNaN(extractedNum) ? 0 : extractedNum;
}

function syncStravaToActualRuns() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const actualSheet = spreadsheet.getSheetByName(ACTUAL_SHEET_NAME);
  const plannedSheet = spreadsheet.getSheetByName(PLANNED_SHEET_NAME);

  if (!actualSheet) {
    Logger.log(`Error: Sheet named "${ACTUAL_SHEET_NAME}" was not found.`);
    return;
  }

  const actualData = actualSheet.getDataRange().getValues();
  if (actualData.length === 0) return;

  const plannedData = plannedSheet ? plannedSheet.getDataRange().getValues() : [];

  const targetMonday = getMostRecentMonday();
  const today = new Date();
  today.setHours(0, 0, 0, 0);

  const diffTime = today.getTime() - targetMonday.getTime();
  const todayOffset = Math.min(6, Math.max(0, Math.round(diffTime / (1000 * 3600 * 24))));

  let targetRowIndex = -1;

  for (let rowIndex = 0; rowIndex < actualData.length; rowIndex++) {
    const colAValue = actualData[rowIndex][0];
    if (!colAValue) continue;

    let rowDate = null;

    if (colAValue instanceof Date) {
      rowDate = new Date(colAValue.getFullYear(), colAValue.getMonth(), colAValue.getDate(), 0, 0, 0, 0);
    } else {
      const strVal = String(colAValue).trim();
      if (strVal.toUpperCase().startsWith('WEEK ')) {
        rowDate = parseWeekMondayDate(strVal);
      }
    }

    if (rowDate && 
        rowDate.getMonth() === targetMonday.getMonth() && 
        rowDate.getDate() === targetMonday.getDate() &&
        rowDate.getFullYear() === targetMonday.getFullYear()) {
      targetRowIndex = rowIndex;
      break;
    }
  }

  if (targetRowIndex === -1) {
    Logger.log(`No matching row found for Monday (${targetMonday.toLocaleDateString()}). Skipping execution.`);
    return;
  }

  Logger.log(`Targeting current week row at index ${targetRowIndex + 1}`);

  const rawActivities = fetchStravaActivitiesForWeek(targetMonday);
  const metersToMiles = 1 / 1609.344;
  const dailyMiles = [0, 0, 0, 0, 0, 0, 0];

  rawActivities.forEach(act => {
    const activityType = String(act.type || '').toLowerCase();
    const sportType = String(act.sport_type || '').toLowerCase();

    if (activityType.includes('run') || sportType.includes('run')) {
      const dateStr = act.start_date_local.split('T')[0];
      const parts = dateStr.split('-');
      const actDate = new Date(parseInt(parts[0], 10), parseInt(parts[1], 10) - 1, parseInt(parts[2], 10), 0, 0, 0, 0);

      const dayDiff = Math.round((actDate.getTime() - targetMonday.getTime()) / (1000 * 3600 * 24));
      if (dayDiff >= 0 && dayDiff <= 6) {
        dailyMiles[dayDiff] += act.distance * metersToMiles;
      }
    }
  });

  const plannedRow = (targetRowIndex < plannedData.length) ? plannedData[targetRowIndex] : null;

  for (let dayIdx = 0; dayIdx <= todayOffset; dayIdx++) {
    const sheetColIndex = dayIdx + 2; 
    const arrayColIndex = dayIdx + 1;

    const targetCell = actualSheet.getRange(targetRowIndex + 1, sheetColIndex);
    const existingVal = String(targetCell.getValue() || "").replace(/\u00A0/g, ' ').trim();

    const isToday = (dayIdx === todayOffset);
    const isRestOnOwnLine = /(^|\n)rest($|\n)/i.test(existingVal);
    const isEmptyOrRest = (existingVal === "" || isRestOnOwnLine || existingVal.toLowerCase().startsWith("data ["));

    if (isToday || isEmptyOrRest) {
      const stravaMiles = Math.floor(dailyMiles[dayIdx] * 100) / 100;
      
      let formattedValue = "";
      if (stravaMiles === 0) {
        formattedValue = isRestOnOwnLine ? existingVal : "Rest";
      } else {
        formattedValue = `${stravaMiles} miles`;
      }

      targetCell.setValue(formattedValue);

      if (plannedRow) {
        const plannedRawVal = String(plannedRow[arrayColIndex] || "").trim();
        const isPlannedRest = /(^|\n)rest($|\n)/i.test(plannedRawVal);

        if (isPlannedRest) {
          if (stravaMiles === 0) {
            targetCell.setBackground(COLORS["green"]);
            targetCell.setFontColor(COLORS["white"]);
          } else {
            targetCell.setBackground(COLORS["red"]);
            targetCell.setFontColor(COLORS["white"]);
          }
        } else {
          const plannedMiles = parseMilesFromCell(plannedRow[arrayColIndex]);
          const diff = Math.abs(stravaMiles - plannedMiles);

          if (diff <= 0.25) {
            targetCell.setBackground(COLORS["green"]);
            targetCell.setFontColor(COLORS["white"]);
          } else if (diff > 0.25 && diff <= 0.50) {
            targetCell.setBackground(COLORS["yellow"]);
            targetCell.setFontColor(null);
          } else {
            targetCell.setBackground(COLORS["red"]);
            targetCell.setFontColor(COLORS["white"]);
          }
        }
      }
    }
  }

  SpreadsheetApp.flush();

  const totalColIndex = 9;
  const totalCell = actualSheet.getRange(targetRowIndex + 1, totalColIndex);
  
  const actualTotalMiles = parseMilesFromCell(totalCell.getValue());

  if (plannedRow) {
    const plannedTotalMiles = parseMilesFromCell(plannedRow[8]);

    if (plannedTotalMiles > 0) {
      const pctDiff = Math.abs((actualTotalMiles - plannedTotalMiles) / plannedTotalMiles);

      if (pctDiff <= 0.03) {
        totalCell.setBackground(COLORS["green"]);
        totalCell.setFontColor(COLORS["white"]);
      } else if (pctDiff <= 0.10) {
        totalCell.setBackground(COLORS["yellow"]);
        totalCell.setFontColor(null);
      } else {
        totalCell.setBackground(COLORS["red"]);
        totalCell.setFontColor(COLORS["white"]);
      }
    }
  }

  Logger.log(`Successfully updated row ${targetRowIndex + 1}. Column I evaluated total: ${actualTotalMiles}`);
}