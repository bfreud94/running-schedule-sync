function getInjuryReportSheetName(date = new Date()) {
  return `${date.getFullYear()} Injury Report`;
}

function getOrCreateInjuryReportSheet(spreadsheet, date = new Date()) {
  const sheetName = getInjuryReportSheetName(date);
  return spreadsheet.getSheetByName(sheetName) || spreadsheet.insertSheet(sheetName);
}

function parseInjuryReport(description) {
  const match = String(description || '').match(/injury report:?\s*\r?\n\s*area:\s*([^\r\n]+)(?:\r?\n|$)([\s\S]*)/i);
  if (!match) return null;

  const bodyPart = match[1].trim();
  const reportBody = match[2].trim();
  const severityMatch = reportBody.match(/(?:^|\r?\n)severity:\s*(\d+(?:\.5)?)\s*\/\s*10(?:\r?\n|$)/i);
  const severity = severityMatch ? Number(severityMatch[1]) : 0;
  return bodyPart ? { bodyPart, description: reportBody, severity } : null;
}

function blendColor(startHex, endHex, ratio) {
  const start = startHex.match(/\w\w/g).map(channel => parseInt(channel, 16));
  const end = endHex.match(/\w\w/g).map(channel => parseInt(channel, 16));
  const channels = start.map((channel, index) => Math.round(channel + (end[index] - channel) * ratio));
  return `#${channels.map(channel => channel.toString(16).padStart(2, '0')).join('')}`;
}

function getSeverityColor(severity) {
  const normalizedSeverity = Math.min(10, Math.max(0, severity)) / 10;
  return blendColor(COLORS.green, '#ff0000', normalizedSeverity);
}

function getDateKey(date) {
  const year = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day = String(date.getDate()).padStart(2, '0');
  return `${year}-${month}-${day}`;
}

function parseSheetDate(value) {
  if (value instanceof Date) return value;
  if (!value) return null;

  const isoDateMatch = String(value).match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (isoDateMatch) {
    return new Date(
      Number(isoDateMatch[1]),
      Number(isoDateMatch[2]) - 1,
      Number(isoDateMatch[3])
    );
  }

  const parsedDate = new Date(value);
  return isNaN(parsedDate.getTime()) ? null : parsedDate;
}

function findDateRowIndex(sheetData, activityDate) {
  const activityDateKey = getDateKey(activityDate);
  return sheetData.findIndex((row, rowIndex) => {
    if (rowIndex === 0) return false;
    const rowDate = parseSheetDate(row[0]);
    return rowDate && getDateKey(rowDate) === activityDateKey;
  });
}

function copyPreviousColumnBackgrounds(sheet, columnIndex, rowCount) {
  if (columnIndex <= 1) return;

  for (let rowIndex = 0; rowIndex < rowCount; rowIndex++) {
    const previousCell = sheet.getRange(rowIndex + 1, columnIndex);
    const targetCell = sheet.getRange(rowIndex + 1, columnIndex + 1);
    targetCell.setBackground(previousCell.getBackground());
  }
}

function findOrCreateBodyPartColumn(sheet, headers, bodyPart, rowCount) {
  const existingColumnIndex = headers.findIndex(header => String(header).trim().toLowerCase() === bodyPart.toLowerCase());
  if (existingColumnIndex >= 1) return existingColumnIndex;

  let columnIndex = headers.findIndex((header, index) => index >= 1 && String(header || '').trim() === '');
  if (columnIndex === -1) columnIndex = Math.max(headers.length, 1);

  sheet.getRange(1, columnIndex + 1).setValue(bodyPart);
  headers[columnIndex] = bodyPart;
  copyPreviousColumnBackgrounds(sheet, columnIndex, rowCount);
  return columnIndex;
}

function getLastHeaderColumnIndex(headers) {
  for (let index = headers.length - 1; index >= 1; index--) {
    if (String(headers[index] || '').trim()) return index;
  }
  return 0;
}

function highlightInjuryRow(sheet, rowIndex, lastHeaderColumnIndex, color) {
  for (let columnIndex = 0; columnIndex <= lastHeaderColumnIndex; columnIndex++) {
    sheet.getRange(rowIndex + 1, columnIndex + 1).setBackground(color);
  }
}

function setInjuryRowValues(sheet, rowIndex, lastHeaderColumnIndex, value) {
  for (let columnIndex = 1; columnIndex <= lastHeaderColumnIndex; columnIndex++) {
    sheet.getRange(rowIndex + 1, columnIndex + 1).setValue(value);
  }
}

function clearInjuryRowBackground(sheet, rowIndex, lastHeaderColumnIndex) {
  highlightInjuryRow(sheet, rowIndex, lastHeaderColumnIndex, null);
}

function clearFutureInjuryRow(sheet, rowIndex, lastHeaderColumnIndex) {
  clearInjuryRowBackground(sheet, rowIndex, lastHeaderColumnIndex);
  setInjuryRowValues(sheet, rowIndex, lastHeaderColumnIndex, '');
}

function formatInjuryReportCells(sheet, rowCount, lastHeaderColumnIndex) {
  for (let rowIndex = 0; rowIndex < rowCount; rowIndex++) {
    for (let columnIndex = 0; columnIndex <= lastHeaderColumnIndex; columnIndex++) {
      const cell = sheet.getRange(rowIndex + 1, columnIndex + 1)
        .setWrap(true)
        .setVerticalAlignment('middle');

      if (columnIndex === 0) cell.setHorizontalAlignment('left');
    }
  }

  for (let columnIndex = 1; columnIndex <= lastHeaderColumnIndex; columnIndex++) {
    sheet.setColumnWidth(columnIndex + 1, 200);
  }
}

function ensureInjuryReportWeek(sheet, sheetData, weekStart) {
  if (!sheetData[0]) sheetData[0] = [];

  if (!String(sheetData[0][0] || '').trim()) {
    sheet.getRange(1, 1).setValue('Date');
    sheetData[0][0] = 'Date';
  }

  const existingDateKeys = new Set(
    sheetData.slice(1)
      .map(row => parseSheetDate(row[0]))
      .filter(date => date !== null)
      .map(getDateKey)
  );

  getWeekDates(weekStart).forEach(date => {
    const dateKey = getDateKey(date);
    if (existingDateKeys.has(dateKey)) return;

    let rowIndex = sheetData.findIndex((row, index) => index >= 1 && !row[0]);
    if (rowIndex === -1) rowIndex = sheetData.length;

    sheet.getRange(rowIndex + 1, 1).setValue(date).setNumberFormat('mmmm d');
    if (!sheetData[rowIndex]) sheetData[rowIndex] = [];
    sheetData[rowIndex][0] = date;
    existingDateKeys.add(dateKey);
  });
}

function updateInjuryReportSheet(sheet, activities, weekStart, currentDate = new Date()) {
  const sheetData = sheet.getDataRange().getValues();
  ensureInjuryReportWeek(sheet, sheetData, weekStart);

  const headers = [...sheetData[0]];
  const activitiesByDate = activities.filter(isRunningActivity).reduce((groups, activity) => {
    const dateKey = getDateKey(parseActivityDate(activity));
    if (!groups[dateKey]) groups[dateKey] = [];
    groups[dateKey].push(activity);
    return groups;
  }, {});

  const reportsByDate = Object.fromEntries(
    Object.entries(activitiesByDate).map(([dateKey, dateActivities]) => [
      dateKey,
      dateActivities
        .map(activity => parseInjuryReport(activity.description))
        .filter(report => report !== null)
    ])
  );

  Object.values(reportsByDate).flat().forEach(report => {
    findOrCreateBodyPartColumn(sheet, headers, report.bodyPart, sheetData.length);
  });

  getWeekDates(weekStart).forEach(activityDate => {
    const dateKey = getDateKey(activityDate);
    const rowIndex = findDateRowIndex(sheetData, activityDate);
    if (rowIndex === -1) {
      Logger.log(`No Injury Report row found for ${activityDate.toLocaleDateString()}.`);
      return;
    }

    if (getDateKey(activityDate) > getDateKey(currentDate)) {
      clearFutureInjuryRow(sheet, rowIndex, getLastHeaderColumnIndex(headers));
      return;
    }

    const reports = reportsByDate[dateKey] || [];
    const rowSeverity = reports.reduce((highest, report) => Math.max(highest, report.severity), 0);
    highlightInjuryRow(
      sheet,
      rowIndex,
      getLastHeaderColumnIndex(headers),
      getSeverityColor(rowSeverity)
    );

    if (reports.length === 0 && activitiesByDate[dateKey] && activitiesByDate[dateKey].length) {
      setInjuryRowValues(sheet, rowIndex, getLastHeaderColumnIndex(headers), 'No injuries reported');
    } else if (reports.length === 0) {
      setInjuryRowValues(sheet, rowIndex, getLastHeaderColumnIndex(headers), 'Rest day');
    }

    reports.forEach(report => {
      const columnIndex = headers.findIndex(header => String(header).trim().toLowerCase() === report.bodyPart.toLowerCase());
      sheet.getRange(rowIndex + 1, columnIndex + 1)
        .setValue(report.description)
        .setBackground(getSeverityColor(report.severity));
    });
  });

  formatInjuryReportCells(sheet, sheetData.length, getLastHeaderColumnIndex(headers));
}