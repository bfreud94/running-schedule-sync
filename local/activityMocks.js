const { existsSync, readdirSync, readFileSync } = require('node:fs');
const path = require('node:path');

const MOCKS_DIR = path.join(__dirname, 'mocks');

function loadActivityMocks(startDate = new Date(0)) {
  if (!existsSync(MOCKS_DIR)) {
    throw new Error('No activity mocks found. Run `npm run sync:mocks` before running locally.');
  }

  const activities = readdirSync(MOCKS_DIR)
    .filter(fileName => fileName.endsWith('.json'))
    .sort()
    .map(fileName => JSON.parse(readFileSync(path.join(MOCKS_DIR, fileName), 'utf8')))
    .filter(activity => activity.start_date_local && new Date(activity.start_date_local) >= startDate)
    .sort((left, right) => left.start_date_local.localeCompare(right.start_date_local));

  if (activities.length === 0) {
    throw new Error(`No activity mocks found on or after ${startDate.toISOString().slice(0, 10)}.`);
  }

  return activities;
}

module.exports = { MOCKS_DIR, loadActivityMocks };
