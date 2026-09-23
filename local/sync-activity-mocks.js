const { execFileSync } = require('node:child_process');
const { mkdirSync, readFileSync, writeFileSync } = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { MOCKS_DIR } = require('./activityMocks');

const cliArgs = process.argv.slice(2);
const sinceArgument = cliArgs.find(argument => argument === '--since' || argument.startsWith('--since='));
const sinceValue = sinceArgument?.startsWith('--since=')
  ? sinceArgument.slice('--since='.length)
  : sinceArgument
    ? cliArgs[cliArgs.indexOf(sinceArgument) + 1]
    : cliArgs.find(argument => !argument.startsWith('-'));
const sinceDate = new Date(sinceValue || new Date().setMonth(new Date().getMonth() - 1));
if (Number.isNaN(sinceDate.getTime())) throw new Error('Invalid --since date. Use YYYY-MM-DD.');
sinceDate.setHours(0, 0, 0, 0);

process.chdir(path.resolve(__dirname, '..'));
mkdirSync(MOCKS_DIR, { recursive: true });

globalThis.Logger = { log: (...args) => console.log(...args) };
globalThis.UrlFetchApp = {
  fetch(url, options = {}) {
    const args = ['--silent', '--show-error', '--location', '--request', String(options.method || 'get').toUpperCase()];
    Object.entries(options.headers || {}).forEach(([name, value]) => args.push('--header', `${name}: ${value}`));
    Object.entries(options.payload || {}).forEach(([name, value]) => args.push('--data-urlencode', `${name}=${value}`));
    args.push('--write-out', '\n%{http_code}', String(url));
    const output = execFileSync('curl', args, { encoding: 'utf8', maxBuffer: 10 * 1024 * 1024 });
    const separator = output.lastIndexOf('\n');
    return {
      getContentText: () => output.slice(0, separator),
      getResponseCode: () => Number(output.slice(separator + 1).trim())
    };
  }
};

const source = [
  'StravaAPI.js',
  'RunningMetrics.js',
  'DateUtils.js'
].map(fileName => readFileSync(fileName, 'utf8')).join('\n');
vm.runInThisContext(`${source}\nglobalThis.__fetchActivities = fetchStravaActivitiesSince;`, { filename: 'strava-source.js' });

const activities = globalThis.__fetchActivities(sinceDate)
  .filter(activity => {
    const type = String(activity.type || '').toLowerCase();
    const sportType = String(activity.sport_type || '').toLowerCase();
    return type.includes('run') || sportType.includes('run');
  });

for (const activity of activities) {
  const date = activity.start_date_local.slice(0, 10);
  const fileName = `${date}-${activity.id}.json`;
  writeFileSync(path.join(MOCKS_DIR, fileName), `${JSON.stringify(activity, null, 2)}\n`);
}

console.log(`Saved ${activities.length} running activity mocks to ${path.relative(process.cwd(), MOCKS_DIR)}.`);
