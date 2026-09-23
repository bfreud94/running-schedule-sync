const { execFileSync } = require('node:child_process');
const { existsSync, readFileSync, writeFileSync } = require('node:fs');
const path = require('node:path');

const root = path.resolve(__dirname, '..');
const sinceDate = process.argv[2];
const mockArgs = [path.join(__dirname, 'sync-activity-mocks.js')];
if (sinceDate) mockArgs.push(sinceDate);

function run(command, args) {
  execFileSync(process.execPath, args, {
    cwd: root,
    stdio: 'inherit'
  });
  console.log(`Completed ${command}.`);
}

run('activity mock download', mockArgs);
const localArgs = [
  path.join(root, 'sync-last-month.js'),
  '--activity-mocks',
  '--sheet-fixture',
  './fixtures/spreadsheet.json'
];
if (sinceDate) localArgs.push(sinceDate);
localArgs.push('mocks');
run('local spreadsheet sync', localArgs);

const outputPath = path.join(root, 'output', 'spreadsheet.json');
const htmlOutputPath = path.join(root, 'output', 'spreadsheet.html');
const fixturePath = path.join(root, 'fixtures', 'spreadsheet.json');
if (!existsSync(outputPath) || !existsSync(htmlOutputPath)) {
  throw new Error('Local sync did not produce both output/spreadsheet.json and output/spreadsheet.html.');
}
const syncedWorkbook = JSON.parse(readFileSync(outputPath, 'utf8').replace(/^\uFEFF/, ''));
writeFileSync(fixturePath, `${JSON.stringify({ sheets: syncedWorkbook.sheets }, null, 2)}\n`);
console.log('Persisted synchronized sheets to fixtures/spreadsheet.json.');
console.log('Updated output/spreadsheet.html.');
