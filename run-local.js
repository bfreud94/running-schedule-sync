const { execFileSync } = require('node:child_process');
const { readFileSync } = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const { createGoogleSheetsMock } = require('./local/GoogleSheetsMock');

const functionName = process.argv[2];
const shouldPrintResult = process.argv.includes('--json');
const fixtureArgumentIndex = process.argv.indexOf('--sheet-fixture');
const fixturePath = fixtureArgumentIndex === -1 ? null : path.resolve(process.argv[fixtureArgumentIndex + 1]);
const outputPath = path.resolve('output', 'spreadsheet.json');

if (!functionName) {
  throw new Error('Usage: node run-local.js <functionName>');
}

if (!/^[A-Za-z_$][\w$]*$/.test(functionName)) {
  throw new Error(`Invalid function name: ${functionName}`);
}

process.chdir(__dirname);

const sheetsMock = fixturePath ? createGoogleSheetsMock(fixturePath, outputPath) : null;
if (sheetsMock) {
  globalThis.SpreadsheetApp = sheetsMock.SpreadsheetApp;
}

globalThis.Logger = {
  log: (...args) => console.log(...args)
};

globalThis.UrlFetchApp = {
  fetch(url, options = {}) {
    const args = [
      '--silent',
      '--show-error',
      '--location',
      '--request',
      String(options.method || 'get').toUpperCase()
    ];

    Object.entries(options.headers || {}).forEach(([name, value]) => {
      args.push('--header', `${name}: ${value}`);
    });

    Object.entries(options.payload || {}).forEach(([name, value]) => {
      args.push('--data-urlencode', `${name}=${value}`);
    });

    args.push('--write-out', '\n%{http_code}', String(url));

    const output = execFileSync('curl', args, {
      encoding: 'utf8',
      maxBuffer: 10 * 1024 * 1024
    });
    const statusSeparator = output.lastIndexOf('\n');
    const content = output.slice(0, statusSeparator);
    const statusCode = Number(output.slice(statusSeparator + 1).trim());

    return {
      getContentText: () => content,
      getResponseCode: () => statusCode
    };
  }
};

const scriptFiles = [
  'DateUtils.js',
  'RunningMetrics.js',
  'SheetUtils.js',
  'InjuryReport.js',
  'StravaAPI.js',
  'DataPlacement.js'
];
const scriptSource = scriptFiles
  .map(fileName => readFileSync(path.join(__dirname, fileName), 'utf8'))
  .join('\n');
const exposeFunction = `\nglobalThis.__localFunction = typeof ${functionName} === 'function' ? ${functionName} : undefined;`;

vm.runInThisContext(scriptSource + exposeFunction, { filename: 'apps-script-source.js' });

if (typeof globalThis.__localFunction !== 'function') {
  throw new Error(`Function "${functionName}" was not found in the Apps Script source files.`);
}

Promise.resolve(globalThis.__localFunction())
  .then(result => {
    if (sheetsMock) {
      const changes = sheetsMock.save();
      console.log(`Saved ${changes.length} sheet changes to ${outputPath}`);
    }
    if (shouldPrintResult && result !== undefined) {
      console.log(JSON.stringify(result, null, 2));
    }
  })
  .catch(error => {
    console.error(error);
    process.exitCode = 1;
  });