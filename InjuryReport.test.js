const assert = require('node:assert/strict');
const { readFileSync } = require('node:fs');
const test = require('node:test');
const vm = require('node:vm');

test('parses the body part from the Area line after Injury Report', () => {
  const context = vm.createContext({ console });
  const source = readFileSync('InjuryReport.js', 'utf8');

  vm.runInContext(`${source}\nglobalThis.injuryTarget = parseInjuryReports;`, context);

  const reports = context.injuryTarget(
    'Injury Report:\nArea: Left Groin\nSeverity: 3/10\nTight after workout'
  );
  assert.equal(reports.length, 1);
  assert.equal(reports[0].bodyPart, 'Left Groin');
  assert.equal(reports[0].description, 'Severity: 3/10\nTight after workout');
  assert.equal(reports[0].severity, 3);
  assert.deepEqual([...context.injuryTarget('Injury Report:\nLeft Groin: Tight after workout')], []);
});

test('cuts off the injury report body at the first blank line', () => {
  const context = vm.createContext({ console });
  const source = readFileSync('InjuryReport.js', 'utf8');

  vm.runInContext(`${source}\nglobalThis.injuryTarget = parseInjuryReports;`, context);

  const reports = context.injuryTarget(
    'Injury Report:\nArea: Left Foot\nSeverity: 3/10\nSymptoms: Mildly sharp pain\n\nSupplemental Workouts:\nCore\n1. Planks (2x1:30)'
  );

  assert.equal(reports.length, 1);
  assert.equal(reports[0].bodyPart, 'Left Foot');
  assert.equal(reports[0].description, 'Severity: 3/10\nSymptoms: Mildly sharp pain');
  assert.equal(reports[0].severity, 3);
});

test('parses every Area block in a multi-body-part Injury Report, including 0/10 severity', () => {
  const context = vm.createContext({ console });
  const source = readFileSync('InjuryReport.js', 'utf8');

  vm.runInContext(`${source}\nglobalThis.injuryTarget = parseInjuryReports;`, context);

  const reports = context.injuryTarget(
    'Injury Report:\nArea: Left Calf\nSeverity: 1/10\nSymptoms: Soreness prior to the run\n\nArea: Left Foot\nSeverity: 0/10\nSymptoms: None'
  );

  assert.equal(reports.length, 2);
  assert.equal(reports[0].bodyPart, 'Left Calf');
  assert.equal(reports[0].severity, 1);
  assert.equal(reports[1].bodyPart, 'Left Foot');
  assert.equal(reports[1].severity, 0);
  assert.equal(reports[1].description, 'Severity: 0/10\nSymptoms: None');
});

test('new body-part columns copy backgrounds from the previous column', () => {
  const backgrounds = {
    '1,2': '#eeeeee',
    '2,2': '#34a853',
    '3,2': 'red'
  };
  const sheet = {
    getRange(row, column) {
      const key = `${row},${column}`;
      return {
        getBackground: () => backgrounds[key] || '#ffffff',
        setBackground: value => { backgrounds[key] = value; }
      };
    }
  };
  const context = vm.createContext({ console });
  const source = readFileSync('InjuryReport.js', 'utf8');

  vm.runInContext(`${source}\nglobalThis.copyTarget = copyPreviousColumnBackgrounds;`, context);
  context.copyTarget(sheet, 2, 3);

  assert.equal(backgrounds['1,3'], '#eeeeee');
  assert.equal(backgrounds['2,3'], '#34a853');
  assert.equal(backgrounds['3,3'], 'red');
});

test('severity colors blend from green to red in half-point increments', () => {
  const context = vm.createContext({
    console,
    COLORS: { green: '#34a853' }
  });
  const source = readFileSync('InjuryReport.js', 'utf8');

  vm.runInContext(`${source}\nglobalThis.colorTarget = getSeverityColor;`, context);

  assert.equal(context.colorTarget(0), '#34a853');
  assert.equal(context.colorTarget(0.5), '#3ea04f');
  assert.equal(context.colorTarget(5), '#9a542a');
  assert.equal(context.colorTarget(9.5), '#f50804');
  assert.equal(context.colorTarget(10), '#ff0000');
});

test('row highlighting includes the date and every body-part column', () => {
  const backgrounds = {};
  const sheet = {
    getRange(row, column) {
      return {
        setBackground(value) {
          backgrounds[`${row},${column}`] = value;
        }
      };
    }
  };
  const context = vm.createContext({ console });
  const source = readFileSync('InjuryReport.js', 'utf8');

  vm.runInContext(`${source}\nglobalThis.highlightTarget = highlightInjuryRow;`, context);
  context.highlightTarget(sheet, 2, 2, '#34a853');

  assert.deepEqual(backgrounds, {
    '3,1': '#34a853',
    '3,2': '#34a853',
    '3,3': '#34a853'
  });
});

test('clearing a future row removes backgrounds and values from body-part cells', () => {
  const backgrounds = {};
  const values = {};
  const sheet = {
    getRange(row, column) {
      return {
        setBackground(value) {
          backgrounds[`${row},${column}`] = value;
        },
        setValue(value) {
          values[`${row},${column}`] = value;
        }
      };
    }
  };
  const context = vm.createContext({ console });
  const source = readFileSync('InjuryReport.js', 'utf8');

  vm.runInContext(`${source}\nglobalThis.clearTarget = clearFutureInjuryRow;`, context);
  context.clearTarget(sheet, 5, 2);

  assert.deepEqual(backgrounds, {
    '6,1': null,
    '6,2': null,
    '6,3': null
  });
  assert.deepEqual(values, {
    '6,2': '',
    '6,3': ''
  });
});

test('status text is written in every body-part cell', () => {
  const values = {};
  const sheet = {
    getRange(row, column) {
      return {
        setValue(value) {
          values[`${row},${column}`] = value;
        }
      };
    }
  };
  const context = vm.createContext({ console });
  const source = readFileSync('InjuryReport.js', 'utf8');

  vm.runInContext(`${source}\nglobalThis.statusTarget = setInjuryRowValues;`, context);
  context.statusTarget(sheet, 1, 2, 'Rest day');

  assert.deepEqual(values, {
    '2,2': 'Rest day',
    '2,3': 'Rest day'
  });
});

test('creates a Status column when no injury body-part columns exist', () => {
  let headerValue;
  const sheet = {
    getRange(row, column) {
      return {
        setValue: value => { if (row === 1 && column === 2) headerValue = value; }
      };
    }
  };
  const context = vm.createContext({ console });
  const source = readFileSync('InjuryReport.js', 'utf8');

  vm.runInContext(`${source}\nglobalThis.statusColumnTarget = ensureStatusColumn;`, context);
  const headers = ['Date'];
  context.statusColumnTarget(sheet, headers);

  assert.equal(headerValue, 'Status');
  assert.deepEqual([...headers], ['Date', 'Status']);
});

test('fills blank body-part cells when another area has an injury', () => {
  const values = { '2,2': '', '2,3': '' };
  const sheet = {
    getRange(row, column) {
      return {
        getValue: () => values[`${row},${column}`],
        setValue(value) {
          values[`${row},${column}`] = value;
        }
      };
    }
  };
  const context = vm.createContext({ console });
  const source = readFileSync('InjuryReport.js', 'utf8');

  vm.runInContext(`${source}\nglobalThis.blankTarget = setBlankInjuryRowValues;`, context);
  context.blankTarget(sheet, 1, 2, 'No injuries reported');

  assert.deepEqual(values, {
    '2,2': 'No injuries reported',
    '2,3': 'No injuries reported'
  });
});