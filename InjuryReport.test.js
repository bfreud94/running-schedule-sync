const assert = require('node:assert/strict');
const { readFileSync } = require('node:fs');
const test = require('node:test');
const vm = require('node:vm');

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

test('clearing a future row removes backgrounds from every populated column', () => {
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

  vm.runInContext(`${source}\nglobalThis.clearTarget = clearInjuryRowBackground;`, context);
  context.clearTarget(sheet, 5, 2);

  assert.deepEqual(backgrounds, {
    '6,1': null,
    '6,2': null,
    '6,3': null
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