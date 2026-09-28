const { runSpreadsheetTabSync } = require('./syncSpreadsheetTabs');

// One pass keeps a single auth round trip and one fixture rewrite.
runSpreadsheetTabSync(['Planned Schedule', 'Supplemental Workouts']);
