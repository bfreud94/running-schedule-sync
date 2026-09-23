const { getAuthenticatedClient } = require('./googleAuth');

getAuthenticatedClient()
  .then(() => console.log('Google Sheets local authentication is ready.'))
  .catch(error => {
    console.error(error.message);
    process.exitCode = 1;
  });