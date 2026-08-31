function getAccessToken() {
  const props = PropertiesService.getScriptProperties();
  const clientId = props.getProperty('STRAVA_CLIENT_ID');
  const clientSecret = props.getProperty('STRAVA_CLIENT_SECRET');
  const refreshToken = props.getProperty('STRAVA_REFRESH_TOKEN');

  if (!clientId || !clientSecret || !refreshToken) {
    throw new Error('Missing credentials. Set STRAVA_CLIENT_ID, STRAVA_CLIENT_SECRET, and STRAVA_REFRESH_TOKEN in Script Properties.');
  }

  const url = 'https://www.strava.com/api/v3/oauth/token';
  const payload = {
    client_id: clientId,
    client_secret: clientSecret,
    refresh_token: refreshToken,
    grant_type: 'refresh_token'
  };

  const options = {
    method: 'post',
    payload: payload,
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const json = JSON.parse(response.getContentText());

  if (response.getResponseCode() !== 200) {
    throw new Error(`Auth Error (${response.getResponseCode()}): ${json.message || response.getContentText()}`);
  }

  if (json.refresh_token && json.refresh_token !== refreshToken) {
    props.setProperty('STRAVA_REFRESH_TOKEN', json.refresh_token);
  }

  return json.access_token;
}

function syncStravaActivities() {
  try {
    const accessToken = getAccessToken();
    const endpoint = 'https://www.strava.com/api/v3/athlete/activities?per_page=50';
    
    const options = {
      method: 'get',
      headers: { Authorization: 'Bearer ' + accessToken },
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(endpoint, options);
    const statusCode = response.getResponseCode();
    const data = JSON.parse(response.getContentText());
    console.log(data)

    if (statusCode !== 200) {
      Logger.log(`API Fetch Failed: ${response.getContentText()}`);
      return;
    }
  } catch (err) {
    Logger.log(`Execution Error: ${err.message}`);
  }
}