function getLocalScriptProperties() {
  process.loadEnvFile('.env');
  return {
    getProperty: key => process.env[key],
    setProperty: (key, value) => { process.env[key] = value; }
  };
}

function getAccessToken() {
  const isLocalEnvironment = typeof PropertiesService === 'undefined';
  const props = isLocalEnvironment
    ? getLocalScriptProperties()
    : PropertiesService.getScriptProperties();
  const clientId = props.getProperty('STRAVA_CLIENT_ID');
  const clientSecret = props.getProperty('STRAVA_CLIENT_SECRET');
  const refreshToken = props.getProperty('STRAVA_REFRESH_TOKEN');

  if (!clientId || !clientSecret || !refreshToken) {
    const credentialSource = isLocalEnvironment ? 'the local .env file' : 'Script Properties';
    throw new Error(`Missing credentials. Set STRAVA_CLIENT_ID, STRAVA_CLIENT_SECRET, and STRAVA_REFRESH_TOKEN in ${credentialSource}.`);
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

function fetchStravaResource(endpoint, accessToken) {
  const response = UrlFetchApp.fetch(endpoint, {
    method: 'get',
    headers: { Authorization: `Bearer ${accessToken}` },
    muteHttpExceptions: true
  });
  const statusCode = response.getResponseCode();

  if (statusCode !== 200) {
    Logger.log(`Strava API request failed (${statusCode}) for ${endpoint}: ${response.getContentText()}`);
    return null;
  }

  return JSON.parse(response.getContentText());
}

function fetchStravaActivityPage(startDate, page, accessToken) {
  const afterTimestamp = Math.floor(startDate.getTime() / 1000) - 1;
  const endpoint = `https://www.strava.com/api/v3/athlete/activities?after=${afterTimestamp}&page=${page}&per_page=100`;
  return fetchStravaResource(endpoint, accessToken) || [];
}

function fetchStravaActivitySummariesSince(startDate, accessToken) {
  const activities = [];
  let page = 1;
  let activityPage;

  do {
    activityPage = fetchStravaActivityPage(startDate, page, accessToken);
    activities.push(...activityPage);
    page++;
  } while (activityPage.length === 100);

  return activities;
}

function fetchDetailedStravaActivity(activityId, accessToken) {
  const endpoint = `https://www.strava.com/api/v3/activities/${activityId}`;
  return fetchStravaResource(endpoint, accessToken);
}

function addDescriptionToStravaActivity(activity, accessToken) {
  const detailedActivity = fetchDetailedStravaActivity(activity.id, accessToken);
  return {
    ...activity,
    description: detailedActivity ? detailedActivity.description ?? null : null
  };
}

function fetchStravaActivitiesSince(startDate = getMostRecentMonday()) {
  const accessToken = getAccessToken();
  const activities = fetchStravaActivitySummariesSince(startDate, accessToken);
  return activities.map(activity => addDescriptionToStravaActivity(activity, accessToken));
}