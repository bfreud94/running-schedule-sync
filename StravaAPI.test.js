const assert = require('node:assert/strict');
const { readFileSync } = require('node:fs');
const test = require('node:test');
const vm = require('node:vm');

function createResponse(statusCode, body) {
  return {
    getResponseCode: () => statusCode,
    getContentText: () => JSON.stringify(body)
  };
}

test('adds descriptions to summary activities since Monday', () => {
  const requestedUrls = [];
  const firstPage = Array.from({ length: 100 }, (_, index) => ({
    id: index + 1,
    name: `Summary ${index + 1}`,
    resource_state: 2
  }));
  const context = vm.createContext({
    console,
    Logger: { log: () => {} },
    PropertiesService: {
      getScriptProperties: () => ({
        getProperty: key => ({
          STRAVA_CLIENT_ID: 'client-id',
          STRAVA_CLIENT_SECRET: 'client-secret',
          STRAVA_REFRESH_TOKEN: 'refresh-token'
        })[key],
        setProperty: () => {}
      })
    },
    UrlFetchApp: {
      fetch(url) {
        requestedUrls.push(url);

        if (url.includes('/oauth/token')) {
          return createResponse(200, { access_token: 'access-token' });
        }

        if (url.includes('/athlete/activities')) {
          const page = new URL(url).searchParams.get('page');
          return createResponse(200, page === '1' ? firstPage : [{ id: 101 }]);
        }

        const activityId = Number(url.split('/').pop());
        return createResponse(200, {
          id: activityId,
          description: `Description ${activityId}`,
          detailed_only: true
        });
      }
    }
  });
  const source = readFileSync('StravaAPI.js', 'utf8');

  vm.runInContext(`${source}\nglobalThis.fetchTarget = fetchStravaActivitiesSince;`, context);
  const activities = context.fetchTarget(new Date(2026, 7, 31));

  assert.equal(activities.length, 101);
  assert.equal(activities[0].name, 'Summary 1');
  assert.equal(activities[0].resource_state, 2);
  assert.equal(activities[0].description, 'Description 1');
  assert.equal(activities[0].detailed_only, undefined);
  assert.equal(activities[100].description, 'Description 101');
  assert.equal(requestedUrls.filter(url => url.includes('/athlete/activities')).length, 2);
  assert.equal(requestedUrls.filter(url => /\/activities\/\d+$/.test(url)).length, 101);
});