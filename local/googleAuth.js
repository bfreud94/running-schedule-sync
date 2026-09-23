const http = require('node:http');
const { existsSync, readFileSync, writeFileSync, mkdirSync } = require('node:fs');
const path = require('node:path');
const { google } = require('googleapis');

const ROOT = path.resolve(__dirname, '..');
const CLIENT_PATH = path.join(__dirname, 'google-oauth-client.json');
const TOKEN_DIR = path.join(ROOT, '.local');
const TOKEN_PATH = path.join(TOKEN_DIR, 'google-token.json');
const SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly'];

function loadClientConfig() {
  if (!existsSync(CLIENT_PATH)) {
    throw new Error(
      `Missing ${path.relative(ROOT, CLIENT_PATH)}. Download a Google Cloud OAuth Desktop app client JSON and save it at that path.`
    );
  }
  const config = JSON.parse(readFileSync(CLIENT_PATH, 'utf8'));
  return config.installed || config.web || config;
}

function getRedirectUri(config) {
  const redirectUri = process.env.GOOGLE_OAUTH_REDIRECT_URI || config.redirect_uris?.[0];
  if (!redirectUri) {
    throw new Error(
      'OAuth client has no redirect URI. Add http://localhost:43821/oauth2callback to Authorized redirect URIs in Google Cloud, or download a Desktop app OAuth client JSON.'
    );
  }
  return redirectUri;
}

async function getAuthenticatedClient() {
  const config = loadClientConfig();
  const redirectUri = getRedirectUri(config);
  const auth = new google.auth.OAuth2(config.client_id, config.client_secret, redirectUri);

  try {
    auth.setCredentials(JSON.parse(readFileSync(TOKEN_PATH, 'utf8')));
    await auth.getAccessToken();
    return auth;
  } catch {
    mkdirSync(TOKEN_DIR, { recursive: true });
    const authorizationUrl = auth.generateAuthUrl({ access_type: 'offline', scope: SCOPES, prompt: 'consent' });
    console.log(`Open this URL in your browser:\n${authorizationUrl}`);
    const callback = new URL(redirectUri);

    const code = await new Promise((resolve, reject) => {
      const server = http.createServer((request, response) => {
        const callbackUrl = new URL(request.url, redirectUri);
        const error = callbackUrl.searchParams.get('error');
        const authorizationCode = callbackUrl.searchParams.get('code');
        response.end(error ? 'Authorization failed. You can close this tab.' : 'Authorized. You can close this tab.');
        server.close();
        if (error) reject(new Error(error));
        else resolve(authorizationCode);
      });
      server.on('error', reject);
      server.listen(Number(callback.port) || 80, callback.hostname);
    });

    const tokenResponse = await auth.getToken(code);
    auth.setCredentials(tokenResponse.tokens);
    writeFileSync(TOKEN_PATH, JSON.stringify(tokenResponse.tokens, null, 2));
    console.log(`Saved Google OAuth token to ${path.relative(ROOT, TOKEN_PATH)}`);
    return auth;
  }
}

module.exports = { getAuthenticatedClient };