const axios = require('axios');
const config = require('../config');
const logger = require('../logger');

let cachedToken = null;
let tokenExpiry = 0;

async function getAccessToken() {
  // Return cached token if valid (with 5-minute buffer)
  if (cachedToken && Date.now() < tokenExpiry - 5 * 60 * 1000) {
    return cachedToken;
  }

  const tokenUrl = `https://login.microsoftonline.com/${config.microsoft.tenantId}/oauth2/v2.0/token`;

  const params = new URLSearchParams({
    client_id: config.microsoft.clientId,
    scope: 'https://graph.microsoft.com/.default',
    client_secret: config.microsoft.clientSecret,
    grant_type: 'client_credentials',
  });

  const response = await axios.post(tokenUrl, params.toString(), {
    headers: { 'Content-Type': 'application/x-www-form-urlencoded' },
  });

  cachedToken = response.data.access_token;
  tokenExpiry = Date.now() + response.data.expires_in * 1000;

  logger.info('Microsoft Graph access token acquired', {
    expiresIn: response.data.expires_in,
  });

  return cachedToken;
}

function graphClient() {
  return {
    async get(url) {
      const token = await getAccessToken();
      return axios.get(url, {
        headers: { Authorization: `Bearer ${token}` },
      });
    },
    async post(url, data) {
      const token = await getAccessToken();
      return axios.post(url, data, {
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
      });
    },
    async patch(url, data) {
      const token = await getAccessToken();
      return axios.patch(url, data, {
        headers: {
          Authorization: `Bearer ${token}`,
          'Content-Type': 'application/json',
        },
      });
    },
    async delete(url) {
      const token = await getAccessToken();
      return axios.delete(url, {
        headers: { Authorization: `Bearer ${token}` },
      });
    },
  };
}

module.exports = { getAccessToken, graphClient };
