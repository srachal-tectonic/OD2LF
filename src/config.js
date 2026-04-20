if (process.env.NODE_ENV !== 'production') {
  require('dotenv').config();
}

const required = [
  'MS_TENANT_ID',
  'MS_CLIENT_ID',
  'MS_CLIENT_SECRET',
  'SHAREPOINT_SITE_ID',
  'SHAREPOINT_TARGET_FOLDER_PATH',
  'WEBHOOK_CLIENT_STATE',
  'LF_SERVER_URL',
  'LF_REPO_NAME',
  'LF_USERNAME',
  'LF_PASSWORD',
];

const missing = required.filter((key) => !process.env[key]);
if (missing.length > 0) {
  console.error(`Missing required environment variables: ${missing.join(', ')}`);
  console.error('See .env.example for reference.');
  process.exit(1);
}

module.exports = {
  port: process.env.PORT || 3000,
  logLevel: process.env.LOG_LEVEL || 'info',
  appUrl: process.env.APP_URL || `http://localhost:${process.env.PORT || 3000}`,

  microsoft: {
    tenantId: process.env.MS_TENANT_ID,
    clientId: process.env.MS_CLIENT_ID,
    clientSecret: process.env.MS_CLIENT_SECRET,
    siteId: process.env.SHAREPOINT_SITE_ID,
    targetFolderPath: process.env.SHAREPOINT_TARGET_FOLDER_PATH,
    webhookClientState: process.env.WEBHOOK_CLIENT_STATE,
  },

  laserfiche: {
    serverUrl: process.env.LF_SERVER_URL,
    repoName: process.env.LF_REPO_NAME,
    username: process.env.LF_USERNAME,
    password: process.env.LF_PASSWORD,
    destinationFolderId: parseInt(process.env.LF_DESTINATION_FOLDER_ID || '1', 10),
    verifySsl: process.env.LF_VERIFY_SSL !== 'false',
  },
};
