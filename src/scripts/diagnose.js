/**
 * Diagnose SharePoint webhook and connectivity issues.
 * Run with: node src/scripts/diagnose.js
 */
require('dotenv').config();
const config = require('../config');
const { graphClient } = require('../services/graph-auth');

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';
const SITE_ID = config.microsoft.siteId;

async function diagnose() {
  const client = graphClient();

  // 1. Test Graph API auth
  console.log('\n=== 1. Testing Graph API Authentication ===');
  try {
    const me = await client.get(`${GRAPH_BASE}/organization`);
    console.log('  AUTH OK - Connected to tenant:', me.data.value?.[0]?.displayName || 'unknown');
  } catch (err) {
    console.error('  AUTH FAILED:', err.response?.data?.error?.message || err.message);
    return;
  }

  // 2. Verify site access
  console.log('\n=== 2. Verifying SharePoint Site ===');
  console.log('  Site ID from .env:', SITE_ID);
  try {
    const site = await client.get(`${GRAPH_BASE}/sites/${SITE_ID}`);
    console.log('  SITE OK -', site.data.displayName);
    console.log('  Web URL:', site.data.webUrl);
  } catch (err) {
    console.error('  SITE FAILED:', err.response?.data?.error?.message || err.message);
    console.log('  --> Your SHAREPOINT_SITE_ID may be wrong.');
    console.log('  --> Try the full format: yourdomain.sharepoint.com,site-guid,web-guid');
    return;
  }

  // 3. Verify drive access
  console.log('\n=== 3. Verifying Default Drive ===');
  try {
    const drive = await client.get(`${GRAPH_BASE}/sites/${SITE_ID}/drive`);
    console.log('  DRIVE OK -', drive.data.name, `(${drive.data.driveType})`);
    console.log('  Drive ID:', drive.data.id);
  } catch (err) {
    console.error('  DRIVE FAILED:', err.response?.data?.error?.message || err.message);
  }

  // 4. Check the target folder path
  console.log('\n=== 4. Checking Target Folder Path ===');
  console.log('  Configured path:', config.microsoft.targetFolderPath);
  try {
    // Try to access the folder using the path
    const folderUrl = `${GRAPH_BASE}/sites/${SITE_ID}${config.microsoft.targetFolderPath}`;
    console.log('  Requesting:', folderUrl);
    const folder = await client.get(folderUrl);
    console.log('  FOLDER OK -', folder.data.name);
    console.log('  Folder ID:', folder.data.id);
  } catch (err) {
    console.error('  FOLDER FAILED:', err.response?.data?.error?.message || err.message);
    console.log('  --> Your SHAREPOINT_TARGET_FOLDER_PATH may be wrong.');

    // Try listing root to show what's there
    console.log('\n  Listing root drive children to help find the right path:');
    try {
      const children = await client.get(`${GRAPH_BASE}/sites/${SITE_ID}/drive/root/children`);
      for (const item of children.data.value || []) {
        const type = item.folder ? 'FOLDER' : 'FILE';
        console.log(`    [${type}] ${item.name}`);
      }
    } catch (e) {
      console.error('  Could not list root:', e.message);
    }
  }

  // 5. List active subscriptions
  console.log('\n=== 5. Active Subscriptions ===');
  try {
    const subs = await client.get(`${GRAPH_BASE}/subscriptions`);
    const subscriptions = subs.data.value || [];
    if (subscriptions.length === 0) {
      console.log('  NO SUBSCRIPTIONS FOUND - webhook was not created or expired');
    } else {
      for (const sub of subscriptions) {
        console.log('  Subscription:', sub.id);
        console.log('    Resource:', sub.resource);
        console.log('    ChangeType:', sub.changeType);
        console.log('    NotificationUrl:', sub.notificationUrl);
        console.log('    Expires:', sub.expirationDateTime);
        console.log('    ClientState:', sub.clientState ? '(set)' : '(not set)');
      }
    }
  } catch (err) {
    console.error('  SUBSCRIPTIONS FAILED:', err.response?.data?.error?.message || err.message);
  }

  // 6. Run a delta query to see recent changes
  console.log('\n=== 6. Recent Drive Changes (Delta Query) ===');
  try {
    const delta = await client.get(`${GRAPH_BASE}/sites/${SITE_ID}/drive/root/delta`);
    const items = delta.data.value || [];
    console.log(`  Found ${items.length} items in delta`);

    // Show a few items with their parentReference.path so user can see the format
    const sample = items.slice(0, 10);
    for (const item of sample) {
      const type = item.folder ? 'FOLDER' : item.file ? 'FILE' : 'OTHER';
      console.log(`    [${type}] ${item.name}`);
      console.log(`      parentReference.path: ${item.parentReference?.path || '(none)'}`);
    }
    if (items.length > 10) {
      console.log(`    ... and ${items.length - 10} more items`);
    }
  } catch (err) {
    console.error('  DELTA FAILED:', err.response?.data?.error?.message || err.message);
  }

  console.log('\n=== Diagnosis Complete ===\n');
}

diagnose().catch((err) => {
  console.error('Diagnosis failed:', err.message);
  process.exit(1);
});
