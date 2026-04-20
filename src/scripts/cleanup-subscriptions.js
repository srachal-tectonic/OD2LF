/**
 * Delete ALL existing Microsoft Graph webhook subscriptions.
 * Run with: node src/scripts/cleanup-subscriptions.js
 */
require('dotenv').config();
require('../config');
const { graphClient } = require('../services/graph-auth');

const GRAPH_BASE = 'https://graph.microsoft.com/v1.0';

async function main() {
  const client = graphClient();

  console.log('Fetching all subscriptions...');
  const response = await client.get(`${GRAPH_BASE}/subscriptions`);
  const subs = response.data.value || [];

  if (subs.length === 0) {
    console.log('No subscriptions found.');
    return;
  }

  console.log(`Found ${subs.length} subscription(s):\n`);
  for (const sub of subs) {
    console.log(`  ID: ${sub.id}`);
    console.log(`  URL: ${sub.notificationUrl}`);
    console.log(`  Expires: ${sub.expirationDateTime}`);

    try {
      await client.delete(`${GRAPH_BASE}/subscriptions/${sub.id}`);
      console.log(`  DELETED\n`);
    } catch (err) {
      console.error(`  DELETE FAILED: ${err.response?.data?.error?.message || err.message}\n`);
    }
  }

  console.log('Cleanup complete.');
}

main().catch((err) => {
  console.error('Failed:', err.message);
  process.exit(1);
});
