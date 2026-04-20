/**
 * Manually renew SharePoint webhook subscription.
 * Run with: npm run renew-webhooks
 * Also used by Heroku Scheduler if configured.
 */
const logger = require('../logger');
const sharepoint = require('../services/sharepoint');

async function main() {
  logger.info('Renewing SharePoint webhook subscription...');

  try {
    const subscription = await sharepoint.renewSubscription();
    logger.info('Subscription renewed successfully!');
    logger.info('New expiration: ' + subscription.expirationDateTime);
  } catch (err) {
    logger.error('Failed to renew subscription', {
      error: err.message,
      response: err.response?.data,
    });
    process.exit(1);
  }

  process.exit(0);
}

main();
