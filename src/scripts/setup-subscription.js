/**
 * Manually create a SharePoint webhook subscription.
 * Run with: npm run setup-subscription
 */
const config = require('../config');
const logger = require('../logger');
const sharepoint = require('../services/sharepoint');

async function main() {
  logger.info('Setting up SharePoint webhook subscription...');
  logger.info('Notification URL: ' + config.appUrl + '/api/webhook');

  try {
    const subscription = await sharepoint.createSubscription();
    logger.info('Subscription created successfully!');
    logger.info('Subscription ID: ' + subscription.id);
    logger.info('Expires: ' + subscription.expirationDateTime);
  } catch (err) {
    logger.error('Failed to create subscription', {
      error: err.message,
      response: err.response?.data,
    });
    process.exit(1);
  }

  process.exit(0);
}

main();
