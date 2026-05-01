const express = require('express');
const helmet = require('helmet');
const compression = require('compression');
const cron = require('node-cron');
const config = require('./config');
const logger = require('./logger');
const sharepoint = require('./services/sharepoint');
const bridge = require('./services/bridge');

const app = express();

// ── Middleware ────────────────────────────────────────────────

app.use(helmet());
app.use(compression());
app.use(express.json());

// Request logging
app.use((req, res, next) => {
  const start = Date.now();
  res.on('finish', () => {
    logger.info(`${req.method} ${req.path}`, {
      status: res.statusCode,
      durationMs: Date.now() - start,
    });
  });
  next();
});

// ── Health Check ─────────────────────────────────────────────

app.get('/health', (req, res) => {
  res.json({
    status: 'healthy',
    uptime: process.uptime(),
    subscription: sharepoint.getSubscriptionState(),
    timestamp: new Date().toISOString(),
  });
});

// List all Graph subscriptions visible to this app (debugging duplicates)
app.get('/api/subscriptions', async (req, res) => {
  try {
    const subs = await sharepoint.listSubscriptions();
    res.json({ count: subs.length, subscriptions: subs });
  } catch (err) {
    logger.error('Failed to list subscriptions', { error: err.message });
    res.status(500).json({ error: err.message });
  }
});

// ── SharePoint Webhook Endpoint ──────────────────────────────

// Debounce: collapse rapid-fire notifications into a single processing call
let debounceTimer = null;
const DEBOUNCE_MS = 5000; // wait 5 seconds after last notification before processing

// Serialize bridge runs. If a notification arrives while a pass is in-flight,
// queue exactly one follow-up pass to drain anything that arrived during it.
// Prevents overlapping LF folder lookups/creates from racing into 409s.
let bridgeRunning = false;
let bridgeRunPending = false;

async function runBridge() {
  if (bridgeRunning) {
    bridgeRunPending = true;
    return;
  }
  bridgeRunning = true;
  try {
    const result = await bridge.processNewItems();
    if (result.foldersProcessed > 0 || result.filesProcessed > 0
        || result.foldersFailed > 0 || result.filesFailed > 0) {
      logger.info('Webhook processing complete', result);
    }
  } catch (err) {
    logger.error('Error processing webhook notification', {
      error: err.message,
      stack: err.stack,
    });
  } finally {
    bridgeRunning = false;
    if (bridgeRunPending) {
      bridgeRunPending = false;
      runBridge();
    }
  }
}

app.post('/api/webhook', async (req, res) => {
  // Validation handshake: SharePoint/Graph sends validationToken on subscription creation
  if (req.query.validationToken) {
    logger.info('Webhook validation handshake received');
    res.set('Content-Type', 'text/plain');
    return res.status(200).send(req.query.validationToken);
  }

  // Also handle lowercase variant (SharePoint REST API)
  if (req.query.validationtoken) {
    logger.info('Webhook validation handshake received (lowercase)');
    res.set('Content-Type', 'text/plain');
    return res.status(200).send(req.query.validationtoken);
  }

  // Validate clientState to verify notification is from Microsoft
  const notifications = req.body.value || [];
  for (const n of notifications) {
    if (n.clientState !== config.microsoft.webhookClientState) {
      logger.warn('Invalid clientState in webhook notification');
      return res.status(403).json({ error: 'Invalid client state' });
    }
  }

  // Respond immediately (must respond within 10 seconds)
  res.status(202).json({ status: 'accepted' });

  // Debounce: reset timer on each notification, only process once things settle
  if (debounceTimer) clearTimeout(debounceTimer);
  debounceTimer = setTimeout(() => {
    debounceTimer = null;
    logger.debug('Processing debounced webhook notification');
    runBridge();
  }, DEBOUNCE_MS);
});

// ── Scheduled Tasks ──────────────────────────────────────────

// Reconcile webhook subscription daily at 2 AM UTC (idempotent)
cron.schedule('0 2 * * *', async () => {
  logger.info('Running scheduled subscription reconciliation');
  try {
    await sharepoint.ensureSubscription();
    logger.info('Scheduled reconciliation completed');
  } catch (err) {
    logger.error('Scheduled reconciliation failed', { error: err.message });
  }
});

// ── Server Startup ───────────────────────────────────────────

const server = app.listen(config.port, '0.0.0.0', async () => {
  logger.info(`Server started on port ${config.port}`);

  // Reconcile webhook subscription on startup (adopts existing or creates new)
  try {
    await sharepoint.ensureSubscription();
    logger.info('Subscription reconciled on startup', {
      state: sharepoint.getSubscriptionState(),
    });
  } catch (err) {
    logger.error('Failed to reconcile subscription on startup', {
      error: err.message,
      stack: err.stack,
    });
    logger.info('The app will keep running. Use /api/setup to retry, or the cron will retry at 2 AM UTC.');
  }
});

// ── Manual Setup Endpoint ────────────────────────────────────

app.post('/api/setup', async (req, res) => {
  try {
    const sub = await sharepoint.ensureSubscription();
    res.json({ status: 'ok', subscription: sub });
  } catch (err) {
    logger.error('Manual setup failed', { error: err.message });
    res.status(500).json({ error: err.message });
  }
});

// ── Graceful Shutdown ────────────────────────────────────────

process.on('SIGTERM', () => {
  logger.info('SIGTERM received, shutting down gracefully');
  server.close(() => {
    logger.info('Server closed');
    process.exit(0);
  });
  setTimeout(() => {
    logger.error('Forced shutdown after timeout');
    process.exit(1);
  }, 14000);
});

process.on('uncaughtException', (err) => {
  logger.error('Uncaught exception', { error: err.message, stack: err.stack });
  process.exit(1);
});

process.on('unhandledRejection', (reason) => {
  logger.error('Unhandled rejection', { reason: String(reason) });
  process.exit(1);
});
