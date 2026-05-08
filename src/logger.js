const winston = require('winston');

const logger = winston.createLogger({
  level: process.env.LOG_LEVEL || 'info',
  format: winston.format.combine(
    winston.format.timestamp(),
    winston.format.errors({ stack: true }),
    winston.format.json()
  ),
  defaultMeta: {
    service: 'od2lf',
    dyno: process.env.DYNO || 'local',
  },
  transports: [
    new winston.transports.Console({
      format: winston.format.combine(
        winston.format.colorize(),
        winston.format.printf(({ level, message, timestamp, ...meta }) => {
          const { service, dyno, ...rest } = meta;
          const extra = Object.keys(rest).length > 0
            ? ' ' + JSON.stringify(rest)
            : '';
          return `${timestamp} [${level}] ${message}${extra}`;
        })
      ),
    }),
  ],
});

// Summarize an HTTP error body for logging. IIS error pages are full HTML
// documents that bloat logs and obscure the underlying problem; collapse them
// to a one-line summary while leaving JSON/short text bodies intact.
function summarizeBody(body) {
  if (body === undefined || body === null) return undefined;
  if (typeof body === 'object') return body;
  if (typeof body !== 'string') return body;
  const trimmed = body.trim();
  if (trimmed.startsWith('<')) {
    const titleMatch = trimmed.match(/<title>([^<]+)<\/title>/i);
    if (titleMatch) return `[HTML: ${titleMatch[1].trim()}]`;
    const h3Match = trimmed.match(/<h3>([^<]+)<\/h3>/i);
    if (h3Match) return `[HTML: ${h3Match[1].trim()}]`;
    return `[HTML response, ${trimmed.length} chars]`;
  }
  return trimmed.length > 500 ? trimmed.slice(0, 500) + '…' : trimmed;
}

module.exports = logger;
module.exports.summarizeBody = summarizeBody;
