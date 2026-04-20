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

module.exports = logger;
