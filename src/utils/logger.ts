/**
 * @file Configures the application logger using Winston.
 */
import winston from 'winston';

const logger = winston.createLogger({
  level: process.env.LOG_LEVEL || 'debug',
  format: winston.format.combine(
    winston.format.timestamp({ format: 'YYYY-MM-DD HH:mm:ss' }),
    winston.format.errors({ stack: true }), // Log stack traces
    winston.format.splat(),
    winston.format.json() // Log in JSON format
  ),
  defaultMeta: { service: 'msoffice-mcp' },
  transports: [
    // Log errors to a separate file
    new winston.transports.File({ filename: 'logs/error.log', level: 'error' }),
    // Log all levels to a combined file
    new winston.transports.File({ filename: 'logs/combined.log' }),
  ],
});

// If not in production AND not in STDIO mode, also log to the console with colorized output
// Check if OFFICE_MCP_PORT is defined to infer if it's running in SSE mode (not STDIO)
if (process.env.NODE_ENV !== 'production' && process.env.OFFICE_MCP_PORT) {
  logger.add(new winston.transports.Console({
    format: winston.format.combine(
      winston.format.colorize(),
      winston.format.simple() // Simple format for console
    ),
  }));
}

export default logger;