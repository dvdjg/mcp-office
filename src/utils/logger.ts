/**
 * @file Logger configuration for the MS Office MCP server.
 * @author David Jurado
 * @date 2025-05-03
 * @copyright Copyright (c) 2025 David Jurado
 * @license MIT
 */
import winston, { Logger } from 'winston';
// Assuming fastmcp is an available module and exports a logger named 'log'
// If not, this import will fail and need adjustment.
// import { log as fastmcpLog } from 'fastmcp'; // Uncomment if fastmcp is available

// Define the environment variable to select the logger
const loggerType = process.env.MCP_LOGGER_TYPE;

let logger: Logger; // Add explicit type annotation

if (loggerType === 'fastmcp') {
  // Use the fastmcp logger if the environment variable indicates
  // logger = fastmcpLog; // Uncomment if fastmcp is available
  // Fallback to Winston if fastmcp is not ready or cannot be imported
  logger = createWinstonLogger();

} else {
  // Use the default Winston logger
  logger = createWinstonLogger();
}

/**
 * Creates and configures the Winston logger instance.
 * @returns {Logger} The configured Winston logger instance.
 */
function createWinstonLogger(): Logger { // Add return type annotation
  const winstonLogger = winston.createLogger({
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
    winstonLogger.add(new winston.transports.Console({
      format: winston.format.combine(
        winston.format.colorize(),
        winston.format.simple() // Simple format for console
      ),
    }));
  }
  return winstonLogger;
}


// Export the selected logger
export default logger;