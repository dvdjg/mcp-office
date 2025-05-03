/**
 * @file Configures the application logger using Winston or FastMCP based on environment variable.
 */
import winston, { Logger } from 'winston';
// Asumimos que fastmcp es un módulo disponible y exporta un logger llamado 'log'
// Si no, esta importación fallará y habrá que ajustarla.
// import { log as fastmcpLog } from 'fastmcp'; // Descomentar si fastmcp está disponible

// Define la variable de entorno para seleccionar el logger
const loggerType = process.env.MCP_LOGGER_TYPE;

let logger: Logger; // Añadir anotación de tipo explícita

if (loggerType === 'fastmcp') {
  // Usar el logger de fastmcp si la variable de entorno lo indica
  // logger = fastmcpLog; // Descomentar si fastmcp está disponible
  // Temporalmente, si fastmcp no está disponible, usamos el logger de Winston
  console.warn("MCP_LOGGER_TYPE is set to 'fastmcp', but fastmcp logger is not fully integrated yet. Using default Winston logger.");
  // Fallback a Winston si fastmcp no está listo o no se puede importar
  logger = createWinstonLogger();

} else {
  // Usar el logger de Winston por defecto
  logger = createWinstonLogger();
}

// Función para crear y configurar el logger de Winston
function createWinstonLogger(): Logger { // Añadir anotación de tipo de retorno
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

  // Si no en producción Y no en modo STDIO, también log a la consola con salida colorizada
  // Verificar si OFFICE_MCP_PORT está definido para inferir si se está ejecutando en modo SSE (no STDIO)
  if (process.env.NODE_ENV !== 'production' && process.env.OFFICE_MCP_PORT) {
    winstonLogger.add(new winston.transports.Console({
      format: winston.format.combine(
        winston.format.colorize(),
        winston.format.simple() // Formato simple para consola
      ),
    }));
  }
  return winstonLogger;
}


// La lógica de fallback `const log = context?.log ?? logger;` se aplicará donde se use este logger.
// Exportar el logger seleccionado
export default logger;