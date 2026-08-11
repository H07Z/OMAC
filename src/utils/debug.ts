/**
 * Debug logging utility
 * Controlled by DEBUG flag, logs to console and optionally to in-app debug panel
 */

const DEBUG = true;

interface LogMessage {
  timestamp: Date;
  category: string;
  message: string;
  data?: unknown;
}

const logs: LogMessage[] = [];

export function debugLog(category: string, message: string, data?: unknown): void {
  if (!DEBUG) return;
  
  const logEntry: LogMessage = {
    timestamp: new Date(),
    category,
    message,
    data
  };
  
  logs.push(logEntry);
  
  const prefix = `[${category}]`;
  if (data !== undefined) {
    console.log(prefix, message, data);
  } else {
    console.log(prefix, message);
  }
}

export function getDebugLogs(): LogMessage[] {
  return [...logs];
}

export function clearDebugLogs(): void {
  logs.length = 0;
}

export function formatDebugLog(log: LogMessage): string {
  const time = log.timestamp.toLocaleTimeString();
  return `[${time}] [${log.category}] ${log.message}${log.data ? ': ' + JSON.stringify(log.data) : ''}`;
}
