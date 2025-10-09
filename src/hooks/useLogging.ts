import { useState } from 'react';
import { LogEntry } from '../types';
import { ExcelServiceResult } from '../io-excel-service';

export function useLogging() {
  const [logs, setLogs] = useState<LogEntry[]>([]);
  const [isLoading, setIsLoading] = useState(false);

  const addLog = (type: LogEntry['type'], method: string, message: string, params?: any) => {
    const newLog: LogEntry = {
      id: Date.now().toString(),
      timestamp: new Date().toLocaleTimeString(),
      type,
      method,
      message,
      params
    };
    setLogs(prev => [newLog, ...prev.slice(0, 49)]); // Keep last 50 logs
  };

  const executeWithLogging = async (method: string, operation: () => Promise<ExcelServiceResult>, params?: any) => {
    setIsLoading(true);
    try {
      const result = await operation();

      // Check if the result indicates an error even though the promise resolved
      if (result && typeof result.success === 'boolean' && !result.success) {
        const errorMessage = result.message || 'Operation failed';
        addLog('error', method, errorMessage, { params, result });
        return result;
      }

      addLog('success', method, 'Operation completed successfully', { params, result });
      return result;
    } catch (error) {
      addLog('error', method, `Error: ${error}`, { params, error });
      throw error;
    } finally {
      setIsLoading(false);
    }
  };

  const clearLogs = () => setLogs([]);

  return {
    logs,
    isLoading,
    addLog,
    executeWithLogging,
    clearLogs
  };
}
