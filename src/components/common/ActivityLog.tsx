import { RefreshCw, CheckCircle, AlertCircle, Info, Activity } from 'lucide-react';
import { LogEntry } from '../../types';

interface ActivityLogProps {
  logs: LogEntry[];
  isLoading: boolean;
  onClearLogs: () => void;
}

export function ActivityLog({ logs, isLoading, onClearLogs }: ActivityLogProps) {
  const getLogIcon = (type: LogEntry['type']) => {
    switch (type) {
      case 'success':
        return <CheckCircle className="w-4 h-4 text-green-600" />;
      case 'error':
        return <AlertCircle className="w-4 h-4 text-red-600" />;
      case 'info':
        return <Info className="w-4 h-4 text-blue-600" />;
    }
  };

  const getLogTextColor = (type: LogEntry['type']) => {
    switch (type) {
      case 'success':
        return 'text-green-700';
      case 'error':
        return 'text-red-700';
      case 'info':
        return 'text-blue-700';
    }
  };

  return (
    <div className="bg-white rounded-xl shadow-sm border border-gray-200 overflow-hidden flex flex-col max-h-[calc(100vh-6rem)]">
      <div className="p-4 border-b border-gray-200 flex items-center justify-between flex-shrink-0">
        <div className="flex items-center space-x-2">
          <Activity className="w-5 h-5 text-gray-600" />
          <h2 className="text-lg font-semibold text-gray-900">Activity Log</h2>
        </div>
        <button
          onClick={onClearLogs}
          className="p-1 text-gray-500 hover:text-gray-700 transition-colors"
          title="Clear logs"
        >
          <RefreshCw className="w-4 h-4" />
        </button>
      </div>

      <div className="flex-1 overflow-y-auto min-h-0">
          {isLoading && (
            <div className="p-4 border-b border-gray-100 bg-yellow-50">
              <div className="flex items-center space-x-2">
                <div className="animate-spin rounded-full h-4 w-4 border-b-2 border-yellow-600"></div>
                <span className="text-sm text-yellow-700">Executing operation...</span>
              </div>
            </div>
          )}

          {logs.length === 0 && !isLoading ? (
            <div className="p-8 text-center text-gray-500">
              <Activity className="w-8 h-8 mx-auto mb-2 text-gray-300" />
              <p className="text-sm">No activity yet</p>
            </div>
          ) : (
            <div className="divide-y divide-gray-100">
              {logs.map((log) => (
                <div key={log.id} className="p-3 hover:bg-gray-50">
                  <div className="flex items-start space-x-2">
                    {getLogIcon(log.type)}
                    <div className="flex-1 min-w-0">
                      <div className="flex items-center space-x-2 mb-1">
                        <span className="text-xs font-medium text-gray-900">{log.method}</span>
                        <span className="text-xs text-gray-500">{log.timestamp}</span>
                      </div>
                      <p className={`text-sm ${getLogTextColor(log.type)} break-words`}>
                        {log.message}
                      </p>
                      {log.params && (
                        <details className="mt-1">
                          <summary className="text-xs text-gray-500 cursor-pointer hover:text-gray-700">
                            View details
                          </summary>
                          <pre className="text-xs bg-gray-100 p-2 rounded mt-1 overflow-x-auto">
                            {JSON.stringify(log.params, null, 2)}
                          </pre>
                        </details>
                      )}
                    </div>
                  </div>
                </div>
              ))}
            </div>
          )}
        </div>
    </div>
  );
}
