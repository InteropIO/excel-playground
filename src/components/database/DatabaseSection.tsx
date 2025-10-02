import { DatabaseConfig } from './DatabaseConfig';
import { CodeBlock } from '../common/CodeBlock';
import { DataSource } from '../../io-excel-service';
import { CodeSnippet } from '../../types';

interface DatabaseSectionProps {
  dataSource: DataSource;
  onDataSourceChange: (dataSource: DataSource) => void;
  queryText: string;
  onQueryTextChange: (text: string) => void;
  queryResults: any[];
  snippets: CodeSnippet[];
  onExecuteSnippet: (index: number) => void;
}

export function DatabaseSection({
  dataSource,
  onDataSourceChange,
  queryText,
  onQueryTextChange,
  queryResults,
  snippets,
  onExecuteSnippet
}: DatabaseSectionProps) {
  return (
    <div className="space-y-8">
      <DatabaseConfig 
        dataSource={dataSource} 
        onDataSourceChange={onDataSourceChange} 
      />

      {/* Database Operations */}
      <div className="space-y-6">
        {snippets.map((snippet, index) => (
          <CodeBlock
            key={index}
            snippet={snippet}
            onExecute={() => onExecuteSnippet(index)}
          />
        ))}
      </div>

      {/* Query Section */}
      <div className="bg-gray-50 rounded-lg p-4">
        <textarea
          value={queryText}
          onChange={(e) => onQueryTextChange(e.target.value)}
          rows={3}
          className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent font-mono text-sm"
          placeholder="Enter your SQL query here..."
        />
      </div>

      {/* Query Results */}
      {queryResults.length > 0 && (
        <div className="bg-white border border-gray-200 rounded-lg overflow-hidden">
          <div className="p-4 border-b border-gray-200">
            <h3 className="font-medium text-gray-900">Query Results</h3>
          </div>
          <div className="overflow-x-auto">
            <table className="min-w-full divide-y divide-gray-200">
              <tbody className="bg-white divide-y divide-gray-200">
                {queryResults.map((row, index) => (
                  <tr key={index}>
                    {Object.values(row).map((cell: any, cellIndex) => (
                      <td key={cellIndex} className="px-6 py-3 whitespace-nowrap text-sm text-gray-900">
                        {cell === null ? 'null' : String(cell)}
                      </td>
                    ))}
                  </tr>
                ))}
              </tbody>
            </table>
          </div>
        </div>
      )}
    </div>
  );
}
