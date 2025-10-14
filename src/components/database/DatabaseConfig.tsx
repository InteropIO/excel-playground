import { Database } from 'lucide-react';
import { DataSource, ColumnType } from '../../io-excel-service';

interface DatabaseConfigProps {
  dataSource: DataSource;
  onDataSourceChange: (dataSource: DataSource) => void;
}

export function DatabaseConfig({ dataSource, onDataSourceChange }: DatabaseConfigProps) {
  const updateDataSource = (updates: Partial<DataSource>) => {
    onDataSourceChange({ ...dataSource, ...updates });
  };

  return (
    <div className="sticky top-4 z-10 bg-white rounded-lg shadow-md border border-gray-200 p-4">
      <h3 className="text-md font-semibold text-gray-900 mb-3 flex items-center">
        <Database className="w-4 h-4 mr-2" />
        Database Configuration
      </h3>

      {/* Primary Configuration - Always Visible */}
      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-3 mb-3">
        <div>
          <label className="block text-xs font-medium text-gray-600 mb-1">Table Name</label>
          <input
            type="text"
            value={dataSource.name}
            onChange={(e) => updateDataSource({ name: e.target.value })}
            className="w-full px-2 py-1 text-sm border border-gray-300 rounded focus:outline-none focus:ring-1 focus:ring-blue-500 focus:border-transparent"
          />
        </div>
        <div>
          <label className="block text-xs font-medium text-gray-600 mb-1">Primary Key</label>
          <input
            type="text"
            value={dataSource.primaryKey?.join(', ') || ''}
            onChange={(e) => updateDataSource({ primaryKey: e.target.value.split(',').map(k => k.trim()) })}
            className="w-full px-2 py-1 text-sm border border-gray-300 rounded focus:outline-none focus:ring-1 focus:ring-blue-500 focus:border-transparent"
            placeholder="ID, Name"
          />
        </div>
        <div>
          <label className="block text-xs font-medium text-gray-600 mb-1">Columns</label>
          <input
            type="text"
            value={dataSource.columns?.map(c => c.name).join(', ') || ''}
            readOnly
            className="w-full px-2 py-1 text-sm border border-gray-300 rounded bg-gray-50 text-gray-600"
            placeholder="Column names (view only)"
          />
        </div>
        <div>
          <label className="block text-xs font-medium text-gray-600 mb-1">Data Rows</label>
          <input
            type="text"
            value={`${dataSource.data?.length || 0} rows`}
            readOnly
            className="w-full px-2 py-1 text-sm border border-gray-300 rounded bg-gray-50 text-gray-600"
          />
        </div>
      </div>

      {/* Collapsible Advanced Configuration */}
      <details className="group">
        <summary className="cursor-pointer text-xs text-gray-600 hover:text-gray-800 flex items-center space-x-1">
          <span>Column Configuration & Data Management</span>
          <svg className="w-3 h-3 transform group-open:rotate-180 transition-transform" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 9l-7 7-7-7" />
          </svg>
        </summary>

        <div className="mt-3 space-y-3">
          {/* Column Configuration */}
          <div>
            <label className="block text-xs font-medium text-gray-600 mb-2">Columns Schema</label>
            <div className="space-y-2">
              {dataSource.columns.map((column, index) => (
                <div key={index} className="grid grid-cols-2 md:grid-cols-5 gap-2 p-2 bg-gray-50 rounded">
                  <input
                    type="text"
                    value={column.name}
                    onChange={(e) => {
                      const newColumns = [...dataSource.columns];
                      newColumns[index] = { ...newColumns[index], name: e.target.value };
                      updateDataSource({ columns: newColumns });
                    }}
                    className="px-2 py-1 text-xs border border-gray-300 rounded focus:outline-none focus:ring-1 focus:ring-blue-500"
                    placeholder="Column name"
                  />
                  <select
                    value={column.type}
                    onChange={(e) => {
                      const newColumns = [...dataSource.columns];
                      newColumns[index] = { ...newColumns[index], type: e.target.value as ColumnType };
                      updateDataSource({ columns: newColumns });
                    }}
                    className="px-2 py-1 text-xs border border-gray-300 rounded focus:outline-none focus:ring-1 focus:ring-blue-500"
                  >
                    <option value={ColumnType.Integer}>Integer</option>
                    <option value={ColumnType.Text}>Text</option>
                    <option value={ColumnType.Boolean}>Boolean</option>
                    <option value={ColumnType.DateTime}>DateTime</option>
                    <option value={ColumnType.Float}>Float</option>
                    <option value={ColumnType.Decimal}>Decimal</option>
                  </select>
                  <label className="flex items-center text-xs">
                    <input
                      type="checkbox"
                      checked={column.pk}
                      onChange={(e) => {
                        const newColumns = [...dataSource.columns];
                        newColumns[index] = { ...newColumns[index], pk: e.target.checked };
                        updateDataSource({ columns: newColumns });
                      }}
                      className="mr-1"
                    />
                    PK
                  </label>
                  <label className="flex items-center text-xs">
                    <input
                      type="checkbox"
                      checked={column.autoIncrement}
                      onChange={(e) => {
                        const newColumns = [...dataSource.columns];
                        newColumns[index] = { ...newColumns[index], autoIncrement: e.target.checked };
                        updateDataSource({ columns: newColumns });
                      }}
                      className="mr-1"
                    />
                    Auto
                  </label>
                  <label className="flex items-center text-xs">
                    <input
                      type="checkbox"
                      checked={!column.nullable}
                      onChange={(e) => {
                        const newColumns = [...dataSource.columns];
                        newColumns[index] = { ...newColumns[index], nullable: !e.target.checked };
                        updateDataSource({ columns: newColumns });
                      }}
                      className="mr-1"
                    />
                    Required
                  </label>
                </div>
              ))}
            </div>
          </div>

          {/* Sample Data Preview */}
          <div>
            <label className="block text-xs font-medium text-gray-600 mb-2">Sample Data (First 3 rows)</label>
            <div className="bg-gray-50 rounded p-2 text-xs font-mono">
              {dataSource.data?.slice(0, 3).map((row, index) => (
                <div key={index} className="mb-1">
                  [{row.map(cell => cell === null ? 'null' : `"${cell}"`).join(', ')}]
                </div>
              ))}
              {(dataSource.data?.length || 0) > 3 && (
                <div className="text-gray-500">... and {(dataSource.data?.length || 0) - 3} more rows</div>
              )}
            </div>
          </div>
        </div>
      </details>
    </div>
  );
}
