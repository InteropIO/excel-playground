import { useState } from 'react';
import { FileSpreadsheet, RefreshCw } from 'lucide-react';
import { ColumnType, DataSource } from './io-excel-service';
import { TabNavigation } from './components/common/TabNavigation';
import { ActivityLog } from './components/common/ActivityLog';
import { DatabaseSection } from './components/database/DatabaseSection';
import { ExcelSection } from './components/excel/ExcelSection';
import { useLogging } from './hooks/useLogging';
import { useDatabaseOperations } from './hooks/useDatabaseOperations';
import { useExcelOperations } from './hooks/useExcelOperations';
import { createDatabaseSnippets, createExcelSnippets, groupSnippetsByCategory } from './utils/snippets';
import { ExcelState, DatabaseState } from './types';

function App() {
  const [activeTab, setActiveTab] = useState<'database' | 'excel'>('excel');
  const { logs, isLoading, addLog, executeWithLogging, clearLogs } = useLogging();

  // Database state
  const [databaseState, setDatabaseState] = useState<DatabaseState>({
    dataSource: {
      name: 'UserTable',
      columns: [
        { name: 'ID', type: ColumnType.Integer, primaryKey: true, autoIncrement: true, nullable: false },
        { name: 'Name', type: ColumnType.Text, primaryKey: false, autoIncrement: false, nullable: false },
        { name: 'Email', type: ColumnType.Text, primaryKey: false, autoIncrement: false, nullable: true }
      ],
      primaryKey: ['ID'],
      data: [
        [null, 'John Doe', 'johndoe@example.com'],
        [null, 'Jane Smith', 'janesmith@example.com'],
        [null, 'Sam Wilson', 'samwilson@example.com']
      ]
    },
    queryText: 'SELECT * FROM UserTable',
    queryResults: []
  });

  // Excel state
  const [excelState, setExcelState] = useState<ExcelState>({
    workbookName: 'Book3',
    worksheetName: 'Sheet1',
    rangeValue: 'A1:C10',
    cellValue: 'Hello World',
    tableName: 'DemoTable',
    contextMenuCaption: 'Send Data',
    ribbonMenuCaption: 'Process Data',
    commentText: 'This is a demo comment',
    backgroundColor: '#FFE4B5',
    foregroundColor: '#000000',
    fileName: 'exported_file.xlsx',
    xlReference: 'Sheet1!A1:B5',
    subscriptionId: 'sub_123',
    menuId: 'menu_123',
    fromRow: 1,
    rowsToRead: 10,
    rowPosition: 1
  });

  // Initialize operations
  const dbOps = useDatabaseOperations();
  const { createOperations } = useExcelOperations();

  // Create Excel operations with current state - this will be recreated when state changes
  const xlOps = createOperations(excelState, databaseState.dataSource, addLog);

  // Update functions
  const updateDatabaseState = (updates: Partial<DatabaseState>) => {
    setDatabaseState(prev => ({ ...prev, ...updates }));
  };

  const updateDataSource = (dataSource: DataSource) => {
    updateDatabaseState({ dataSource });
  };

  const updateExcelState = (updates: Partial<ExcelState>) => {
    setExcelState(prev => ({ ...prev, ...updates }));
  };

  // Database operations with logging
  const dbOperations = [
    () => executeWithLogging('DB.Init', () => dbOps.initDatabase(databaseState.dataSource), databaseState.dataSource),
    () => executeWithLogging('DB.CreateTable', () => dbOps.createTable(databaseState.dataSource), databaseState.dataSource),
    () => executeWithLogging('DB.InsertData', () => dbOps.insertData(databaseState.dataSource), databaseState.dataSource),
    async () => {
      const result = await executeWithLogging('DB.ExecuteQuery', () => dbOps.executeQuery(databaseState.dataSource, databaseState.queryText), { query: databaseState.queryText });
      if (result?.data) {
        updateDatabaseState({ queryResults: result.data });
      }
    },
    () => executeWithLogging('DB.UpdateRow', () => dbOps.updateRow(databaseState.dataSource, ['Updated Name', 'updated@example.com'], 1), { rowData: ['Updated Name', 'updated@example.com'], primaryKeyValue: 1 }),
    () => executeWithLogging('DB.UpdateColumns', () => dbOps.updateColumns(databaseState.dataSource, { Name: 'Updated Name', Email: 'updated@example.com' }, 1), { updates: { Name: 'Updated Name', Email: 'updated@example.com' }, primaryKeyValue: 1 }),
    () => executeWithLogging('DB.Dispose', () => dbOps.disposeDatabase(databaseState.dataSource), databaseState.dataSource)
  ];

  // Excel operations grouped by category
  const excelOperations = {
    'Basic': [
      () => executeWithLogging('XL.CreateWorkbook', xlOps.createWorkbook, { workbookName: excelState.workbookName, worksheetName: excelState.worksheetName }),
      () => executeWithLogging('XL.OpenWorkbook', xlOps.openWorkbook, { fileName: excelState.fileName }),
      () => executeWithLogging('XL.SaveAs', xlOps.saveWorkbook, { fileName: excelState.fileName }),
      () => executeWithLogging('XL.Activate', xlOps.activateRange, { range: excelState.rangeValue })
    ],
    'Read/Write': [
      () => executeWithLogging('XL.Read', xlOps.readRange, { workbook: excelState.workbookName, worksheet: excelState.worksheetName, range: excelState.rangeValue }),
      () => executeWithLogging('XL.Write', xlOps.writeRange, { workbook: excelState.workbookName, worksheet: excelState.worksheetName, range: excelState.rangeValue, value: excelState.cellValue }),
      () => executeWithLogging('XL.ReadXlRef', xlOps.readExcelRef, { reference: excelState.xlReference }),
      () => executeWithLogging('XL.WriteXlRef', xlOps.writeExcelRef, { reference: excelState.xlReference, value: excelState.cellValue })
    ],
    'Subscriptions': [
      () => executeWithLogging('XL.Subscribe', xlOps.subscribeToRange, { range: excelState.rangeValue }),
      () => executeWithLogging('XL.SubscribeDeltas', xlOps.subscribeDeltas, { range: excelState.rangeValue }),
      () => executeWithLogging('XL.DestroySubscription', xlOps.destroySubscription, { subscriptionId: excelState.subscriptionId })
    ],
    'Tables': [
      () => executeWithLogging('XL.CreateTable', xlOps.createExcelTable, { tableName: excelState.tableName, range: excelState.rangeValue }),
      () => executeWithLogging('XL.CreateLinkedTable', xlOps.createLinkedTable, { range: excelState.rangeValue, dataSource: databaseState.dataSource }),
      () => executeWithLogging('XL.RefreshTable', xlOps.refreshTable, { range: excelState.rangeValue, tableName: excelState.tableName }),
      () => executeWithLogging('XL.WriteTableRows', xlOps.writeTableRows, { tableName: excelState.tableName, rowPosition: excelState.rowPosition, data: [['3', 'New User', 'newuser@example.com']] }),
      () => executeWithLogging('XL.ReadTableRows', xlOps.readTableRows, { tableName: excelState.tableName, fromRow: excelState.fromRow, rowsToRead: excelState.rowsToRead }),
      () => executeWithLogging('XL.UpdateTableColumns', xlOps.updateTableColumns, { tableName: excelState.tableName, columnOperations: [{ currentName: 'Email', newName: 'EmailAddress', position: null, operation: 'Rename' }] }),
      () => executeWithLogging('XL.DescribeTableColumns', xlOps.describeTableColumns, { tableName: excelState.tableName })
    ],
    'Menus': [
      () => executeWithLogging('XL.CreateContextMenu', xlOps.createContextMenu, { caption: excelState.contextMenuCaption, range: excelState.rangeValue }),
      () => executeWithLogging('XL.CreateContextMenuRaw', xlOps.createContextMenuRaw, { caption: excelState.contextMenuCaption, range: excelState.rangeValue }),
      () => executeWithLogging('XL.DestroyContextMenu', xlOps.destroyContextMenu, { menuId: excelState.menuId }),
      () => executeWithLogging('XL.CreateDynamicRibbonMenu', xlOps.createRibbonMenu, { caption: excelState.ribbonMenuCaption, range: excelState.rangeValue }),
      () => executeWithLogging('XL.CreateDynamicRibbonMenuRaw', xlOps.createRibbonMenuRaw, { caption: excelState.ribbonMenuCaption, range: excelState.rangeValue }),
      () => executeWithLogging('XL.DestroyRibbonMenu', xlOps.destroyRibbonMenu, { menuId: excelState.menuId })
    ],
    'Styling': [
      () => executeWithLogging('XL.WriteComment', xlOps.writeComment, { range: excelState.rangeValue, comment: excelState.commentText }),
      () => executeWithLogging('XL.ClearComments', xlOps.clearComments, { range: excelState.rangeValue }),
      () => executeWithLogging('XL.ClearContents', xlOps.clearContents, { range: excelState.rangeValue }),
      () => executeWithLogging('XL.ApplyStyles', xlOps.applyStyles, { range: excelState.rangeValue, backgroundColor: excelState.backgroundColor, foregroundColor: excelState.foregroundColor })
    ]
  };

  // Execute operations
  const executeDatabaseOperation = (index: number) => {
    if (dbOperations[index]) {
      dbOperations[index]();
    }
  };

  const executeExcelOperation = (index: number, category: string) => {
    const categoryOps = excelOperations[category as keyof typeof excelOperations];
    if (categoryOps && categoryOps[index]) {
      categoryOps[index]();
    }
  };

  // Generate snippets
  const dbSnippets = createDatabaseSnippets(databaseState);
  const xlSnippets = createExcelSnippets(excelState);
  const groupedXlSnippets = groupSnippetsByCategory(xlSnippets);

  return (
    <div className="min-h-screen bg-gradient-to-br from-slate-50 via-blue-50 to-indigo-100">
      {/* Header */}
      <header className="bg-white shadow-sm border-b border-gray-200">
        <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8">
          <div className="flex justify-between items-center h-16">
            <div className="flex items-center space-x-3">
              <div className="bg-gradient-to-r from-blue-600 to-indigo-600 p-2 rounded-lg">
                <FileSpreadsheet className="w-6 h-6 text-white" />
              </div>
              <div>
                <h1 className="text-xl font-bold text-gray-900">Complete IO Excel Service Guide</h1>
                <p className="text-sm text-gray-500">Comprehensive API Documentation & Testing Tool - All Methods Included</p>
              </div>
            </div>
            <div className="flex items-center space-x-4">
              {isLoading && (
                <div className="flex items-center space-x-2 text-blue-600">
                  <RefreshCw className="w-4 h-4 animate-spin" />
                  <span className="text-sm">Processing...</span>
                </div>
              )}
              <button
                onClick={clearLogs}
                className="px-4 py-2 text-sm bg-gray-100 hover:bg-gray-200 text-gray-700 rounded-lg transition-colors duration-200"
              >
                Clear Logs
              </button>
            </div>
          </div>
        </div>
      </header>

      <div className="max-w-7xl mx-auto px-4 sm:px-6 lg:px-8 py-8">
        <div className="grid grid-cols-1 lg:grid-cols-4 gap-8">
          {/* Main Content */}
          <div className="lg:col-span-3 space-y-6">
            <div className="bg-white rounded-xl shadow-sm border border-gray-200">
              <div className="border-b border-gray-200">
                <TabNavigation activeTab={activeTab} onTabChange={setActiveTab} />
              </div>

              <div className="p-6">
                {activeTab === 'database' && (
                  <DatabaseSection
                    dataSource={databaseState.dataSource}
                    onDataSourceChange={updateDataSource}
                    queryText={databaseState.queryText}
                    onQueryTextChange={(text) => updateDatabaseState({ queryText: text })}
                    queryResults={databaseState.queryResults}
                    snippets={dbSnippets}
                    onExecuteSnippet={executeDatabaseOperation}
                  />
                )}

                {activeTab === 'excel' && (
                  <ExcelSection
                    state={excelState}
                    onStateChange={updateExcelState}
                    groupedSnippets={groupedXlSnippets}
                    onExecuteSnippet={executeExcelOperation}
                  />
                )}
              </div>
            </div>
          </div>

          {/* Activity Log */}
          <div className="sticky top-4 self-start z-10">
            <ActivityLog
              logs={logs}
              isLoading={isLoading}
              onClearLogs={clearLogs}
            />
          </div>
        </div>
      </div>
    </div>
  );
}

export default App;
