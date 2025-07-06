import { useState, useRef, useContext } from 'react';
import { IOConnectContext } from "@interopio/react-hooks";
import { GlueDBService, GlueExcelService, DataSource, ColumnType } from './io-excel-service';
import {
  Database,
  FileSpreadsheet,
  Play,
  Edit3,
  Settings,
  Menu,
  RefreshCw,
  Palette,
  Activity,
  CheckCircle,
  AlertCircle,
  Info,
  Code,
  Copy,
  Check,
  Table,
  Zap,
} from 'lucide-react';

interface LogEntry {
  id: string;
  timestamp: string;
  type: 'success' | 'error' | 'info';
  method: string;
  message: string;
  params?: any;
}

interface CodeSnippet {
  title: string;
  code: string;
  description: string;
  category: string;
}

function CodeBlock({ snippet, onExecute }: { snippet: CodeSnippet; onExecute: () => void }) {
  const [copied, setCopied] = useState(false);

  const copyToClipboard = async () => {
    await navigator.clipboard.writeText(snippet.code);
    setCopied(true);
    setTimeout(() => setCopied(false), 2000);
  };

  return (
    <div className="bg-gray-50 rounded-lg border border-gray-200 overflow-hidden">
      <div className="bg-gray-100 px-4 py-3 border-b border-gray-200 flex items-center justify-between">
        <div className="flex items-center space-x-2">
          <Code className="w-4 h-4 text-gray-600" />
          <h4 className="font-medium text-gray-900">{snippet.title}</h4>
          <span className="text-xs bg-blue-100 text-blue-800 px-2 py-1 rounded">{snippet.category}</span>
        </div>
        <div className="flex items-center space-x-2">
          <button
            onClick={copyToClipboard}
            className="p-1 text-gray-500 hover:text-gray-700 transition-colors"
            title="Copy code"
          >
            {copied ? <Check className="w-4 h-4 text-green-600" /> : <Copy className="w-4 h-4" />}
          </button>
          <button
            onClick={onExecute}
            className="px-3 py-1 bg-blue-600 hover:bg-blue-700 text-white text-sm rounded transition-colors flex items-center space-x-1"
          >
            <Play className="w-3 h-3" />
            <span>Run</span>
          </button>
        </div>
      </div>
      <div className="p-4">
        <p className="text-sm text-gray-600 mb-3">{snippet.description}</p>
        <pre className="bg-gray-900 text-gray-100 p-3 rounded text-sm overflow-x-auto">
          <code>{snippet.code}</code>
        </pre>
      </div>
    </div>
  );
}

function App() {
  const [activeTab, setActiveTab] = useState<'database' | 'excel'>('database');
  const [logs, setLogs] = useState<LogEntry[]>([]);
  const [isLoading, setIsLoading] = useState(false);
  const ioAPI = useContext(IOConnectContext);


  // Database service state
  const [dbDataSource, setDbDataSource] = useState<DataSource>({
    file: 'interop.io/io.Connect Desktop/UserData/DEMO-INTEROP.IO/io.db',
    name: 'UserTable',
    columns: [
      { name: 'ID', type: ColumnType.Integer, pk: true, autoIncrement: true, nullable: false },
      { name: 'Name', type: ColumnType.Text, pk: false, autoIncrement: false, nullable: false },
      { name: 'Email', type: ColumnType.Text, pk: false, autoIncrement: false, nullable: true }
    ],
    primaryKey: ['ID'],
    data: [
      [null, 'John Doe', 'johndoe@example.com'],
      [null, 'Jane Smith', 'janesmith@example.com'],
      [null, 'Sam Wilson', 'samwilson@example.com']
    ]
  });

  const [queryText, setQueryText] = useState('SELECT * FROM UserTable');
  const [queryResults, setQueryResults] = useState<any[]>([]);

  // Excel service state
  const [workbookName, setWorkbookName] = useState('Book3');
  const [worksheetName, setWorksheetName] = useState('Sheet1');
  const [rangeValue, setRangeValue] = useState('A1:C10');
  const [cellValue, setCellValue] = useState('Hello World');
  const [tableName, setTableName] = useState('DemoTable');
  const [contextMenuCaption, setContextMenuCaption] = useState('Send Data');
  const [ribbonMenuCaption, setRibbonMenuCaption] = useState('Process Data');
  const [commentText, setCommentText] = useState('This is a demo comment');
  const [backgroundColor, setBackgroundColor] = useState('#FFE4B5');
  const [foregroundColor, setForegroundColor] = useState('#000000');
  const [fileName, setFileName] = useState('exported_file.xlsx');
  const [xlReference, setXlReference] = useState('Sheet1!A1:B5');
  const [subscriptionId, setSubscriptionId] = useState('sub_123');
  const [menuId, setMenuId] = useState('menu_123');
  const [fromRow, setFromRow] = useState(1);
  const [rowsToRead, setRowsToRead] = useState(10);
  const [rowPosition, setRowPosition] = useState(1);

  const dbService = useRef(new GlueDBService(ioAPI));
  const xlService = useRef(new GlueExcelService(ioAPI));

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

  const executeWithLogging = async (method: string, operation: () => Promise<any>, params?: any) => {
    setIsLoading(true);
    try {
      const result = await operation();
      addLog('success', method, 'Operation completed successfully', { params, result });
      return result;
    } catch (error) {
      addLog('error', method, `Error: ${error}`, { params, error });
      throw error;
    } finally {
      setIsLoading(false);
    }
  };

  // Database operations
  const initDatabase = () => executeWithLogging('DB.Init',
    () => dbService.current.init(dbDataSource), dbDataSource);

  const createTable = () => executeWithLogging('DB.CreateTable',
    () => dbService.current.createTable(dbDataSource), dbDataSource);

  const insertData = () => executeWithLogging('DB.InsertData',
    () => dbService.current.insertData(dbDataSource), dbDataSource);

  const executeQuery = async () => {
    const result = await executeWithLogging('DB.ExecuteQuery',
      () => dbService.current.executeQuery(dbDataSource, queryText), { query: queryText });
    if (result?.data) {
      setQueryResults(result.data);
    }
  };

  const updateRow = () => executeWithLogging('DB.UpdateRow',
    () => dbService.current.updateRow(dbDataSource, ['Updated Name', 'updated@example.com'], 1),
    { rowData: ['Updated Name', 'updated@example.com'], pkValue: 1 });

  const updateColumns = () => executeWithLogging('DB.UpdateColumns',
    () => dbService.current.updateColumns(dbDataSource, { Name: 'Updated Name', Email: 'updated@example.com' }, 1),
    { updates: { Name: 'Updated Name', Email: 'updated@example.com' }, pkValue: 1 });

  const disposeDatabase = () => executeWithLogging('DB.Dispose',
    () => dbService.current.dispose(dbDataSource), dbDataSource);

  // Excel operations - Basic
  const createWorkbook = () => executeWithLogging('XL.CreateWorkbook',
    () => xlService.current.createWorkbook(workbookName, worksheetName),
    { workbookName, worksheetName });

  const readRange = () => executeWithLogging('XL.Read',
    () => xlService.current.read({ workbook: workbookName, worksheet: worksheetName, range: rangeValue }),
    { workbook: workbookName, worksheet: worksheetName, range: rangeValue });

  const writeRange = () => executeWithLogging('XL.Write',
    () => xlService.current.write({ workbook: workbookName, worksheet: worksheetName, range: rangeValue }, cellValue),
    { workbook: workbookName, worksheet: worksheetName, range: rangeValue, value: cellValue });

  const readXlRef = () => executeWithLogging('XL.ReadXlRef',
    () => xlService.current.readXlRef(xlReference),
    { reference: xlReference });

  const writeXlRef = () => executeWithLogging('XL.WriteXlRef',
    () => xlService.current.writeXlRef(xlReference, cellValue),
    { reference: xlReference, value: cellValue });

  const openWorkbook = () => executeWithLogging('XL.OpenWorkbook',
    () => xlService.current.openWorkbook(fileName),
    { fileName });

  const saveWorkbook = () => executeWithLogging('XL.SaveAs',
    () => xlService.current.saveAs(
      { workbook: workbookName, worksheet: worksheetName, range: rangeValue },
      fileName
    ),
    { fileName });

  const activateRange = () => executeWithLogging('XL.Activate',
    () => xlService.current.activate({ workbook: workbookName, worksheet: worksheetName, range: rangeValue }),
    { range: rangeValue });

  // Excel operations - Subscriptions
  const subscribeToRange = () => executeWithLogging('XL.Subscribe',
    () => xlService.current.subscribe(
      { workbook: workbookName, worksheet: worksheetName, range: rangeValue },
      { callbackEndpoint: 'xlServiceCxtMenuCallback' }
    ),
    { range: rangeValue });

  const subscribeDeltas = () => executeWithLogging('XL.SubscribeDeltas',
    () => xlService.current.subscribeDeltas(
      { workbook: workbookName, worksheet: worksheetName, range: rangeValue },
      { callbackEndpoint: 'xlServiceCxtMenuCallback' }
    ),
    { range: rangeValue });

  const destroySubscription = () => executeWithLogging('XL.DestroySubscription',
    () => xlService.current.destroySubscription(subscriptionId),
    { subscriptionId });

  // Excel operations - Tables
  const createExcelTable = () => executeWithLogging('XL.CreateTable',
    () => xlService.current.createTable(
      { workbook: workbookName, worksheet: worksheetName, range: rangeValue },
      tableName,
      'TableStyleMedium2',
      ['ID', 'Name', 'Email'],
      [['1', 'John Doe', 'john@example.com'], ['2', 'Jane Smith', 'jane@example.com']],
      (origin, subscriptionId, ...props) => addLog('info', 'XL.TableCallback', 'Table callback triggered', { origin, subscriptionId, props })
    ),
    { tableName, range: rangeValue });

  const createLinkedTable = () => executeWithLogging('XL.CreateLinkedTable',
    () => xlService.current.createLinkedTable(
      { workbook: workbookName, worksheet: worksheetName, range: rangeValue },
      dbDataSource,
      { callbackEndpoint: 'xlServiceCxtMenuCallback' }
    ),
    { range: rangeValue, dataSource: dbDataSource });

  const refreshTable = () => executeWithLogging('XL.RefreshTable',
    () => xlService.current.refreshTable(
      { workbook: workbookName, worksheet: worksheetName, range: rangeValue },
      tableName
    ),
    { range: rangeValue, tableName });

  const writeTableRows = () => executeWithLogging('XL.WriteTableRows',
    () => xlService.current.writeTableRows(
      { workbook: workbookName, worksheet: worksheetName, range: rangeValue },
      tableName,
      rowPosition,
      [['3', 'New User', 'newuser@example.com']]
    ),
    { tableName, rowPosition, data: [['3', 'New User', 'newuser@example.com']] });

  const readTableRows = () => executeWithLogging('XL.ReadTableRows',
    () => xlService.current.readTableRows(
      { workbook: workbookName, worksheet: worksheetName, range: rangeValue },
      tableName,
      fromRow,
      rowsToRead
    ),
    { tableName, fromRow, rowsToRead });

  const updateTableColumns = () => executeWithLogging('XL.UpdateTableColumns',
    () => xlService.current.updateTableColumns(
      { workbook: workbookName, worksheet: worksheetName, range: rangeValue },
      tableName,
      [{ OldName: 'Email', Name: 'EmailAddress', Position: null, Op: 'Rename' as const }]
    ),
    { tableName, columnOps: [{ OldName: 'Email', Name: 'EmailAddress', Position: null, Op: 'Rename' }] });

  const describeTableColumns = () => executeWithLogging('XL.DescribeTableColumns',
    () => xlService.current.describeTableColumns(
      { workbook: workbookName, worksheet: worksheetName, range: rangeValue },
      tableName
    ),
    { tableName });

  // Excel operations - Menus
  const createContextMenu = () => executeWithLogging('XL.CreateContextMenu',
    () => xlService.current.createContextMenu(
      contextMenuCaption,
      ['io', 'actions'],
      { workbook: workbookName, worksheet: worksheetName, range: rangeValue },
      (origin, subscriptionId, ...props) => addLog('info', 'XL.ContextMenuCallback', 'Context menu clicked', { origin, subscriptionId, props })
    ),
    { caption: contextMenuCaption, range: rangeValue });

  const createContextMenuRaw = () => executeWithLogging('XL.CreateContextMenuRaw',
    () => xlService.current.createContextMenuRaw(
      contextMenuCaption,
      ['io', 'actions'],
      { workbook: workbookName, worksheet: worksheetName, range: rangeValue },
      { callbackEndpoint: 'xlServiceCxtMenuCallback' }
    ),
    { caption: contextMenuCaption, range: rangeValue });

  const destroyContextMenu = () => executeWithLogging('XL.DestroyContextMenu',
    () => xlService.current.destroyContextMenu(menuId),
    { menuId });

  const createRibbonMenu = () => executeWithLogging('XL.CreateDynamicRibbonMenu',
    () => xlService.current.createDynamicRibbonMenu(
      ribbonMenuCaption,
      { workbook: workbookName, worksheet: worksheetName, range: rangeValue },
      (origin, subscriptionId, ...props) => addLog('info', 'XL.RibbonMenuCallback', 'Ribbon menu clicked', { origin, subscriptionId, props })
    ),
    { caption: ribbonMenuCaption, range: rangeValue });

  const createRibbonMenuRaw = () => executeWithLogging('XL.CreateDynamicRibbonMenuRaw',
    () => xlService.current.createDynamicRibbonMenuRaw(
      ribbonMenuCaption,
      { workbook: workbookName, worksheet: worksheetName, range: rangeValue },
      { callbackEndpoint: 'xlServiceCxtMenuCallback' }
    ),
    { caption: ribbonMenuCaption, range: rangeValue });

  const destroyRibbonMenu = () => executeWithLogging('XL.DestroyRibbonMenu',
    () => xlService.current.destroyRibbonMenu(menuId),
    { menuId });

  // Excel operations - Styling & Comments
  const writeComment = () => executeWithLogging('XL.WriteComment',
    () => xlService.current.writeComment(
      { workbook: workbookName, worksheet: worksheetName, range: rangeValue },
      commentText
    ),
    { range: rangeValue, comment: commentText });

  const clearComments = () => executeWithLogging('XL.ClearComments',
    () => xlService.current.clearComments(
      { workbook: workbookName, worksheet: worksheetName, range: rangeValue }
    ),
    { range: rangeValue });

  const clearContents = () => executeWithLogging('XL.ClearContents',
    () => xlService.current.clearContents(
      { workbook: workbookName, worksheet: worksheetName, range: rangeValue }
    ),
    { range: rangeValue });

  const applyStyles = () => executeWithLogging('XL.ApplyStyles',
    () => xlService.current.applyStyles(
      { workbook: workbookName, worksheet: worksheetName, range: rangeValue },
      backgroundColor,
      foregroundColor
    ),
    { range: rangeValue, backgroundColor, foregroundColor });

  const clearLogs = () => setLogs([]);

  // Database code snippets
  const dbCodeSnippets: CodeSnippet[] = [
    {
      title: "Initialize Database",
      category: "Setup",
      description: "Initialize a database connection with a data source configuration",
      code: `const dataSource = {
  file: '${dbDataSource.file}',
  name: '${dbDataSource.name}',
  columns: [
    { name: 'ID', type: ColumnType.Integer, pk: true, autoIncrement: true, nullable: false },
    { name: 'Name', type: ColumnType.Text, pk: false, autoIncrement: false, nullable: false },
    { name: 'Email', type: ColumnType.Text, pk: false, autoIncrement: false, nullable: true }
  ],
  primaryKey: ['ID'],
  data: [
    [null, 'John Doe', 'johndoe@example.com'],
    [null, 'Jane Smith', 'janesmith@example.com']
  ]
};

await dbService.init(dataSource);`
    },
    {
      title: "Create Table",
      category: "Setup",
      description: "Create a new table in the database using the data source schema",
      code: `await dbService.createTable(dataSource);`
    },
    {
      title: "Insert Data",
      category: "Data",
      description: "Insert the data rows defined in the data source into the table",
      code: `await dbService.insertData(dataSource);`
    },
    {
      title: "Execute Query",
      category: "Data",
      description: "Execute a SQL query against the database",
      code: `const result = await dbService.executeQuery(dataSource, '${queryText}');
console.log(result.data);`
    },
    {
      title: "Update Row",
      category: "Data",
      description: "Update a specific row in the table using primary key",
      code: `const rowData = ['Updated Name', 'updated@example.com'];
const pkValue = 1;
await dbService.updateRow(dataSource, rowData, pkValue);`
    },
    {
      title: "Update Columns",
      category: "Data",
      description: "Update specific columns in a row using primary key",
      code: `const updates = { Name: 'Updated Name', Email: 'updated@example.com' };
const pkValue = 1;
await dbService.updateColumns(dataSource, updates, pkValue);`
    },
    {
      title: "Dispose Database",
      category: "Cleanup",
      description: "Clean up and dispose of the database connection",
      code: `await dbService.dispose(dataSource);`
    }
  ];

  // Excel code snippets
  const xlCodeSnippets: CodeSnippet[] = [
    // Basic Operations
    {
      title: "Create Workbook",
      category: "Basic",
      description: "Create a new Excel workbook with a specified worksheet",
      code: `await xlService.createWorkbook('${workbookName}', '${worksheetName}');`
    },
    {
      title: "Open Workbook",
      category: "Basic",
      description: "Open an existing Excel workbook from file",
      code: `await xlService.openWorkbook('${fileName}');`
    },
    {
      title: "Save Workbook",
      category: "Basic",
      description: "Save the workbook to a specific file path",
      code: `const range = { workbook: '${workbookName}', worksheet: '${worksheetName}', range: '${rangeValue}' };
await xlService.saveAs(range, '${fileName}');`
    },
    {
      title: "Activate Range",
      category: "Basic",
      description: "Activate and select a specific range in Excel",
      code: `const range = { workbook: '${workbookName}', worksheet: '${worksheetName}', range: '${rangeValue}' };
await xlService.activate(range);`
    },

    // Read/Write Operations
    {
      title: "Read Range",
      category: "Read/Write",
      description: "Read data from a specific range in an Excel worksheet",
      code: `const range = { workbook: '${workbookName}', worksheet: '${worksheetName}', range: '${rangeValue}' };
const result = await xlService.read(range);`
    },
    {
      title: "Write Range",
      category: "Read/Write",
      description: "Write data to a specific range in an Excel worksheet",
      code: `const range = { workbook: '${workbookName}', worksheet: '${worksheetName}', range: '${rangeValue}' };
await xlService.write(range, '${cellValue}');`
    },
    {
      title: "Read Excel Reference",
      category: "Read/Write",
      description: "Read data using Excel reference notation",
      code: `const result = await xlService.readXlRef('${xlReference}');`
    },
    {
      title: "Write Excel Reference",
      category: "Read/Write",
      description: "Write data using Excel reference notation",
      code: `await xlService.writeXlRef('${xlReference}', '${cellValue}');`
    },

    // Subscription Operations
    {
      title: "Subscribe to Range",
      category: "Subscriptions",
      description: "Subscribe to changes in a specific Excel range",
      code: `const range = { workbook: '${workbookName}', worksheet: '${worksheetName}', range: '${rangeValue}' };
const subscriptionInfo = { callbackEndpoint: 'xlServiceCxtMenuCallback' };
const result = await xlService.subscribe(range, subscriptionInfo);`
    },
    {
      title: "Subscribe to Deltas",
      category: "Subscriptions",
      description: "Subscribe to delta changes in a range with data top-left position",
      code: `const range = { workbook: '${workbookName}', worksheet: '${worksheetName}', range: '${rangeValue}' };
const subscriptionInfo = { callbackEndpoint: 'xlServiceCxtMenuCallback' };
await xlService.subscribeDeltas(range,  subscriptionInfo);`
    },
    {
      title: "Destroy Subscription",
      category: "Subscriptions",
      description: "Remove an active subscription",
      code: `await xlService.destroySubscription('${subscriptionId}');`
    },

    // Table Operations
    {
      title: "Create Excel Table",
      category: "Tables",
      description: "Create a formatted table in Excel with data and styling",
      code: `const range = { workbook: '${workbookName}', worksheet: '${worksheetName}', range: '${rangeValue}' };
const columns = ['ID', 'Name', 'Email'];
const data = [['1', 'John Doe', 'john@example.com'], ['2', 'Jane Smith', 'jane@example.com']];

await xlService.createTable(
  range, 
  '${tableName}', 
  'TableStyleMedium2', 
  columns, 
  data,
  (origin, subscriptionId, ...props) => {
    console.log('Table callback triggered', { origin, subscriptionId, props });
  }
);`
    },
    {
      title: "Create Linked Table",
      category: "Tables",
      description: "Create a table linked to a database data source",
      code: `const range = { workbook: '${workbookName}', worksheet: '${worksheetName}', range: '${rangeValue}' };
const subscriptionInfo = { callbackEndpoint: 'xlServiceCxtMenuCallback' };

await xlService.createLinkedTable(range, dataSource, subscriptionInfo);`
    },
    {
      title: "Refresh Table",
      category: "Tables",
      description: "Refresh data in an existing table",
      code: `const range = { workbook: '${workbookName}', worksheet: '${worksheetName}', range: '${rangeValue}' };
await xlService.refreshTable(range, '${tableName}');`
    },
    {
      title: "Write Table Rows",
      category: "Tables",
      description: "Write new rows to an existing table",
      code: `const range = { workbook: '${workbookName}', worksheet: '${worksheetName}', range: '${rangeValue}' };
const newRows = [['3', 'New User', 'newuser@example.com']];
await xlService.writeTableRows(range, '${tableName}', ${rowPosition}, newRows);`
    },
    {
      title: "Read Table Rows",
      category: "Tables",
      description: "Read specific rows from a table",
      code: `const range = { workbook: '${workbookName}', worksheet: '${worksheetName}', range: '${rangeValue}' };
const result = await xlService.readTableRows(range, '${tableName}', ${fromRow}, ${rowsToRead});`
    },
    {
      title: "Update Table Columns",
      category: "Tables",
      description: "Add, remove, or rename columns in a table",
      code: `const range = { workbook: '${workbookName}', worksheet: '${worksheetName}', range: '${rangeValue}' };
const columnOps = [
  { OldName: 'Email', Name: 'EmailAddress', Position: null, Op: 'Rename' }
];
await xlService.updateTableColumns(range, '${tableName}', columnOps);`
    },
    {
      title: "Describe Table Columns",
      category: "Tables",
      description: "Get information about table columns",
      code: `const range = { workbook: '${workbookName}', worksheet: '${worksheetName}', range: '${rangeValue}' };
const result = await xlService.describeTableColumns(range, '${tableName}');`
    },

    // Menu Operations
    {
      title: "Create Context Menu",
      category: "Menus",
      description: "Add a custom context menu item to Excel with callback",
      code: `const range = { workbook: '${workbookName}', worksheet: '${worksheetName}', range: '${rangeValue}' };

await xlService.createContextMenu(
  '${contextMenuCaption}',
  ['io', 'actions'],
  range,
  (origin, subscriptionId, ...props) => {
    console.log('Context menu clicked', { origin, subscriptionId, props });
  }
);`
    },
    {
      title: "Create Context Menu (Raw)",
      category: "Menus",
      description: "Add a custom context menu item with raw subscription info",
      code: `const range = { workbook: '${workbookName}', worksheet: '${worksheetName}', range: '${rangeValue}' };
const subscriptionInfo = { callbackEndpoint: 'xlServiceCxtMenuCallback' };

await xlService.createContextMenuRaw(
  '${contextMenuCaption}',
  ['io', 'actions'],
  range,
  subscriptionInfo
);`
    },
    {
      title: "Destroy Context Menu",
      category: "Menus",
      description: "Remove a context menu item",
      code: `await xlService.destroyContextMenu('${menuId}');`
    },
    {
      title: "Create Ribbon Menu",
      category: "Menus",
      description: "Add a custom ribbon menu item to Excel with callback",
      code: `const range = { workbook: '${workbookName}', worksheet: '${worksheetName}', range: '${rangeValue}' };

await xlService.createDynamicRibbonMenu(
  '${ribbonMenuCaption}',
  range,
  (origin, subscriptionId, ...props) => {
    console.log('Ribbon menu clicked', { origin, subscriptionId, props });
  }
);`
    },
    {
      title: "Create Ribbon Menu (Raw)",
      category: "Menus",
      description: "Add a custom ribbon menu item with raw subscription info",
      code: `const range = { workbook: '${workbookName}', worksheet: '${worksheetName}', range: '${rangeValue}' };
const subscriptionInfo = { callbackEndpoint: 'xlServiceCxtMenuCallback' };

await xlService.createDynamicRibbonMenuRaw(
  '${ribbonMenuCaption}',
  range,
  subscriptionInfo
);`
    },
    {
      title: "Destroy Ribbon Menu",
      category: "Menus",
      description: "Remove a ribbon menu item",
      code: `await xlService.destroyRibbonMenu('${menuId}');`
    },

    // Styling & Comments
    {
      title: "Write Comment",
      category: "Styling",
      description: "Add a comment to a specific cell or range",
      code: `const range = { workbook: '${workbookName}', worksheet: '${worksheetName}', range: '${rangeValue}' };
await xlService.writeComment(range, '${commentText}');`
    },
    {
      title: "Clear Comments",
      category: "Styling",
      description: "Remove all comments from a range",
      code: `const range = { workbook: '${workbookName}', worksheet: '${worksheetName}', range: '${rangeValue}' };
await xlService.clearComments(range);`
    },
    {
      title: "Clear Contents",
      category: "Styling",
      description: "Clear all content from a range",
      code: `const range = { workbook: '${workbookName}', worksheet: '${worksheetName}', range: '${rangeValue}' };
await xlService.clearContents(range);`
    },
    {
      title: "Apply Styles",
      category: "Styling",
      description: "Apply background and foreground colors to a range",
      code: `const range = { workbook: '${workbookName}', worksheet: '${worksheetName}', range: '${rangeValue}' };
await xlService.applyStyles(range, '${backgroundColor}', '${foregroundColor}');`
    }
  ];

  // Group Excel snippets by category
  const groupedXlSnippets = xlCodeSnippets.reduce((acc, snippet) => {
    if (!acc[snippet.category]) {
      acc[snippet.category] = [];
    }
    acc[snippet.category].push(snippet);
    return acc;
  }, {} as Record<string, CodeSnippet[]>);

  const executeExcelOperation = (snippetIndex: number, category: string) => {
    const operations = {
      'Basic': [createWorkbook, openWorkbook, saveWorkbook, activateRange],
      'Read/Write': [readRange, writeRange, readXlRef, writeXlRef],
      'Subscriptions': [subscribeToRange, subscribeDeltas, destroySubscription],
      'Tables': [createExcelTable, createLinkedTable, refreshTable, writeTableRows, readTableRows, updateTableColumns, describeTableColumns],
      'Menus': [createContextMenu, createContextMenuRaw, destroyContextMenu, createRibbonMenu, createRibbonMenuRaw, destroyRibbonMenu],
      'Styling': [writeComment, clearComments, clearContents, applyStyles]
    };

    const categoryOps = operations[category as keyof typeof operations];
    if (categoryOps && categoryOps[snippetIndex]) {
      categoryOps[snippetIndex]();
    }
  };

  const executeDatabaseOperation = (snippetIndex: number) => {
    const operations = [initDatabase, createTable, insertData, executeQuery, updateRow, updateColumns, disposeDatabase];
    if (operations[snippetIndex]) {
      operations[snippetIndex]();
    }
  };

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
            {/* Tab Navigation */}
            <div className="bg-white rounded-xl shadow-sm border border-gray-200">
              <div className="border-b border-gray-200">
                <nav className="flex space-x-8 px-6">
                  <button
                    onClick={() => setActiveTab('database')}
                    className={`py-4 px-1 border-b-2 font-medium text-sm transition-colors duration-200 ${activeTab === 'database'
                      ? 'border-blue-500 text-blue-600'
                      : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'
                      }`}
                  >
                    <div className="flex items-center space-x-2">
                      <Database className="w-4 h-4" />
                      <span>Database Service ({dbCodeSnippets.length} methods)</span>
                    </div>
                  </button>
                  <button
                    onClick={() => setActiveTab('excel')}
                    className={`py-4 px-1 border-b-2 font-medium text-sm transition-colors duration-200 ${activeTab === 'excel'
                      ? 'border-blue-500 text-blue-600'
                      : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'
                      }`}
                  >
                    <div className="flex items-center space-x-2">
                      <FileSpreadsheet className="w-4 h-4" />
                      <span>Excel Service ({xlCodeSnippets.length} methods)</span>
                    </div>
                  </button>
                </nav>
              </div>

              <div className="p-6">
                {activeTab === 'database' && (
                  <div className="space-y-8">
                    {/* Database Configuration */}
                    <div className="bg-gray-50 rounded-lg p-6">
                      <h3 className="text-lg font-semibold text-gray-900 mb-4 flex items-center">
                        <Settings className="w-5 h-5 mr-2" />
                        Database Configuration
                      </h3>
                      <div className="grid grid-cols-1 md:grid-cols-2 gap-4">
                        <div>
                          <label className="block text-sm font-medium text-gray-700 mb-2">Database File</label>
                          <input
                            type="text"
                            value={dbDataSource.file}
                            onChange={(e) => setDbDataSource(prev => ({ ...prev, file: e.target.value }))}
                            className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                          />
                        </div>
                        <div>
                          <label className="block text-sm font-medium text-gray-700 mb-2">Table Name</label>
                          <input
                            type="text"
                            value={dbDataSource.name}
                            onChange={(e) => setDbDataSource(prev => ({ ...prev, name: e.target.value }))}
                            className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                          />
                        </div>
                      </div>
                    </div>

                    {/* Database Code Examples */}
                    <div className="space-y-6">
                      {dbCodeSnippets.map((snippet, index) => (
                        <CodeBlock
                          key={index}
                          snippet={snippet}
                          onExecute={() => executeDatabaseOperation(index)}
                        />
                      ))}
                    </div>

                    {/* Query Section */}
                    <div className="bg-gray-50 rounded-lg p-4">
                      <h3 className="text-lg font-semibold text-gray-900 mb-4">Custom SQL Query</h3>
                      <textarea
                        value={queryText}
                        onChange={(e) => setQueryText(e.target.value)}
                        rows={3}
                        className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent font-mono text-sm"
                        placeholder="Enter your SQL query here..."
                      />
                    </div>

                    {/* Query Results */}
                    {queryResults.length > 0 && (
                      <div className="bg-white border border-gray-200 rounded-lg overflow-hidden">
                        <div className="px-4 py-3 bg-gray-50 border-b border-gray-200">
                          <h3 className="text-lg font-semibold text-gray-900">Query Results</h3>
                        </div>
                        <div className="overflow-x-auto">
                          <table className="min-w-full divide-y divide-gray-200">
                            <thead className="bg-gray-50">
                              <tr>
                                {Object.keys(queryResults[0] || {}).map((key) => (
                                  <th key={key} className="px-6 py-3 text-left text-xs font-medium text-gray-500 uppercase tracking-wider">
                                    {key}
                                  </th>
                                ))}
                              </tr>
                            </thead>
                            <tbody className="bg-white divide-y divide-gray-200">
                              {queryResults.map((row, index) => (
                                <tr key={index} className="hover:bg-gray-50">
                                  {Object.values(row).map((value: any, cellIndex) => (
                                    <td key={cellIndex} className="px-6 py-4 whitespace-nowrap text-sm text-gray-900">
                                      {String(value)}
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
                )}

                {activeTab === 'excel' && (
                  <div className="space-y-8">
                    {/* Excel Configuration */}
                    <div className="bg-gray-50 rounded-lg p-6">
                      <h3 className="text-lg font-semibold text-gray-900 mb-4 flex items-center">
                        <Settings className="w-5 h-5 mr-2" />
                        Excel Configuration
                      </h3>
                      <div className="grid grid-cols-1 md:grid-cols-3 gap-4 mb-4">
                        <div>
                          <label className="block text-sm font-medium text-gray-700 mb-2">Workbook</label>
                          <input
                            type="text"
                            value={workbookName}
                            onChange={(e) => setWorkbookName(e.target.value)}
                            className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                          />
                        </div>
                        <div>
                          <label className="block text-sm font-medium text-gray-700 mb-2">Worksheet</label>
                          <input
                            type="text"
                            value={worksheetName}
                            onChange={(e) => setWorksheetName(e.target.value)}
                            className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                          />
                        </div>
                        <div>
                          <label className="block text-sm font-medium text-gray-700 mb-2">Range</label>
                          <input
                            type="text"
                            value={rangeValue}
                            onChange={(e) => setRangeValue(e.target.value)}
                            className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                          />
                        </div>
                      </div>

                      <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-4 gap-4 mb-4">
                        <div>
                          <label className="block text-sm font-medium text-gray-700 mb-2">Cell Value</label>
                          <input
                            type="text"
                            value={cellValue}
                            onChange={(e) => setCellValue(e.target.value)}
                            className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                          />
                        </div>
                        <div>
                          <label className="block text-sm font-medium text-gray-700 mb-2">Table Name</label>
                          <input
                            type="text"
                            value={tableName}
                            onChange={(e) => setTableName(e.target.value)}
                            className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                          />
                        </div>
                        <div>
                          <label className="block text-sm font-medium text-gray-700 mb-2">File Name</label>
                          <input
                            type="text"
                            value={fileName}
                            onChange={(e) => setFileName(e.target.value)}
                            className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                          />
                        </div>
                        <div>
                          <label className="block text-sm font-medium text-gray-700 mb-2">Excel Reference</label>
                          <input
                            type="text"
                            value={xlReference}
                            onChange={(e) => setXlReference(e.target.value)}
                            className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                          />
                        </div>
                        <div>
                          <label className="block text-sm font-medium text-gray-700 mb-2">Context Menu</label>
                          <input
                            type="text"
                            value={contextMenuCaption}
                            onChange={(e) => setContextMenuCaption(e.target.value)}
                            className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                          />
                        </div>
                        <div>
                          <label className="block text-sm font-medium text-gray-700 mb-2">Ribbon Menu</label>
                          <input
                            type="text"
                            value={ribbonMenuCaption}
                            onChange={(e) => setRibbonMenuCaption(e.target.value)}
                            className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                          />
                        </div>
                        <div>
                          <label className="block text-sm font-medium text-gray-700 mb-2">Comment Text</label>
                          <input
                            type="text"
                            value={commentText}
                            onChange={(e) => setCommentText(e.target.value)}
                            className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                          />
                        </div>
                        <div>
                          <label className="block text-sm font-medium text-gray-700 mb-2">Subscription ID</label>
                          <input
                            type="text"
                            value={subscriptionId}
                            onChange={(e) => setSubscriptionId(e.target.value)}
                            className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                          />
                        </div>
                        <div>
                          <label className="block text-sm font-medium text-gray-700 mb-2">Menu ID</label>
                          <input
                            type="text"
                            value={menuId}
                            onChange={(e) => setMenuId(e.target.value)}
                            className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                          />
                        </div>
                        <div>
                          <label className="block text-sm font-medium text-gray-700 mb-2">From Row</label>
                          <input
                            type="number"
                            value={fromRow}
                            onChange={(e) => setFromRow(parseInt(e.target.value))}
                            className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                          />
                        </div>
                        <div>
                          <label className="block text-sm font-medium text-gray-700 mb-2">Rows to Read</label>
                          <input
                            type="number"
                            value={rowsToRead}
                            onChange={(e) => setRowsToRead(parseInt(e.target.value))}
                            className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                          />
                        </div>
                        <div>
                          <label className="block text-sm font-medium text-gray-700 mb-2">Row Position</label>
                          <input
                            type="number"
                            value={rowPosition}
                            onChange={(e) => setRowPosition(parseInt(e.target.value))}
                            className="w-full px-3 py-2 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                          />
                        </div>
                        <div>
                          <label className="block text-sm font-medium text-gray-700 mb-2">Background Color</label>
                          <input
                            type="color"
                            value={backgroundColor}
                            onChange={(e) => setBackgroundColor(e.target.value)}
                            className="w-full h-10 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                          />
                        </div>
                        <div>
                          <label className="block text-sm font-medium text-gray-700 mb-2">Foreground Color</label>
                          <input
                            type="color"
                            value={foregroundColor}
                            onChange={(e) => setForegroundColor(e.target.value)}
                            className="w-full h-10 border border-gray-300 rounded-md focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent"
                          />
                        </div>
                      </div>
                    </div>

                    {/* Excel Code Examples by Category */}
                    {Object.entries(groupedXlSnippets).map(([category, snippets]) => (
                      <div key={category} className="space-y-6">
                        <div className="flex items-center space-x-2 mb-4">
                          <div className="bg-blue-100 p-2 rounded-lg">
                            {category === 'Basic' && <FileSpreadsheet className="w-5 h-5 text-blue-600" />}
                            {category === 'Read/Write' && <Edit3 className="w-5 h-5 text-blue-600" />}
                            {category === 'Subscriptions' && <Zap className="w-5 h-5 text-blue-600" />}
                            {category === 'Tables' && <Table className="w-5 h-5 text-blue-600" />}
                            {category === 'Menus' && <Menu className="w-5 h-5 text-blue-600" />}
                            {category === 'Styling' && <Palette className="w-5 h-5 text-blue-600" />}
                          </div>
                          <h3 className="text-xl font-bold text-gray-900">{category} Operations</h3>
                          <span className="text-sm bg-gray-100 text-gray-600 px-2 py-1 rounded">{snippets.length} methods</span>
                        </div>
                        {snippets.map((snippet, index) => (
                          <CodeBlock
                            key={`${category}-${index}`}
                            snippet={snippet}
                            onExecute={() => executeExcelOperation(index, category)}
                          />
                        ))}
                      </div>
                    ))}
                  </div>
                )}
              </div>
            </div>
          </div>

          {/* Activity Log Sidebar */}
          <div className="lg:col-span-1">
            <div className="bg-white rounded-xl shadow-sm border border-gray-200 h-fit max-h-[800px] flex flex-col sticky top-8">
              <div className="px-6 py-4 border-b border-gray-200">
                <h3 className="text-lg font-semibold text-gray-900 flex items-center">
                  <Activity className="w-5 h-5 mr-2" />
                  Activity Log
                </h3>
                <p className="text-sm text-gray-500 mt-1">{logs.length} operations logged</p>
              </div>
              <div className="flex-1 overflow-y-auto p-4 space-y-3">
                {logs.length === 0 ? (
                  <div className="text-center py-8 text-gray-500">
                    <Activity className="w-8 h-8 mx-auto mb-2 opacity-50" />
                    <p>No operations yet</p>
                    <p className="text-xs">Execute operations to see logs here</p>
                  </div>
                ) : (
                  logs.map((log) => (
                    <div
                      key={log.id}
                      className={`p-3 rounded-lg border-l-4 ${log.type === 'success'
                        ? 'bg-green-50 border-green-400'
                        : log.type === 'error'
                          ? 'bg-red-50 border-red-400'
                          : 'bg-blue-50 border-blue-400'
                        }`}
                    >
                      <div className="flex items-start justify-between">
                        <div className="flex items-center space-x-2">
                          {log.type === 'success' && <CheckCircle className="w-4 h-4 text-green-600" />}
                          {log.type === 'error' && <AlertCircle className="w-4 h-4 text-red-600" />}
                          {log.type === 'info' && <Info className="w-4 h-4 text-blue-600" />}
                          <span className="text-xs font-medium text-gray-600">{log.timestamp}</span>
                        </div>
                      </div>
                      <div className="mt-1">
                        <p className="text-sm font-medium text-gray-900">{log.method}</p>
                        <p className="text-xs text-gray-600 mt-1">{log.message}</p>
                        {log.params && (
                          <details className="mt-2">
                            <summary className="text-xs text-gray-500 cursor-pointer hover:text-gray-700">
                              View details
                            </summary>
                            <pre className="text-xs text-gray-600 mt-1 bg-gray-100 p-2 rounded overflow-x-auto">
                              {JSON.stringify(log.params, null, 2)}
                            </pre>
                          </details>
                        )}
                      </div>
                    </div>
                  ))
                )}
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  );
}

export default App;