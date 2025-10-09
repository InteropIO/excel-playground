import { CodeSnippet, DatabaseState, ExcelState } from '../types';

export function createDatabaseSnippets(state: DatabaseState): CodeSnippet[] {
  return [
    {
      title: "Initialize Database",
      category: "Setup",
      description: "Initialize a database connection with a data source configuration",
      code: `const dataSource = {
  name: '${state.dataSource.name}',
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
      code: `const result = await dbService.executeQuery(dataSource, '${state.queryText}');
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
}

export function createExcelSnippets(state: ExcelState): CodeSnippet[] {
  return [
    // Basic Operations
    {
      title: "Create Workbook",
      category: "Basic",
      description: "Create a new Excel workbook with a specified worksheet",
      code: `await xlService.createWorkbook('${state.workbookName}', '${state.worksheetName}');`
    },
    {
      title: "Open Workbook",
      category: "Basic",
      description: "Open an existing Excel workbook from file",
      code: `await xlService.openWorkbook('${state.fileName}');`
    },
    {
      title: "Save Workbook",
      category: "Basic",
      description: "Save the workbook to a specific file path",
      code: `const range = { workbook: '${state.workbookName}', worksheet: '${state.worksheetName}', range: '${state.rangeValue}' };
await xlService.saveAs(range, '${state.fileName}');`
    },
    {
      title: "Activate Range",
      category: "Basic",
      description: "Activate and select a specific range in Excel",
      code: `const range = { workbook: '${state.workbookName}', worksheet: '${state.worksheetName}', range: '${state.rangeValue}' };
await xlService.activate(range);`
    },

    // Read/Write Operations
    {
      title: "Read Range",
      category: "Read/Write",
      description: "Read data from a specific range in an Excel worksheet",
      code: `const range = { workbook: '${state.workbookName}', worksheet: '${state.worksheetName}', range: '${state.rangeValue}' };
const result = await xlService.read(range);`
    },
    {
      title: "Write Range",
      category: "Read/Write",
      description: "Write data to a specific range in an Excel worksheet",
      code: `const range = { workbook: '${state.workbookName}', worksheet: '${state.worksheetName}', range: '${state.rangeValue}' };
await xlService.write(range, '${state.cellValue}');`
    },
    {
      title: "Read Excel Reference",
      category: "Read/Write",
      description: "Read data using Excel reference notation",
      code: `const result = await xlService.readRef('${state.xlReference}');`
    },
    {
      title: "Write Excel Reference",
      category: "Read/Write",
      description: "Write data using Excel reference notation",
      code: `await xlService.writeRef('${state.xlReference}', '${state.cellValue}');`
    },

    // Subscription Operations
    {
      title: "Subscribe to Range",
      category: "Subscriptions",
      description: "Subscribe to changes in a specific Excel range",
      code: `const range = { workbook: '${state.workbookName}', worksheet: '${state.worksheetName}', range: '${state.rangeValue}' };
const callback = (origin, subscriptionId, ...props) => console.log('Subscribe callback triggered', origin, subscriptionId, props);
const result = await xlService.subscribe(range, callback);`
    },
    {
      title: "Subscribe to Deltas",
      category: "Subscriptions",
      description: "Subscribe to delta changes in a range with data top-left position",
      code: `const range = { workbook: '${state.workbookName}', worksheet: '${state.worksheetName}', range: '${state.rangeValue}' };
const callback = (origin, subscriptionId, ...props) => console.log('Subscribe deltas callback triggered', origin, subscriptionId, props);
await xlService.subscribeDeltas(range, callback);`
    },
    {
      title: "Destroy Subscription",
      category: "Subscriptions",
      description: "Remove an active subscription",
      code: `await xlService.destroySubscription('${state.subscriptionId}');`
    },

    // Table Operations
    {
      title: "Create Excel Table",
      category: "Tables",
      description: "Create a formatted table in Excel with data and styling",
      code: `const range = { workbook: '${state.workbookName}', worksheet: '${state.worksheetName}', range: '${state.rangeValue}' };
const columns = ['ID', 'Name', 'Email'];
const data = [['1', 'John Doe', 'john@example.com'], ['2', 'Jane Smith', 'jane@example.com']];

await xlService.createTable(
  range,
  '${state.tableName}',
  'TableStyleMedium2',
  columns,
  data,
  (origin, subscriptionId, ...props) => {
    console.log('Table callback triggered', origin, subscriptionId, props);
  }
);`
    },
    {
      title: "Create Linked Table",
      category: "Tables",
      description: "Create a table linked to a database data source",
      code: `const range = { workbook: '${state.workbookName}', worksheet: '${state.worksheetName}', range: '${state.rangeValue}' };
const subscriptionInfo = { callbackEndpoint: 'xlServiceCxtMenuCallback' };

await xlService.createLinkedTable(range, dataSource, subscriptionInfo);`
    },
    {
      title: "Refresh Table",
      category: "Tables",
      description: "Refresh data in an existing table",
      code: `const range = { workbook: '${state.workbookName}', worksheet: '${state.worksheetName}', range: '${state.rangeValue}' };
await xlService.refreshTable(range, '${state.tableName}');`
    },
    {
      title: "Write Table Rows",
      category: "Tables",
      description: "Write new rows to an existing table",
      code: `const range = { workbook: '${state.workbookName}', worksheet: '${state.worksheetName}', range: '${state.rangeValue}' };
const newRows = [['3', 'New User', 'newuser@example.com']];
await xlService.writeTableRows(range, '${state.tableName}', ${state.rowPosition}, newRows);`
    },
    {
      title: "Read Table Rows",
      category: "Tables",
      description: "Read specific rows from a table",
      code: `const range = { workbook: '${state.workbookName}', worksheet: '${state.worksheetName}', range: '${state.rangeValue}' };
const result = await xlService.readTableRows(range, '${state.tableName}', ${state.fromRow}, ${state.rowsToRead});`
    },
    {
      title: "Update Table Columns",
      category: "Tables",
      description: "Add, remove, or rename columns in a table",
      code: `const range = { workbook: '${state.workbookName}', worksheet: '${state.worksheetName}', range: '${state.rangeValue}' };
const columnOps = [
  { oldName: 'Email', name: 'EmailAddress', position: null, op: 'Rename' }
];
await xlService.updateTableColumns(range, '${state.tableName}', columnOps);`
    },
    {
      title: "Describe Table Columns",
      category: "Tables",
      description: "Get information about table columns",
      code: `const range = { workbook: '${state.workbookName}', worksheet: '${state.worksheetName}', range: '${state.rangeValue}' };
const result = await xlService.describeTableColumns(range, '${state.tableName}');`
    },

    // Menu Operations
    {
      title: "Create Context Menu",
      category: "Menus",
      description: "Add a custom context menu item to Excel with callback",
      code: `const range = { workbook: '${state.workbookName}', worksheet: '${state.worksheetName}', range: '${state.rangeValue}' };

await xlService.createContextMenu(
  '${state.contextMenuCaption}',
  ['io', 'actions'],
  range,
  (origin, subscriptionId, ...props) => {
    console.log('Context menu clicked', origin, subscriptionId, props);
  }
);`
    },
    {
      title: "Create Context Menu (Raw)",
      category: "Menus",
      description: "Add a custom context menu item with raw subscription info",
      code: `const range = { workbook: '${state.workbookName}', worksheet: '${state.worksheetName}', range: '${state.rangeValue}' };
const subscriptionInfo = { callbackEndpoint: 'xlServiceCxtMenuCallback' };

await xlService.createContextMenuRaw(
  '${state.contextMenuCaption}',
  ['io', 'actions'],
  range,
  subscriptionInfo
);`
    },
    {
      title: "Destroy Context Menu",
      category: "Menus",
      description: "Remove a context menu item",
      code: `await xlService.destroyContextMenu('${state.menuId}');`
    },
    {
      title: "Create Ribbon Menu",
      category: "Menus",
      description: "Add a custom ribbon menu item to Excel with callback",
      code: `const range = { workbook: '${state.workbookName}', worksheet: '${state.worksheetName}', range: '${state.rangeValue}' };

await xlService.createDynamicRibbonMenu(
  '${state.ribbonMenuCaption}',
  range,
  (origin, subscriptionId, ...props) => {
    console.log('Ribbon menu clicked', origin, subscriptionId, props);
  }
);`
    },
    {
      title: "Create Ribbon Menu (Raw)",
      category: "Menus",
      description: "Add a custom ribbon menu item with raw subscription info",
      code: `const range = { workbook: '${state.workbookName}', worksheet: '${state.worksheetName}', range: '${state.rangeValue}' };
const subscriptionInfo = { callbackEndpoint: 'xlServiceCxtMenuCallback' };

await xlService.createDynamicRibbonMenuRaw(
  '${state.ribbonMenuCaption}',
  range,
  subscriptionInfo
);`
    },
    {
      title: "Destroy Ribbon Menu",
      category: "Menus",
      description: "Remove a ribbon menu item",
      code: `await xlService.destroyRibbonMenu('${state.menuId}');`
    },

    // Styling & Comments
    {
      title: "Write Comment",
      category: "Styling",
      description: "Add a comment to a specific cell or range",
      code: `const range = { workbook: '${state.workbookName}', worksheet: '${state.worksheetName}', range: '${state.rangeValue}' };
await xlService.writeComment(range, '${state.commentText}');`
    },
    {
      title: "Clear Comments",
      category: "Styling",
      description: "Remove all comments from a range",
      code: `const range = { workbook: '${state.workbookName}', worksheet: '${state.worksheetName}', range: '${state.rangeValue}' };
await xlService.clearComments(range);`
    },
    {
      title: "Clear Contents",
      category: "Styling",
      description: "Clear all content from a range",
      code: `const range = { workbook: '${state.workbookName}', worksheet: '${state.worksheetName}', range: '${state.rangeValue}' };
await xlService.clearContents(range);`
    },
    {
      title: "Apply Styles",
      category: "Styling",
      description: "Apply background and foreground colors to a range",
      code: `const range = { workbook: '${state.workbookName}', worksheet: '${state.worksheetName}', range: '${state.rangeValue}' };
await xlService.applyStyles(range, '${state.backgroundColor}', '${state.foregroundColor}');`
    }
  ];
}

export function groupSnippetsByCategory(snippets: CodeSnippet[]): Record<string, CodeSnippet[]> {
  return snippets.reduce((acc, snippet) => {
    if (!acc[snippet.category]) {
      acc[snippet.category] = [];
    }
    acc[snippet.category].push(snippet);
    return acc;
  }, {} as Record<string, CodeSnippet[]>);
}
