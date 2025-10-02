export interface LogEntry {
  id: string;
  timestamp: string;
  type: 'success' | 'error' | 'info';
  method: string;
  message: string;
  params?: any;
}

export interface CodeSnippet {
  title: string;
  code: string;
  description: string;
  category: string;
}

export interface DatabaseState {
  dataSource: import('../io-excel-service').DataSource;
  queryText: string;
  queryResults: any[];
}

export interface ExcelState {
  workbookName: string;
  worksheetName: string;
  rangeValue: string;
  cellValue: string;
  tableName: string;
  contextMenuCaption: string;
  ribbonMenuCaption: string;
  commentText: string;
  backgroundColor: string;
  foregroundColor: string;
  fileName: string;
  xlReference: string;
  subscriptionId: string;
  menuId: string;
  fromRow: number;
  rowsToRead: number;
  rowPosition: number;
}
