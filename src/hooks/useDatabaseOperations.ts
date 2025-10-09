import { useRef, useContext } from 'react';
import { IOConnectContext } from "@interopio/react-hooks";
import { IOConnectDBService, DataSource } from '../io-excel-service';

export function useDatabaseOperations() {
  const ioAPI = useContext(IOConnectContext);
  const dbService = useRef(new IOConnectDBService(ioAPI));

  const operations = {
    initDatabase: (dataSource: DataSource) =>
      dbService.current.init(dataSource),

    createTable: (dataSource: DataSource) =>
      dbService.current.createTable(dataSource),

    insertData: (dataSource: DataSource) =>
      dbService.current.insertData(dataSource),

    executeQuery: (dataSource: DataSource, queryText: string) =>
      dbService.current.executeQuery(dataSource, queryText),

    updateRow: (dataSource: DataSource, rowData: any[], pkValue: any) =>
      dbService.current.updateRow(dataSource, rowData, pkValue),

    updateColumns: (dataSource: DataSource, updates: Record<string, any>, pkValue: any) =>
      dbService.current.updateColumns(dataSource, updates, pkValue),

    disposeDatabase: (dataSource: DataSource) =>
      dbService.current.dispose(dataSource)
  };

  return operations;
}
