import { useRef, useContext, useCallback } from 'react';
import { IOConnectContext } from "@interopio/react-hooks";
import { IOConnectExcelService, DataSource } from '../io-excel-service';
import { ExcelState } from '../types';

export function useExcelOperations() {
  const ioAPI = useContext(IOConnectContext);
  const xlService = useRef<IOConnectExcelService | null>(null);

  // Initialize service only once
  if (!xlService.current) {
    xlService.current = new IOConnectExcelService(ioAPI);
    window.xlService = xlService.current;
    window.io = ioAPI;
    console.log('Created new IOConnectExcelService instance:', xlService.current);
  }

  const createOperations = useCallback((state: ExcelState, dataSource: DataSource, addLog: any) => {
    const createRange = () => ({
      workbook: state.workbookName,
      worksheet: state.worksheetName,
      range: state.rangeValue
    });

    return {
      // Basic Operations
      createWorkbook: () =>
        xlService.current!.createWorkbook(state.workbookName, state.worksheetName),

      openWorkbook: () =>
        xlService.current!.openWorkbook(state.fileName),

      saveWorkbook: () =>
        xlService.current!.saveAs(createRange(), state.fileName),

      activateRange: () =>
        xlService.current!.activate(createRange()),

      // Read/Write Operations
      readRange: () =>
        xlService.current!.read(createRange()),

      writeRange: () =>
        xlService.current!.write(createRange(), state.cellValue as any),

      readExcelRef: () =>
        xlService.current!.readExcelRef(state.xlReference),

      writeExcelRef: () =>
        xlService.current!.writeExcelRef(state.xlReference, state.cellValue as any),

      // Subscription Operations
      subscribeToRange: () =>
        xlService.current!.subscribe(
          createRange(),
          (origin, subscriptionId, ...props) => addLog('info', 'XL.SubscribeCallback', 'Subscribe callback triggered', { origin, subscriptionId, props })
        ),

      subscribeDeltas: () =>
        xlService.current!.subscribeDeltas(
          createRange(),
          (origin, subscriptionId, ...props) => addLog('info', 'XL.SubscribeDeltasCallback', 'Subscribe deltas callback triggered', { origin, subscriptionId, props })
        ),

      destroySubscription: () =>
        xlService.current!.destroySubscription(state.subscriptionId),

      // Table Operations
      createExcelTable: () =>
        xlService.current!.createTable(
          createRange(),
          state.tableName,
          'TableStyleMedium2',
          ['ID', 'Name', 'Email'],
          [['1', 'John Doe', 'john@example.com'], ['2', 'Jane Smith', 'jane@example.com']] as any,
          (origin, subscriptionId, ...props) => addLog('info', 'XL.TableCallback', 'Table callback triggered', { origin, subscriptionId, props })
        ),

      createLinkedTable: () =>
        xlService.current!.createLinkedTable(
          createRange(),
          dataSource,
          { callbackEndpoint: 'xlServiceCxtMenuCallback' }
        ),

      refreshTable: () =>
        xlService.current!.refreshTable(createRange(), state.tableName),

      writeTableRows: () =>
        xlService.current!.writeTableRows(
          createRange(),
          state.tableName,
          state.rowPosition,
          [['3', 'New User', 'newuser@example.com']] as any
        ),

      readTableRows: () =>
        xlService.current!.readTableRows(
          createRange(),
          state.tableName,
          state.fromRow,
          state.rowsToRead
        ),

      updateTableColumns: () =>
        xlService.current!.updateTableColumns(
          createRange(),
          state.tableName,
          [{ currentName: 'Email', newName: 'EmailAddress', position: null, operation: 'Rename' as const }]
        ),

      describeTableColumns: () =>
        xlService.current!.describeTableColumns(createRange(), state.tableName),

      // Menu Operations
      createContextMenu: () =>
        xlService.current!.createContextMenu(
          state.contextMenuCaption,
          ['io', 'actions'],
          createRange(),
          (origin, subscriptionId, ...props) => addLog('info', 'XL.ContextMenuCallback', 'Context menu clicked', { origin, subscriptionId, props })
        ),

      createContextMenuRaw: () =>
        xlService.current!.createContextMenuRaw(
          state.contextMenuCaption,
          ['io', 'actions'],
          createRange(),
          { callbackEndpoint: 'xlServiceCxtMenuCallback' }
        ),

      destroyContextMenu: () =>
        xlService.current!.destroyContextMenu(state.menuId),

      createRibbonMenu: () =>
        xlService.current!.createDynamicRibbonMenu(
          state.ribbonMenuCaption,
          createRange(),
          (origin, subscriptionId, ...props) => addLog('info', 'XL.RibbonMenuCallback', 'Ribbon menu clicked', { origin, subscriptionId, props })
        ),

      createRibbonMenuRaw: () =>
        xlService.current!.createDynamicRibbonMenuRaw(
          state.ribbonMenuCaption,
          createRange(),
          { callbackEndpoint: 'xlServiceCxtMenuCallback' }
        ),

      destroyRibbonMenu: () =>
        xlService.current!.destroyRibbonMenu(state.menuId),

      // Styling & Comments
      writeComment: () =>
        xlService.current!.writeComment(createRange(), state.commentText),

      clearComments: () =>
        xlService.current!.clearComments(createRange()),

      clearContents: () =>
        xlService.current!.clearContents(createRange()),

      applyStyles: () =>
        xlService.current!.applyStyles(createRange(), state.backgroundColor, state.foregroundColor)
    };
  }, []);

  return { service: xlService.current, createOperations };
}
