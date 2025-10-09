import { Settings } from 'lucide-react';
import { ExcelState } from '../../types';

interface ExcelConfigProps {
  state: ExcelState;
  onStateChange: (updates: Partial<ExcelState>) => void;
}

export function ExcelConfig({ state, onStateChange }: ExcelConfigProps) {
  const updateState = (updates: Partial<ExcelState>) => {
    onStateChange(updates);
  };

  return (
    <div className="sticky top-4 z-10 bg-white rounded-lg shadow-md border border-gray-200 p-4">
      <h3 className="text-md font-semibold text-gray-900 mb-3 flex items-center">
        <Settings className="w-4 h-4 mr-2" />
        Excel Configuration
      </h3>

      {/* Primary Configuration - Always Visible */}
      <div className="grid grid-cols-2 md:grid-cols-3 lg:grid-cols-6 gap-3 mb-3">
        <div>
          <label className="block text-xs font-medium text-gray-600 mb-1">Workbook</label>
          <input
            type="text"
            value={state.workbookName}
            onChange={(e) => updateState({ workbookName: e.target.value })}
            className="w-full px-2 py-1 text-sm border border-gray-300 rounded focus:outline-none focus:ring-1 focus:ring-blue-500 focus:border-transparent"
          />
        </div>
        <div>
          <label className="block text-xs font-medium text-gray-600 mb-1">Worksheet</label>
          <input
            type="text"
            value={state.worksheetName}
            onChange={(e) => updateState({ worksheetName: e.target.value })}
            className="w-full px-2 py-1 text-sm border border-gray-300 rounded focus:outline-none focus:ring-1 focus:ring-blue-500 focus:border-transparent"
          />
        </div>
        <div>
          <label className="block text-xs font-medium text-gray-600 mb-1">Range</label>
          <input
            type="text"
            value={state.rangeValue}
            onChange={(e) => updateState({ rangeValue: e.target.value })}
            className="w-full px-2 py-1 text-sm border border-gray-300 rounded focus:outline-none focus:ring-1 focus:ring-blue-500 focus:border-transparent"
          />
        </div>
        <div>
          <label className="block text-xs font-medium text-gray-600 mb-1">Cell Value</label>
          <input
            type="text"
            value={state.cellValue}
            onChange={(e) => updateState({ cellValue: e.target.value })}
            className="w-full px-2 py-1 text-sm border border-gray-300 rounded focus:outline-none focus:ring-1 focus:ring-blue-500 focus:border-transparent"
          />
        </div>
        <div>
          <label className="block text-xs font-medium text-gray-600 mb-1">Table Name</label>
          <input
            type="text"
            value={state.tableName}
            onChange={(e) => updateState({ tableName: e.target.value })}
            className="w-full px-2 py-1 text-sm border border-gray-300 rounded focus:outline-none focus:ring-1 focus:ring-blue-500 focus:border-transparent"
          />
        </div>
        <div>
          <label className="block text-xs font-medium text-gray-600 mb-1">File Name</label>
          <input
            type="text"
            value={state.fileName}
            onChange={(e) => updateState({ fileName: e.target.value })}
            className="w-full px-2 py-1 text-sm border border-gray-300 rounded focus:outline-none focus:ring-1 focus:ring-blue-500 focus:border-transparent"
          />
        </div>
      </div>

      {/* Collapsible Advanced Configuration */}
      <details className="group">
        <summary className="cursor-pointer text-xs text-gray-600 hover:text-gray-800 flex items-center space-x-1">
          <span>Advanced Options</span>
          <svg className="w-3 h-3 transform group-open:rotate-180 transition-transform" fill="none" stroke="currentColor" viewBox="0 0 24 24">
            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 9l-7 7-7-7" />
          </svg>
        </summary>

        <div className="mt-3 grid grid-cols-2 md:grid-cols-3 lg:grid-cols-5 gap-3">
          <div>
            <label className="block text-xs font-medium text-gray-600 mb-1">Excel Ref</label>
            <input
              type="text"
              value={state.xlReference}
              onChange={(e) => updateState({ xlReference: e.target.value })}
              className="w-full px-2 py-1 text-sm border border-gray-300 rounded focus:outline-none focus:ring-1 focus:ring-blue-500 focus:border-transparent"
            />
          </div>
          <div>
            <label className="block text-xs font-medium text-gray-600 mb-1">Context Menu</label>
            <input
              type="text"
              value={state.contextMenuCaption}
              onChange={(e) => updateState({ contextMenuCaption: e.target.value })}
              className="w-full px-2 py-1 text-sm border border-gray-300 rounded focus:outline-none focus:ring-1 focus:ring-blue-500 focus:border-transparent"
            />
          </div>
          <div>
            <label className="block text-xs font-medium text-gray-600 mb-1">Ribbon Menu</label>
            <input
              type="text"
              value={state.ribbonMenuCaption}
              onChange={(e) => updateState({ ribbonMenuCaption: e.target.value })}
              className="w-full px-2 py-1 text-sm border border-gray-300 rounded focus:outline-none focus:ring-1 focus:ring-blue-500 focus:border-transparent"
            />
          </div>
          <div>
            <label className="block text-xs font-medium text-gray-600 mb-1">Comment</label>
            <input
              type="text"
              value={state.commentText}
              onChange={(e) => updateState({ commentText: e.target.value })}
              className="w-full px-2 py-1 text-sm border border-gray-300 rounded focus:outline-none focus:ring-1 focus:ring-blue-500 focus:border-transparent"
            />
          </div>
          <div>
            <label className="block text-xs font-medium text-gray-600 mb-1">Sub ID</label>
            <input
              type="text"
              value={state.subscriptionId}
              onChange={(e) => updateState({ subscriptionId: e.target.value })}
              className="w-full px-2 py-1 text-sm border border-gray-300 rounded focus:outline-none focus:ring-1 focus:ring-blue-500 focus:border-transparent"
            />
          </div>
          <div>
            <label className="block text-xs font-medium text-gray-600 mb-1">Menu ID</label>
            <input
              type="text"
              value={state.menuId}
              onChange={(e) => updateState({ menuId: e.target.value })}
              className="w-full px-2 py-1 text-sm border border-gray-300 rounded focus:outline-none focus:ring-1 focus:ring-blue-500 focus:border-transparent"
            />
          </div>
          <div>
            <label className="block text-xs font-medium text-gray-600 mb-1">From Row</label>
            <input
              type="number"
              value={state.fromRow}
              onChange={(e) => updateState({ fromRow: parseInt(e.target.value) })}
              className="w-full px-2 py-1 text-sm border border-gray-300 rounded focus:outline-none focus:ring-1 focus:ring-blue-500 focus:border-transparent"
            />
          </div>
          <div>
            <label className="block text-xs font-medium text-gray-600 mb-1">Rows to Read</label>
            <input
              type="number"
              value={state.rowsToRead}
              onChange={(e) => updateState({ rowsToRead: parseInt(e.target.value) })}
              className="w-full px-2 py-1 text-sm border border-gray-300 rounded focus:outline-none focus:ring-1 focus:ring-blue-500 focus:border-transparent"
            />
          </div>
          <div>
            <label className="block text-xs font-medium text-gray-600 mb-1">Row Position</label>
            <input
              type="number"
              value={state.rowPosition}
              onChange={(e) => updateState({ rowPosition: parseInt(e.target.value) })}
              className="w-full px-2 py-1 text-sm border border-gray-300 rounded focus:outline-none focus:ring-1 focus:ring-blue-500 focus:border-transparent"
            />
          </div>
          <div>
            <label className="block text-xs font-medium text-gray-600 mb-1">BG Color</label>
            <input
              type="color"
              value={state.backgroundColor}
              onChange={(e) => updateState({ backgroundColor: e.target.value })}
              className="w-full h-7 border border-gray-300 rounded focus:outline-none focus:ring-1 focus:ring-blue-500 focus:border-transparent"
            />
          </div>
          <div>
            <label className="block text-xs font-medium text-gray-600 mb-1">FG Color</label>
            <input
              type="color"
              value={state.foregroundColor}
              onChange={(e) => updateState({ foregroundColor: e.target.value })}
              className="w-full h-7 border border-gray-300 rounded focus:outline-none focus:ring-1 focus:ring-blue-500 focus:border-transparent"
            />
          </div>
        </div>
      </details>
    </div>
  );
}
