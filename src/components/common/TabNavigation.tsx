import { Database, FileSpreadsheet } from 'lucide-react';

interface TabNavigationProps {
  activeTab: 'database' | 'excel';
  onTabChange: (tab: 'database' | 'excel') => void;
}

export function TabNavigation({ activeTab, onTabChange }: TabNavigationProps) {
  return (
    <nav className="flex space-x-8 px-6">
      <button
        onClick={() => onTabChange('excel')}
        className={`py-4 px-1 border-b-2 font-medium text-sm transition-colors duration-200 ${
          activeTab === 'excel'
            ? 'border-blue-500 text-blue-600'
            : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'
        }`}
      >
        <div className="flex items-center space-x-2">
          <FileSpreadsheet className="w-4 h-4" />
          <span>Excel Service (25 methods)</span>
        </div>
      </button>
      <button
        onClick={() => onTabChange('database')}
        className={`py-4 px-1 border-b-2 font-medium text-sm transition-colors duration-200 ${
          activeTab === 'database'
            ? 'border-blue-500 text-blue-600'
            : 'border-transparent text-gray-500 hover:text-gray-700 hover:border-gray-300'
        }`}
      >
        <div className="flex items-center space-x-2">
          <Database className="w-4 h-4" />
          <span>Database Service (7 methods)</span>
        </div>
      </button>
    </nav>
  );
}
