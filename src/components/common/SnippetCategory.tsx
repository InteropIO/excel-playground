import { FileSpreadsheet, Edit3, Zap, Table, Menu, Palette } from 'lucide-react';
import { CodeSnippet } from '../../types';
import { CodeBlock } from '../common/CodeBlock';

interface SnippetCategoryProps {
  category: string;
  snippets: CodeSnippet[];
  isCollapsed: boolean;
  onToggle: () => void;
  onExecuteSnippet: (index: number) => void;
}

export function SnippetCategory({ 
  category, 
  snippets, 
  isCollapsed, 
  onToggle, 
  onExecuteSnippet 
}: SnippetCategoryProps) {
  const getCategoryIcon = () => {
    switch (category) {
      case 'Basic':
        return <FileSpreadsheet className="w-5 h-5 text-blue-600" />;
      case 'Read/Write':
        return <Edit3 className="w-5 h-5 text-blue-600" />;
      case 'Subscriptions':
        return <Zap className="w-5 h-5 text-blue-600" />;
      case 'Tables':
        return <Table className="w-5 h-5 text-blue-600" />;
      case 'Menus':
        return <Menu className="w-5 h-5 text-blue-600" />;
      case 'Styling':
        return <Palette className="w-5 h-5 text-blue-600" />;
      default:
        return <FileSpreadsheet className="w-5 h-5 text-blue-600" />;
    }
  };

  return (
    <div className="space-y-6">
      <div 
        className="flex items-center justify-between p-4 bg-gray-50 rounded-lg border border-gray-200 cursor-pointer hover:bg-gray-100 transition-colors"
        onClick={onToggle}
      >
        <div className="flex items-center space-x-2">
          <div className="bg-blue-100 p-2 rounded-lg">
            {getCategoryIcon()}
          </div>
          <h3 className="text-xl font-bold text-gray-900">{category} Operations</h3>
          <span className="text-sm bg-gray-100 text-gray-600 px-2 py-1 rounded">{snippets.length} methods</span>
        </div>
        <div className="flex items-center space-x-2">
          <span className="text-sm text-gray-500">
            {isCollapsed ? 'Show' : 'Hide'}
          </span>
          <div className={`transform transition-transform duration-200 ${isCollapsed ? 'rotate-180' : ''}`}>
            <svg className="w-5 h-5 text-gray-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
              <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 9l-7 7-7-7" />
            </svg>
          </div>
        </div>
      </div>
      
      {!isCollapsed && (
        <div className="space-y-6 pl-4">
          {snippets.map((snippet, index) => (
            <CodeBlock
              key={`${category}-${index}`}
              snippet={snippet}
              onExecute={() => onExecuteSnippet(index)}
            />
          ))}
        </div>
      )}
    </div>
  );
}
