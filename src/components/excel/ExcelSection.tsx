import { useState } from 'react';
import { ExcelConfig } from './ExcelConfig';
import { SnippetCategory } from '../common/SnippetCategory';
import { ExcelState, CodeSnippet } from '../../types';

interface ExcelSectionProps {
  state: ExcelState;
  onStateChange: (updates: Partial<ExcelState>) => void;
  groupedSnippets: Record<string, CodeSnippet[]>;
  onExecuteSnippet: (index: number, category: string) => void;
}

export function ExcelSection({
  state,
  onStateChange,
  groupedSnippets,
  onExecuteSnippet
}: ExcelSectionProps) {
  const [collapsedCategories, setCollapsedCategories] = useState<Record<string, boolean>>({});

  const toggleCategory = (category: string) => {
    setCollapsedCategories(prev => ({
      ...prev,
      [category]: !prev[category]
    }));
  };

  return (
    <div className="space-y-8">
      <ExcelConfig 
        state={state} 
        onStateChange={onStateChange} 
      />

      {/* Excel Code Examples by Category */}
      {Object.entries(groupedSnippets).map(([category, snippets]) => (
        <SnippetCategory
          key={category}
          category={category}
          snippets={snippets}
          isCollapsed={collapsedCategories[category] || false}
          onToggle={() => toggleCategory(category)}
          onExecuteSnippet={(index) => onExecuteSnippet(index, category)}
        />
      ))}
    </div>
  );
}
