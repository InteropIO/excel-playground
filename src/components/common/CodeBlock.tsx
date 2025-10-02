import { useState } from 'react';
import { Code, Copy, Check, Play } from 'lucide-react';
import SyntaxHighlighter from 'react-syntax-highlighter';
import { dracula } from 'react-syntax-highlighter/dist/esm/styles/hljs';
import { CodeSnippet } from '../../types';

interface CodeBlockProps {
  snippet: CodeSnippet;
  onExecute: () => void;
}

export function CodeBlock({ snippet, onExecute }: CodeBlockProps) {
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
        <SyntaxHighlighter language="javascript" style={dracula} className="rounded text-sm overflow-x-auto">
          {snippet.code}
        </SyntaxHighlighter>
      </div>
    </div>
  );
}
