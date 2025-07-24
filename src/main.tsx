import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import App from './App.tsx';
import IODesktop from "@interopio/desktop";
import { IOConnectProvider } from "@interopio/react-hooks";

import './index.css';

createRoot(document.getElementById('root')!).render(
    <IOConnectProvider fallback={<h2>Loading...</h2>} settings={{
      desktop: {
        config: {},
        factory: IODesktop,
      }
    }} >
      {/* @ts-ignore */}
      <App />
    </IOConnectProvider>
);
