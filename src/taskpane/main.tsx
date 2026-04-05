import { StrictMode } from 'react';
import { createRoot } from 'react-dom/client';
import { FluentProvider, webLightTheme } from '@fluentui/react-components';
import App from './components/App';

/* global Office */

Office.onReady(() => {
  const root = document.getElementById('root');
  if (!root) return;

  createRoot(root).render(
    <StrictMode>
      <FluentProvider theme={webLightTheme}>
        <App />
      </FluentProvider>
    </StrictMode>
  );
});
