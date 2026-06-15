import { createRoot } from 'react-dom/client';
import App from './App';
import { useStore } from '../store/index';

Office.onReady(() => {
  // Hydrate persisted config and auth before first render. Auth decrypts
  // asynchronously; the store re-renders subscribers when it lands.
  useStore.getState().loadConfigFromStorage();
  useStore.getState().loadChatHistory();
  void useStore.getState().loadAuthFromStorage();

  const container = document.getElementById('root');
  if (!container) throw new Error('Root element not found');
  createRoot(container).render(<App />);
});
