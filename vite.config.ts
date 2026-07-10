/// <reference types="vitest" />
import { readFileSync } from 'node:fs';
import { defineConfig } from 'vite';
import react from '@vitejs/plugin-react';
import { getHttpsServerOptions } from 'office-addin-dev-certs';

const packageJson = JSON.parse(
  readFileSync(new URL('./package.json', import.meta.url), 'utf-8'),
) as { version?: string };

export default defineConfig(async ({ command }) => {
  const httpsOptions = await getHttpsServerOptions();
  return {
    base: command === 'build' ? '/SheetClaw/' : '/',
    define: {
      __APP_VERSION__: JSON.stringify(packageJson.version ?? '0.0.0'),
    },
    plugins: [react()],
    server: {
      port: 3000,
      https: httpsOptions,
    },
    preview: {
      port: 3000,
      https: httpsOptions,
    },
    build: {
      rollupOptions: {
        input: {
          taskpane: 'taskpane.html',
          oauthStart: 'oauth-start.html',
          oauthCallback: 'oauth-callback.html',
        },
      },
      outDir: 'dist',
    },
    test: {
      environment: 'node',
      include: ['src/**/*.test.ts'],
    },
  };
});
