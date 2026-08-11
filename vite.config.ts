/// <reference types="vitest" />
import { copyFileSync, readFileSync } from 'node:fs';
import { defineConfig, type Plugin } from 'vite';
import react from '@vitejs/plugin-react';
import { getHttpsServerOptions } from 'office-addin-dev-certs';

const packageJson = JSON.parse(
  readFileSync(new URL('./package.json', import.meta.url), 'utf-8'),
) as { version?: string };

/**
 * Publish manifest.xml alongside the built app.
 *
 * It cannot live in public/ because `npm run validate-manifest` and
 * `npm run sideload` both address it at the repo root, so it is copied at build
 * time instead. Without this the GitHub Pages deploy (which uploads dist/) has
 * no manifest.xml and install.ps1's download 404s.
 */
function copyManifest(): Plugin {
  return {
    name: 'sheetclaw-copy-manifest',
    apply: 'build',
    closeBundle() {
      copyFileSync('manifest.xml', 'dist/manifest.xml');
    },
  };
}

export default defineConfig(async ({ command }) => {
  const httpsOptions = await getHttpsServerOptions();
  return {
    base: command === 'build' ? '/SheetClaw/' : '/',
    define: {
      __APP_VERSION__: JSON.stringify(packageJson.version ?? '0.0.0'),
    },
    plugins: [react(), copyManifest()],
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
