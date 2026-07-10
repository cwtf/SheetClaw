import { defineConfig } from 'vitest/config';

// Live-network integration checks (npm run test:providers). Kept out of the
// default `npm test` run: they hit real endpoints and need connectivity.
export default defineConfig({
  test: {
    environment: 'node',
    include: ['src/**/*.integration.ts'],
    testTimeout: 120_000,
  },
});
