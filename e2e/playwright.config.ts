import { defineConfig } from '@playwright/test';

export default defineConfig({
  testDir: './tests',
  timeout: 120_000,
  expect: {
    toHaveScreenshot: {
      maxDiffPixelRatio: 0.02,
    },
  },
  use: {
    baseURL: 'http://localhost:8080/web/',
  },
  webServer: {
    command: 'python -m http.server 8080',
    cwd: '..',
    port: 8080,
    reuseExistingServer: true,
    timeout: 10_000,
  },
  projects: [
    {
      name: 'chromium',
      use: {
        browserName: 'chromium',
        viewport: { width: 1280, height: 900 },
      },
    },
  ],
});
