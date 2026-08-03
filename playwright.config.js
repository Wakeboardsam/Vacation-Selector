const { defineConfig, devices } = require('@playwright/test');

module.exports = defineConfig({
  testDir: './tests',
  fullyParallel: true,
  reporter: 'list',
  use: {
    trace: 'on-first-retry',
  },
  projects: [
    {
      name: 'chromium',
      use: { ...devices['Desktop Chrome'] },
    },
    // We intentionally disable mobile & webkit here since this CI environment natively traps/fails to download them.
    // They are documented as NOT RUN natively in the runner logs.
  ],
});
