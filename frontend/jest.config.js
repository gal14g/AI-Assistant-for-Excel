/** @type {import('ts-jest').JestConfigWithTsJest} */
module.exports = {
  preset: "ts-jest",
  testEnvironment: "node",
  rootDir: "src",
  testMatch: ["**/__tests__/**/*.test.ts"],
  moduleNameMapper: {
    "^@engine/(.*)$": "<rootDir>/engine/$1",
    "^@services/(.*)$": "<rootDir>/services/$1",
    "^@shared/(.*)$": "<rootDir>/shared/$1",
    "^@components/(.*)$": "<rootDir>/taskpane/components/$1",
    "^@hooks/(.*)$": "<rootDir>/taskpane/hooks/$1",
  },
  // Mock Office.js globals that aren't available in Node
  globals: {
    Excel: {},
  },
  setupFiles: ["<rootDir>/__tests__/setup.ts"],
};
