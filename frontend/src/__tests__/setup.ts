/**
 * Jest setup — provide minimal Office.js globals so engine code can import
 * without crashing, even though we never execute real Excel operations.
 */

// eslint-disable-next-line @typescript-eslint/no-explicit-any
(globalThis as any).Excel = {
  run: jest.fn(),
  RequestContext: jest.fn(),
};

// eslint-disable-next-line @typescript-eslint/no-explicit-any
(globalThis as any).Office = {
  onReady: jest.fn(),
};
