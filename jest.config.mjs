// jest.config.mjs
export default {
  preset: 'ts-jest/presets/default-esm',
  testEnvironment: 'node',
  extensionsToTreatAsEsm: ['.ts'],
  transform: {
    // Transform both .ts/.tsx and .js/.jsx files with ts-jest in ESM mode
    '^.+\\.(t|j)sx?$': [
      'ts-jest',
      {
        useESM: true,
      },
    ],
  },
  // Add .ts to the list of extensions Jest should look for
  moduleFileExtensions: ['ts', 'tsx', 'js', 'jsx', 'json', 'node'],
  moduleNameMapper: {
    // Map imports ending in .js back to .ts files for resolution
    '^(\\.{1,2}/.*)\\.js$': '$1',
    // Ensure @/ alias maps correctly using <rootDir>
    '^@/(.*)\\.js$': '<rootDir>/src/$1', // Explicitly map .js alias
    '^@/(.*)$': '<rootDir>/src/$1'      // Keep original mapping
  },
  // Ignore all node_modules EXCEPT pptxgenjs (which needs transformation)
  transformIgnorePatterns: ['/node_modules/(?!pptxgenjs/)'],
  testPathIgnorePatterns: ['/node_modules/', '/dist/'],
  collectCoverageFrom: ['src/**/*.ts'],
  coveragePathIgnorePatterns: ['/node_modules/', '/dist/'],
};