// jest.config.mjs
export default {
  preset: 'ts-jest',
  testEnvironment: 'node',
  transform: {
    '^.+\\.tsx?$': ['ts-jest', {
      tsconfig: 'tsconfig.test.json'
    }]
  },
  moduleFileExtensions: ['ts', 'js', 'json', 'node'],
  moduleNameMapper: {
    // Map imports ending in .js back to .ts files for resolution
    '^(\\.{1,2}/.*)\\.js$': '$1',
    // Ensure @/ alias maps correctly using <rootDir>
    '^@/(.*)\\.js$': '<rootDir>/src/$1', // Explicitly map .js alias
    '^@/(.*)$': '<rootDir>/src/$1'      // Keep original mapping
  },
  // Ignore all node_modules EXCEPT pptxgenjs (which needs transformation)
  transformIgnorePatterns: ['/node_modules/(?!pptxgenjs|fs-extra/)'],
  testPathIgnorePatterns: ['/node_modules/', '/dist/'],
  collectCoverageFrom: ['src/**/*.ts'],
  coveragePathIgnorePatterns: ['/node_modules/', '/dist/'],
};