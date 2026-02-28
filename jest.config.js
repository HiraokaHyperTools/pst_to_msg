export default {
  preset: 'ts-jest/presets/default-esm',
  testEnvironment: 'node',
  transform: {
    '^.+\\.(ts|tsx|js)$': ['ts-jest', {
      tsconfig: 'tsconfig.test.json'
    }]
  },
  moduleFileExtensions: ['ts', 'tsx', 'js', 'jsx', 'json', 'node'],
  moduleNameMapper: {
    '^(\\.{1,2}/.*)\\.js$': '$1',
    '^@kenjiuno/msgreader-web-ng$': '<rootDir>/node_modules/@kenjiuno/msgreader-web-ng/lib/index.js',
    '^@hiraokahypertools/pst-extractor$': '<rootDir>/node_modules/@hiraokahypertools/pst-extractor/dist/index.js',
  },
  transformIgnorePatterns: [
    "/node_modules/(?!(@kenjiuno/msgreader-web-ng|@hiraokahypertools/pst-extractor)/)"
  ]
};