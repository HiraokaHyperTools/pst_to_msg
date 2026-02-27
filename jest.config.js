export default {
  testEnvironment: 'node',
  transform: {
    "^.+\\.(t|j)sx?$": "babel-jest",
  },
  moduleNameMapper: {
    '^(\\.{1,2}/.*)\\.js$': '$1',
    '^@kenjiuno/msgreader-web-ng$': '<rootDir>/node_modules/@kenjiuno/msgreader-web-ng/lib/index.js',
    '^@hiraokahypertools/pst-extractor$': '<rootDir>/node_modules/@hiraokahypertools/pst-extractor/dist/index.js',
  },
  transformIgnorePatterns: [
    "/node_modules/(?!(@kenjiuno/msgreader-web-ng|@hiraokahypertools/pst-extractor)/)"
  ],
};