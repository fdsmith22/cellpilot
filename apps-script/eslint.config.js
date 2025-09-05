const js = require("@eslint/js");
const googleConfig = require("eslint-config-google");

module.exports = [
  js.configs.recommended,
  {
    files: ["**/*.js"],
    languageOptions: {
      ecmaVersion: 2020,
      sourceType: "script",
      globals: {
        // Google Apps Script globals
        SpreadsheetApp: "readonly",
        SlidesApp: "readonly",
        Charts: "readonly",
        DriveApp: "readonly",
        DocumentApp: "readonly",
        FormApp: "readonly",
        CalendarApp: "readonly",
        UrlFetchApp: "readonly",
        Utilities: "readonly",
        Logger: "readonly",
        console: "readonly",
        HtmlService: "readonly",
        ContentService: "readonly",
        ScriptApp: "readonly",
        PropertiesService: "readonly",
        CacheService: "readonly",
        LockService: "readonly",
        Session: "readonly",
        Browser: "readonly",
        // Custom globals
        CellPilot: "readonly",
        CellM8: "writable"
      }
    },
    rules: {
      ...googleConfig.rules,
      "no-unused-vars": ["warn", { "args": "none" }],
      "no-undef": "error",
      "no-redeclare": "off", // CellM8 is a global
      "max-len": ["warn", { "code": 120 }],
      "require-jsdoc": "off",
      "valid-jsdoc": "off",
      "camelcase": "warn",
      "new-cap": "off",
      "space-before-function-paren": ["error", {
        "anonymous": "always",
        "named": "never",
        "asyncArrow": "always"
      }],
      "object-curly-spacing": ["error", "always"],
      "indent": ["error", 2],
      "comma-dangle": ["error", "only-multiline"]
    }
  }
];