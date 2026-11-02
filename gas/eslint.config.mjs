import js from "@eslint/js";
import { defineConfig, globalIgnores } from "eslint/config";

export default defineConfig([
    globalIgnores(["node_modules/", "dist/"]),
    js.configs.recommended,
    {
        files: ["**/*.js"],
        languageOptions: {
            ecmaVersion: "latest",
            sourceType: "script",
            globals: {
                SpreadsheetApp: "readonly",
                DocumentApp: "readonly",
                DriveApp: "readonly",
                GmailApp: "readonly",
                Session: "readonly",
                Logger: "readonly",
                Utilities: "readonly",
                PropertiesService: "readonly",
                CacheService: "readonly",
                UrlFetchApp: "readonly",
                HtmlService: "readonly",
                ContentService: "readonly",
                LockService: "readonly",
                ScriptApp: "readonly",
                Sheets: "readonly",
                logToSheet: "readonly",
                getSignFolderId: "readonly",
                getReportFolderId: "readonly",
                parseDriveId: "readonly",
            },
        },
        rules: {
            eqeqeq: "error",
            "no-unused-vars": ["warn", { args: "after-used" }],
            "no-use-before-define": ["error", { functions: false, classes: true }],
            "prefer-const": ["warn", { destructuring: "all" }],
            "no-var": "warn",
            "no-undef": "warn",
            "no-console": "off",
            curly: ["error", "all"],
            semi: ["error", "always"],
            quotes: ["warn", "double"],
        },
    },
]);