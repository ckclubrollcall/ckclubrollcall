export default [
    {
        files: ["gas/**/*.js"],
        languageOptions: {
            globals: {
                SpreadsheetApp: "readonly",
                DriveApp: "readonly",
                GmailApp: "readonly",
                Session: "readonly",
                Logger: "readonly",
                Utilities: "readonly",
                PropertiesService: "readonly",
                UrlFetchApp: "readonly",
                HtmlService: "readonly",
                ContentService: "readonly",
                LockService: "readonly",
                CacheService: "readonly",
                ScriptApp: "readonly",
                CalendarApp: "readonly",
            },
        },
    },
];