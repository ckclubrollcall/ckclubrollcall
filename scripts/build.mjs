import { writeFile } from "node:fs/promises";

const GAS_URL = process.env.GAS_URL;

if (!GAS_URL) {
    throw new Error("缺少 GAS_URL 環境變數");
}

const url = new URL(GAS_URL);

if (!["http:", "https:"].includes(url.protocol)) {
    throw new Error("GAS_URL 必須使用 HTTP 或 HTTPS");
}

await writeFile(
    "public/config.js",
    `window.APP_CONFIG = Object.freeze(${JSON.stringify({
        GAS_URL: GAS_URL.replace(/\/+$/, ""),
    })});\n`,
    "utf8",
);