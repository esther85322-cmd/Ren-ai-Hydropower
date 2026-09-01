// After `vite build`, duplicate dist/index.html into dist/checkin.html with
// its own <title>/Open Graph tags, so sharing the /checkin link (LINE, etc.)
// shows "仁愛水電打卡" in the link preview instead of the main app's title.
// It's a straight copy of the built index.html (same hashed script tag), so
// it always stays in sync with whatever main.jsx/CheckinPage.jsx built to —
// nothing here references file names that could go stale between builds.
// Requires firebase.json's "cleanUrls": true so /checkin resolves to this file.
import { readFileSync, writeFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";

const distDir = join(dirname(fileURLToPath(import.meta.url)), "..", "dist");
const html = readFileSync(join(distDir, "index.html"), "utf8");

const checkinHtml = html
  .replace(/<title>.*?<\/title>/, "<title>仁愛水電打卡</title>")
  .replace(
    /<meta name="description"[^>]*>\s*<meta property="og:title"[^>]*>\s*<meta property="og:description"[^>]*>\s*<meta property="og:type"[^>]*>/,
    [
      '<meta name="description" content="仁愛水電 出工打卡" />',
      '<meta property="og:title" content="仁愛水電打卡" />',
      '<meta property="og:description" content="出工打卡，記錄上下班時間" />',
      '<meta property="og:type" content="website" />',
    ].join("\n    ")
  );

writeFileSync(join(distDir, "checkin.html"), checkinHtml);
console.log("postbuild: wrote dist/checkin.html");
