// After `vite build`, duplicate dist/index.html into one HTML file per
// simplified front-end route, each with its own <title>/Open Graph tags —
// so sharing e.g. the /checkin or /tasks link (LINE, etc.) shows a title
// specific to that page instead of the main app's title.
// Each is a straight copy of the built index.html (same hashed script tag),
// so it always stays in sync with whatever main.jsx built to — nothing here
// references file names that could go stale between builds.
// Requires firebase.json's "cleanUrls": true so e.g. /checkin resolves to
// checkin.html.
import { readFileSync, writeFileSync } from "node:fs";
import { fileURLToPath } from "node:url";
import { dirname, join } from "node:path";

const distDir = join(dirname(fileURLToPath(import.meta.url)), "..", "dist");
const html = readFileSync(join(distDir, "index.html"), "utf8");

const pages = [
  { file: "checkin.html", title: "仁愛水電打卡", description: "出工打卡，記錄上下班時間" },
  { file: "tasks.html", title: "仁愛水電待確認事項", description: "工地每日待確認事項清單" },
];

for (const { file, title, description } of pages) {
  const out = html
    .replace(/<title>.*?<\/title>/, `<title>${title}</title>`)
    .replace(
      /<meta name="description"[^>]*>\s*<meta property="og:title"[^>]*>\s*<meta property="og:description"[^>]*>\s*<meta property="og:type"[^>]*>/,
      [
        `<meta name="description" content="${description}" />`,
        `<meta property="og:title" content="${title}" />`,
        `<meta property="og:description" content="${description}" />`,
        `<meta property="og:type" content="website" />`,
      ].join("\n    ")
    );
  writeFileSync(join(distDir, file), out);
  console.log(`postbuild: wrote dist/${file}`);
}
