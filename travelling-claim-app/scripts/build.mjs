import fs from "node:fs/promises";

const required = [
  "index.html",
  "manifest.json",
  "service-worker.js",
  "src/main.js",
  "src/lib/xlsxTemplatePatcher.js",
  "src/styles/app.css",
  "config/templateMapping.js",
  "templates/Travelling claim form - updated version.xlsx"
];

for (const file of required) {
  await fs.access(new URL(`../${file}`, import.meta.url));
}

console.log("Static app build check passed. Serve the project root or deploy these files.");
