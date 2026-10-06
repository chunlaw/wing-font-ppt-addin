// node build.mjs vendor — copy the browser-loadable library files out of
//                         node_modules into ./vendor (runs on npm install).
// node build.mjs        — also assemble the static site into ./dist.
import { cpSync, mkdirSync, rmSync } from "node:fs";

const VENDOR = {
  "vendor/harfbuzzjs/index.mjs": "node_modules/harfbuzzjs/dist/index.mjs",
  "vendor/harfbuzzjs/harfbuzz.js": "node_modules/harfbuzzjs/dist/harfbuzz.js",
  "vendor/harfbuzzjs/harfbuzz.wasm": "node_modules/harfbuzzjs/dist/harfbuzz.wasm",
  "vendor/wawoff2/decompress_binding.js": "node_modules/wawoff2/build/decompress_binding.js",
};

rmSync("vendor", { recursive: true, force: true });
mkdirSync("vendor/harfbuzzjs", { recursive: true });
mkdirSync("vendor/wawoff2");
for (const [to, from] of Object.entries(VENDOR)) cpSync(from, to);

if (process.argv[2] !== "vendor") {
  rmSync("dist", { recursive: true, force: true });
  for (const f of ["index.html", "app.js", "convert.js", "i18n.js", "assets", "vendor"]) {
    cpSync(f, "dist/" + f, { recursive: true });
  }
}
