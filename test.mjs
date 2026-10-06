// node test.mjs — checks the converter against test/WingSmall.ttf.
// 行 has three readings; selectors follow the annotation strings' sort
// order, not their weight rank: VS17 = haang4, VS18 = hang4, VS19 = hong4
// (the default), VS256 = bare. By weight, 行２ = hang4 and 行３ = haang4.
// DIY marks follow CSV row order: zaa1 = U+E000, hong2 = U+E000 + VS17.
// Word rule 銀行 → hong4.
import assert from "node:assert/strict";
import { readFileSync } from "node:fs";
import { createRequire } from "node:module";
import * as hb from "harfbuzzjs";
import { applyChanges, loadFont, stripChanges } from "./convert.js";

const VS17 = "\u{E0100}", VS18 = "\u{E0101}", VS19 = "\u{E0102}", BARE = "\u{E01EF}";
const ZAA1 = "\u{E000}", HONG2 = "\u{E000}\u{E0100}";
const f = loadFont(hb, readFileSync("test/WingSmall.ttf"));
const conv = (s) => {
  const r = f.convert(s);
  assert(r.cmapOk && r.gsubOk, `self-check failed for ${s}`);
  assert.deepEqual(r.failed, []);
  return r.text;
};

assert.equal(conv("行"), "行");
assert.equal(conv("銀行"), "銀行"); // default reading: nothing to pin
assert.equal(conv("行２"), "行" + VS18);
assert.equal(conv("行３"), "行" + VS17);
assert.equal(conv("行０"), "行" + BARE);
assert.equal(conv("行０ｚａａ１"), "行" + BARE + ZAA1);
assert.equal(conv("行０ｈｏｎｇ２"), "行" + BARE + HONG2);
assert.equal(conv("行" + VS19), "行" + VS19); // explicit default selector: left as typed
assert.equal(conv("行" + VS17), "行" + VS17); // idempotent
assert.equal(conv("abc 行２\r一行３\v銀"), `abc 行${VS18}\r一行${VS17}\v銀`);
assert.equal(conv("😀行２"), "😀行" + VS18); // not in font: left alone

// Offsets are UTF-16 code units into the original text.
assert.deepEqual(f.convert("😀行２").changes, [{ start: 2, end: 4, replacement: "行" + VS18 }]);

// cmap-only rendering of the output equals the fully shaped input.
const gids = (s) => f.shape(s).map((g) => g.gid);
assert.deepEqual(f.cmapOnly(conv("行０ｚａａ１")), gids("行０ｚａａ１"));

assert.deepEqual(f.variants("行").map((v) => v.text), ["行", "行" + VS17, "行" + VS18, "行" + VS19, "行" + BARE]);
const marked = "行" + BARE + ZAA1 + "銀";
assert.equal(applyChanges(marked, stripChanges(marked)), "行銀");
const r = f.convert("a行２b行３");
assert.equal(applyChanges("a行２b行３", r.changes), r.text);

// woff2 → sfnt gives the same result as the ttf.
const woff2 = createRequire(import.meta.url)("wawoff2/build/decompress_binding.js");
await new Promise((r) => (woff2.calledRun ? r() : (woff2.onRuntimeInitialized = r)));
const sfnt = woff2.decompress(readFileSync("test/WingSmall.woff2"));
assert.equal(loadFont(hb, sfnt).convert("行２").text, "行" + VS18);

console.log("ok");
