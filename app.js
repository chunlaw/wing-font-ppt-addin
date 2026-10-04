import * as hb from "./vendor/harfbuzzjs/index.mjs";
import { applyChanges, loadFont, stripChanges } from "./convert.js";

const CDN = "https://wing-font.chunlaw.io/fonts/";
const $ = (id) => document.getElementById(id);
const say = (msg, bad = false) => {
  $("status").textContent = msg;
  $("status").className = bad ? "bad" : "";
};

// ── Fonts ────────────────────────────────────────────────────────────
// A font "name" is an official family name (fetched from the CDN), a URL
// (anything with a slash), or "local:<file name>" for a user-picked file.
// Everything is kept in the Cache API so it's downloaded / picked once.

const cacheKey = (name) =>
  name.startsWith("local:") ? "local/" + encodeURIComponent(name.slice(6))
  : name.includes("/") ? name
  : CDN + encodeURIComponent(name) + ".woff2";

let woff2; // wawoff2's emscripten module, loaded on first use
const decompress = async (bytes) => {
  woff2 ??= new Promise((resolve, reject) => {
    window.Module = { onRuntimeInitialized: () => resolve(window.Module) };
    const s = document.createElement("script");
    s.src = "vendor/wawoff2/decompress_binding.js";
    s.onerror = reject;
    document.head.append(s);
  });
  const out = (await woff2).decompress(bytes);
  if (!out) throw new Error("woff2 解壓失敗");
  return out;
};

async function fontBytes(name) {
  const key = cacheKey(name);
  const cache = await globalThis.caches?.open("wing-fonts");
  let res = await cache?.match(key);
  if (!res) {
    if (name.startsWith("local:")) throw new Error(`請重新揀本機字型檔「${name.slice(6)}」`);
    res = await fetch(key);
    if (!res.ok) throw new Error(`搵唔到字型「${name}」（${res.status}）。請填官方字型名，或者揀本機字型檔。`);
    await cache?.put(key, res.clone());
  }
  const bytes = new Uint8Array(await res.arrayBuffer());
  return String.fromCharCode(...bytes.subarray(0, 4)) === "wOF2" ? decompress(bytes) : bytes;
}

let current; // { name, font, sha, css }
async function useFont(name, bytes) {
  if (current?.name === name) return current;
  say(`載入字型 ${name}…`);
  bytes ??= await fontBytes(name);
  const digest = await crypto.subtle.digest("SHA-256", bytes);
  const sha = [...new Uint8Array(digest)].map((b) => b.toString(16).padStart(2, "0")).join("").slice(0, 16);
  const css = "wf-" + sha;
  document.fonts.add(await new FontFace(css, bytes).load());
  $("sandbox").style.fontFamily = css;
  current = { name, font: loadFont(hb, bytes), sha, css };
  // The deck remembers which build converted it; warn when it differs.
  const rec = settings()?.get("wingFont");
  say(rec && rec.name === name && rec.sha !== sha
    ? `注意：字型「${name}」同呢份文件上次轉換時用嘅版本唔同，舊文字嘅讀音可能有變。`
    : `字型 ${name} 已載入。`, rec && rec.name === name && rec.sha !== sha);
  return current;
}

// Manual choice wins, remembered per document family name.
const overrideKey = (family) => "font:" + (family || "");
function fontFor(family) {
  const name = $("font").value.trim() || localStorage.getItem(overrideKey(family)) || family;
  if (!name) throw new Error("唔知用邊個字型：請喺上面填字型名，或者揀本機字型檔。");
  if ($("font").value.trim()) localStorage.setItem(overrideKey(family), name);
  return useFont(name);
}

$("file").onchange = async () => {
  const file = $("file").files[0];
  if (!file) return;
  const name = "local:" + file.name;
  const cache = await globalThis.caches?.open("wing-fonts");
  await cache?.put(cacheKey(name), new Response(file)); // stays on this machine
  current = null;
  $("font").value = name;
  let bytes = new Uint8Array(await file.arrayBuffer());
  if (file.name.endsWith(".woff2")) bytes = await decompress(bytes);
  await useFont(name, bytes).catch((e) => say(e.message, true));
};

// ── Host adapters ────────────────────────────────────────────────────
// withSelection(fn): fn(text, family) → [{ start, end, replacement }];
// the adapter applies the edits back-to-front so offsets stay valid and
// untouched text keeps its formatting.

const settings = () => globalThis.Office?.context?.document?.settings;

const adapters = {
  async powerpoint(fn) {
    await PowerPoint.run(async (ctx) => {
      let ranges;
      try {
        ranges = [ctx.presentation.getSelectedTextRange()];
        ranges[0].load("text");
        ranges[0].font.load("name");
        await ctx.sync();
      } catch {
        // No text selection: fall back to whole selected text boxes.
        const shapes = ctx.presentation.getSelectedShapes();
        shapes.load("items");
        await ctx.sync();
        ranges = shapes.items.map((s) => s.textFrame.textRange);
        for (const r of ranges) r.load("text"), r.font.load("name");
        await ctx.sync();
      }
      for (const r of ranges) {
        for (const c of (await fn(r.text, r.font.name)).reverse()) {
          r.getSubstring(c.start, c.end - c.start).text = c.replacement;
        }
      }
      await ctx.sync();
    });
  },
  // Word, and PowerPoint older than PowerPointApi 1.5.
  // ponytail: replaces the whole selection, so mixed formatting inside it
  // collapses and the font can't be auto-detected. Use Word.run ranges if
  // the Word version (M5) needs better.
  async common(fn) {
    const doc = Office.context.document;
    const call = (method, ...args) =>
      new Promise((resolve, reject) =>
        doc[method](...args, (r) => (r.status === "succeeded" ? resolve(r.value) : reject(new Error(r.error.message)))));
    const text = await call("getSelectedDataAsync", Office.CoercionType.Text);
    const changes = await fn(text, null);
    if (changes.length) await call("setSelectedDataAsync", applyChanges(text, changes));
  },
  // Plain browser: the sandbox textarea stands in for the document.
  async standalone(fn) {
    const box = $("sandbox");
    const [s, e] = box.selectionStart === box.selectionEnd ? [0, box.value.length] : [box.selectionStart, box.selectionEnd];
    const changes = await fn(box.value.slice(s, e), null);
    if (changes.length) box.setRangeText(applyChanges(box.value.slice(s, e), changes), s, e, "select");
  },
};

let withSelection = adapters.standalone;
let busy = false;
async function run(fn) {
  if (busy) return;
  busy = true;
  try {
    await withSelection(fn);
  } catch (e) {
    say(e.message || String(e), true);
  } finally {
    busy = false;
  }
}

// ── Actions ──────────────────────────────────────────────────────────

$("convert").onclick = () =>
  run(async (text, family) => {
    if (!text) throw new Error("請先選取文字或者文字方塊。");
    const f = await fontFor(family);
    const r = f.font.convert(text);
    if (!r.cmapOk) throw new Error("自我驗證失敗：轉換結果同字型排版唔一致，冇改動文字。");
    const notes = [`已轉換 ${r.changes.length} 處。`];
    if (r.failed.length) {
      notes.push(`有 ${r.failed.length} 處冇 cmap 路徑，保留原文：` +
        r.failed.map((x) => `「${text.slice(x.start, x.end)}」`).join(""));
    }
    if (!r.gsubOk) notes.push("注意：轉換後喺會做 GSUB 嘅程式（Word、網頁）讀音可能唔同。");
    if (r.changes.length && settings()) {
      settings().set("wingFont", { name: f.name, sha: f.sha });
      settings().saveAsync();
    }
    say(notes.join("\n"), r.failed.length > 0 || !r.gsubOk);
    return r.changes;
  });

$("restore").onclick = () =>
  run(async (text) => {
    const changes = stripChanges(text);
    say(`已還原 ${changes.length} 處。`);
    return changes;
  });

// Picker: when exactly one character is selected, draw each of its cmap-14
// variants with the loaded font itself; clicking one swaps the selection.
const ONE_CHAR = /^([^\u{E000}-\u{F8FF}\u{E0100}-\u{E01EF}])[\u{E000}-\u{F8FF}\u{E0100}-\u{E01EF}]*$/u;
async function refreshPicker() {
  const box = $("variants");
  await run(async (text, family) => {
    const char = text.match(ONE_CHAR)?.[1];
    const f = char && (await fontFor(family).catch(() => null));
    const seen = new Set();
    const list = f ? f.font.variants(char).filter((v) => !seen.has(v.gid) && seen.add(v.gid)) : [];
    box.replaceChildren(...(list.length > 1 ? list : []).map((v) => {
      const b = document.createElement("button");
      b.textContent = v.text;
      b.style.fontFamily = f.css;
      b.onclick = () => run(async (t) => [{ start: 0, end: t.length, replacement: v.text }]);
      return b;
    }));
    return [];
  });
}

// ── Boot ─────────────────────────────────────────────────────────────

// office.js may be missing (offline, plain browser): fall back to the sandbox.
const info = await (globalThis.Office?.onReady() ?? {});
if (info.host === "PowerPoint" && Office.context.requirements.isSetSupported("PowerPointApi", "1.5")) {
  withSelection = adapters.powerpoint;
} else if (info.host) {
  withSelection = adapters.common;
}
if (info.host) {
  Office.context.document.addHandlerAsync(Office.EventType.DocumentSelectionChanged, refreshPicker);
} else {
  $("sandbox").hidden = false;
  $("sandbox").onselect = $("sandbox").onkeyup = $("sandbox").onmouseup = refreshPicker;
  $("font").value = new URLSearchParams(location.search).get("font") ?? "";
}
say("準備好。");
