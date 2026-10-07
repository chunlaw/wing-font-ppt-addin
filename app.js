import * as hb from "./vendor/harfbuzzjs/index.mjs";
import { applyChanges, clusters, loadFont, stripChanges } from "./convert.js";
import { setLanguage, t } from "./i18n.js";

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
  if (!out) throw new Error(t("woff2Failed"));
  return out;
};

// When the published copy of a font was last changed (null if it can't be
// told: offline, not an official font, ...). Last-Modified is one of the few
// headers a cross-origin page may read.
const published = (url) =>
  fetch(url, { method: "HEAD", cache: "no-store" })
    .then((r) => (r.ok ? r.headers.get("Last-Modified") : null))
    .catch(() => null);

async function fontBytes(name) {
  const key = cacheKey(name);
  const cache = await globalThis.caches?.open("wing-fonts");
  let res = await cache?.match(key);
  if (res && !name.startsWith("local:")) {
    // A remembered download goes stale when the font is republished: the
    // user installs the new build and the pane would keep converting with
    // the old one. Ask the server whether it changed and fetch again if so.
    const now = await published(key);
    if (now && now !== res.headers.get("Last-Modified")) res = null;
  }
  if (!res) {
    if (name.startsWith("local:")) throw new Error(t("repickLocal", name.slice(6)));
    res = await fetch(key, { cache: "no-store" });
    if (!res.ok) {
      showFontInfo(t("notFound", name, res.status) + "\n" + t("loadLocal"), true);
      throw new Error(t("notLoaded", name));
    }
    await cache?.put(key, res.clone());
  }
  const bytes = new Uint8Array(await res.arrayBuffer());
  return String.fromCharCode(...bytes.subarray(0, 4)) === "wOF2" ? decompress(bytes) : bytes;
}

// The add-in can't read installed fonts: it converts with its own copy,
// PowerPoint renders with the installed one. Selector numbering can differ
// between builds, so show which build was loaded and, when that can't be
// told (no build date, or no copy to download), ask for the installed file.
const showFontInfo = (msg, bad = false) => {
  $("fontinfo").textContent = msg;
  $("fontinfo").className = bad ? "bad" : "";
  if (bad) $("fontopts").open = true; // the fix is in there: show it
};
async function describeFont({ name, font }) {
  const local = name.startsWith("local:");
  const built = font.version.match(/Wing Font (\d{4}-\d{2}-\d{2})/)?.[1];
  const lines = [t("infoFont", local ? t("infoLocal", name.slice(6)) : name), t("infoVersion", font.version)];
  let bad = false;
  if (!local) {
    lines.push(built ? t("checkVersion") : t("noBuildDate") + "\n" + t("loadLocal"));
    bad = !built;
  } else if (built && font.family) {
    // A picked file is remembered under its file name. If the official font
    // of that family has been republished since this copy was built, the
    // user may well have installed the newer one: say so.
    const now = await published(CDN + encodeURIComponent(font.family) + ".woff2");
    const newer = now && new Date(now).toISOString().slice(0, 10);
    if (newer && newer > built) {
      lines.push(t("newerBuild", newer, built));
      bad = true;
    }
  }
  showFontInfo(lines.join("\n"), bad);
}

let current; // { name, font, sha, css }
async function useFont(name, bytes) {
  if (current?.name === name) return current;
  say(t("loading", name));
  bytes ??= await fontBytes(name);
  const digest = await crypto.subtle.digest("SHA-256", bytes);
  const sha = [...new Uint8Array(digest)].map((b) => b.toString(16).padStart(2, "0")).join("").slice(0, 16);
  const css = "wf-" + sha;
  document.fonts.add(await new FontFace(css, bytes).load());
  $("sandbox").style.fontFamily = css;
  current = { name, font: loadFont(hb, bytes), sha, css };
  await describeFont(current);
  // The deck remembers which build converted it; warn when it differs.
  const rec = settings()?.get("wingFont");
  say(rec && rec.name === name && rec.sha !== sha
    ? t("versionChanged", name)
    : t("loaded", name), rec && rec.name === name && rec.sha !== sha);
  return current;
}

// Manual choice wins, remembered per document family name.
const overrideKey = (family) => "font:" + (family || "");
function fontFor(family) {
  // PowerPoint reports no font name when the selection mixes fonts (say, a
  // title and a body, or text that is partly Latin-font). Keep using the
  // font already loaded rather than refusing: it is shown in the Font card.
  const name = $("font").value.trim() || localStorage.getItem(overrideKey(family)) || family || current?.name;
  if (!name) throw new Error(t("noFont"));
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
// withSelection(fn, wholeIfCaret): fn(text, family) → [{ start, end, replacement }];
// the adapter applies the edits back-to-front so offsets stay valid and
// untouched text keeps its formatting.

const settings = () => globalThis.Office?.context?.document?.settings;

const adapters = {
  async powerpoint(fn, wholeIfCaret) {
    // With only a caret in the text, getSelectedTextRange() does not return
    // an empty range. Measured in PowerPoint 16 (API 1.8):
    //   Windows — the word around the caret (銀行攞錢), caret position unknown;
    //   Mac     — the single character right after the caret.
    // The common API returns "" for a caret on both, which is how a caret
    // is told apart from a real selection of the same text.
    const caret = await new Promise((resolve) =>
      Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
        (r) => resolve(r.status === "succeeded" && r.value === "")));
    await PowerPoint.run(async (ctx) => {
      let ranges;
      try {
        let r = ctx.presentation.getSelectedTextRange();
        if (caret) {
          const all = r.getParentTextFrame().textRange;
          if (wholeIfCaret) {
            r = all; // Convert / Restore: the whole text box
          } else {
            // Picker: if only one character came back (Mac), also offer the
            // one before the caret — that's the one just typed.
            r.load("start,length");
            all.load("text");
            await ctx.sync();
            const cs = clusters(all.text);
            const i = cs.findIndex((c) => c.start <= r.start && r.start < c.end);
            if (i >= 0 && r.length <= cs[i].end - cs[i].start) {
              const prev = cs[i - 1];
              const first = prev && !/[\r\n\v]/.test(prev.char) ? prev : cs[i];
              r = all.getSubstring(first.start, cs[i].end - first.start);
            }
          }
        }
        r.load("text");
        r.font.load("name");
        await ctx.sync();
        ranges = [r];
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
        for (const c of (await fn(r.text, r.font.name || "")).reverse()) {
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
// Actions run one at a time, in order. A click must never be dropped just
// because a picker refresh (fired by every selection change) is in flight.
let queue = Promise.resolve();
// `wholeIfCaret`: with just a caret in the text, act on the whole text box
// (Convert, Restore) rather than on the word around the caret (the picker).
const run = (fn, { quiet = false, wholeIfCaret = false } = {}) =>
  (queue = queue
    .then(() => withSelection(fn, wholeIfCaret))
    .catch((e) => quiet || say(e.message || String(e), true)));

// ── Actions ──────────────────────────────────────────────────────────

// PowerPoint reports no font name ("" here) when the selected text mixes
// fonts; other hosts pass null because they can't tell. Conversion then
// uses one font for all of it, so say so under the Convert button.
const noteFonts = (text, family) => {
  $("mixed").hidden = !(text && family === "");
};

$("convert").onclick = () =>
  run(async (text, family) => {
    if (!text) throw new Error(t("noSelection"));
    noteFonts(text, family);
    const f = await fontFor(family);
    const r = f.font.convert(text);
    if (!r.cmapOk) throw new Error(t("selfCheckFailed"));
    const notes = [t("converted", r.changes.length)];
    if (r.failed.length) {
      notes.push(t("failed", r.failed.length, r.failed.map((x) => t("quote", text.slice(x.start, x.end))).join("")));
    }
    if (!r.gsubOk) notes.push(t("gsubWarning"));
    if (r.changes.length && settings()) {
      settings().set("wingFont", { name: f.name, sha: f.sha });
      settings().saveAsync();
    }
    say(notes.join("\n"), r.failed.length > 0 || !r.gsubOk);
    return r.changes;
  }, { wholeIfCaret: true });

$("restore").onclick = () =>
  run(async (text) => {
    const changes = stripChanges(text);
    say(t("restored", changes.length));
    return changes;
  }, { wholeIfCaret: true });

// Picker: one row per character of a short selection, each reading drawn
// with the loaded font itself; clicking one swaps that character. A bare
// caret counts as selecting the word around it (see the adapter), so
// putting the cursor in a word lists the readings of each of its characters.
const PICKER_MAX = 12;
let pickerQueued = false; // collapse bursts of selection-change events
async function refreshPicker() {
  if (pickerQueued) return;
  pickerQueued = true;
  let rows = []; // stays empty if the selection can't be read at all
  let hint = t("pickEmpty");
  await run(async (text, family) => {
    pickerQueued = false;
    noteFonts(text, family);
    const chars = clusters(text).filter((c) => c.char.trim());
    if (!chars.length) return [];
    if (chars.length > PICKER_MAX) return (hint = t("pickTooMany", PICKER_MAX)), [];
    const f = await fontFor(family).catch(() => null);
    if (!f) return [];
    for (const c of chars) {
      const seen = new Set();
      const list = f.font.variants(c.char).filter((v) => !seen.has(v.gid) && seen.add(v.gid));
      if (list.length < 2) continue;
      const row = document.createElement("div");
      row.className = "row";
      row.append(...list.map((v) => {
        const b = document.createElement("button");
        b.textContent = v.text;
        b.setAttribute("aria-pressed", String(v.text === c.text)); // the one in the document now
        b.style.fontFamily = f.css;
        // Swap only if the selection still holds the same text: never
        // overwrite something else the user selected in the meantime.
        b.onclick = () => run(async (now) =>
          now === text ? [{ start: c.start, end: c.end, replacement: v.text }] : []).then(refreshPicker);
        return b;
      }));
      rows.push(row);
    }
    if (!rows.length) hint = t("pickNone");
    return [];
  }, { quiet: true }); // quiet: selecting a picture is not an error worth showing
  pickerQueued = false;
  $("variants").dataset.empty = hint;
  $("variants").replaceChildren(...rows);
}

// ── Boot ─────────────────────────────────────────────────────────────

// office.js may be missing (offline, plain browser): fall back to the sandbox.
const info = await (globalThis.Office?.onReady() ?? {});
// Follow the Office UI language; in a plain browser, ?lang= or the browser's.
setLanguage(info.host ? Office.context.displayLanguage
  : new URLSearchParams(location.search).get("lang") || navigator.language);
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
}
const params = new URLSearchParams(location.search);
$("font").value = params.get("font") ?? "";
say(t("ready"));
// Development only: test/ isn't part of the published site.
if (params.get("selftest")) (await import("./test/office/selftest.js")).run(params.get("selftest"));
