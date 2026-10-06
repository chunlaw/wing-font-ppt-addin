// End-to-end check inside a real Office host. Loaded only when the pane URL
// has ?selftest=<report URL> (see test/office/README.md). It drives the
// pane's own buttons against a text box it adds to slide 1, then POSTs what
// it saw. Nothing here ships: build.mjs doesn't copy test/ into dist/.

const $ = (id) => document.getElementById(id);
const wait = (ms) => new Promise((r) => setTimeout(r, ms));
const cps = (s) => Array.from(s, (c) => c.codePointAt(0).toString(16)).join(" ");

// Wait until the status line stops changing and isn't a "loading…" message.
async function settle() {
  let prev;
  for (let i = 0; i < 240; i++) {
    await wait(250);
    const s = $("status").textContent;
    if (s === prev && !s.endsWith("…")) return s;
    prev = s;
  }
  return prev;
}

export async function run(reportUrl) {
  const p = new URLSearchParams(location.search);
  const family = p.get("family");
  const RAW = "我去銀行攞錢，行路返屋企。行２ 行３ 行０ 行０ｚａａ１ 😀行２";
  const log = { ua: navigator.userAgent, raw: RAW, steps: {} };
  const send = () => fetch(reportUrl, { method: "POST", body: JSON.stringify(log, null, 1) }).catch(() => {});
  try {
    log.host = String(Office.context.host);
    log.platform = String(Office.context.platform);
    log.displayLanguage = Office.context.displayLanguage;
    log.contentLanguage = Office.context.contentLanguage;
    log.api = ["1.1", "1.2", "1.3", "1.4", "1.5", "1.6", "1.7", "1.8"].filter((v) =>
      Office.context.requirements.isSetSupported("PowerPointApi", v));
    log.ui = { lang: document.documentElement.lang, convert: $("convert").textContent };

    const box = (ctx) => ctx.presentation.slides.getItemAt(0).shapes.getItem(log.shapeId).textFrame.textRange;
    const read = async (what) => {
      const out = {};
      await PowerPoint.run(async (ctx) => {
        const tr = box(ctx);
        tr.load("text");
        tr.font.load("name");
        await ctx.sync();
        out.text = cps(tr.text);
        log.len = tr.text.length;
        out.fontName = tr.font.name;
        // Formatting probe: is 路 still red + bold, 我 still not?
        for (const [key, ch] of [["lou", "路"], ["ngo", "我"]]) {
          const f = tr.getSubstring(tr.text.indexOf(ch), 1).font;
          f.load("color,bold,name");
          await ctx.sync();
          out[key] = { color: f.color, bold: f.bold, name: f.name };
        }
      });
      out.status = $("status").textContent;
      log.steps[what] = out;
    };

    // Select a text range; if the host refuses, select the whole shape
    // instead (the pane handles both) and note that it happened.
    const select = async (what, pick) => {
      try {
        await PowerPoint.run(async (ctx) => {
          pick(box(ctx)).setSelected();
          await ctx.sync();
        });
        log.selected[what] = "text";
      } catch (e) {
        await PowerPoint.run(async (ctx) => {
          ctx.presentation.slides.getItemAt(0).setSelectedShapes([log.shapeId]);
          await ctx.sync();
        });
        log.selected[what] = "shape (text selection failed: " + e.message + ")";
      }
      await wait(1500);
    };
    log.selected = {};

    // 1. Set up: a text box in the wing font, with 行路 red and bold.
    await PowerPoint.run(async (ctx) => {
      const shape = ctx.presentation.slides.getItemAt(0).shapes.addTextBox(RAW, { left: 30, top: 30, width: 660, height: 320 });
      shape.name = "wingfont-selftest";
      const tr = shape.textFrame.textRange;
      tr.font.name = family;
      tr.font.size = 30;
      const hot = tr.getSubstring(RAW.indexOf("行路"), 2).font;
      hot.color = "#FF0000";
      hot.bold = true;
      shape.load("id");
      await ctx.sync();
      log.shapeId = shape.id;
    });
    await select("setup", (tr) => tr.getSubstring(0, RAW.length));
    await read("setup");

    // 2. Convert the selection.
    $("convert").click();
    await settle();
    await read("converted");
    log.fontinfo = $("fontinfo").textContent;
    log.settings = Office.context.document.settings.get("wingFont");

    // 3. Picker: select the 行 after 。 (with its selector), expect tiles, click the first.
    let pick;
    await PowerPoint.run(async (ctx) => {
      const tr = box(ctx);
      tr.load("text");
      await ctx.sync();
      const i = tr.text.indexOf("行", tr.text.indexOf("。"));
      pick = { i, len: /^.[\u{E0100}-\u{E01EF}]/u.test(tr.text.slice(i)) ? 3 : 1 };
    });
    log.pickSelection = pick;
    await select("pick", (tr) => tr.getSubstring(pick.i, pick.len));
    await wait(4000); // DocumentSelectionChanged → picker
    const tiles = [...document.querySelectorAll("#variants button")];
    log.tiles = tiles.map((b) => ({ text: cps(b.textContent), pressed: b.getAttribute("aria-pressed") }));
    if (tiles.length) {
      tiles[0].click();
      await wait(3000);
    }
    await read("picked");

    // 4. Restore everything.
    await select("restore", (tr) => tr.getSubstring(0, log.len));
    $("restore").click();
    await settle();
    await read("restored");

    // 5. Leave the slide showing the converted text, for a screenshot.
    await PowerPoint.run(async (ctx) => {
      box(ctx).text = RAW;
      await ctx.sync();
    });
    await select("final", (tr) => tr.getSubstring(0, RAW.length));
    $("convert").click();
    await settle();
    await read("final");
    log.done = true;
  } catch (e) {
    log.error = String(e) + (e.debugInfo ? " " + JSON.stringify(e.debugInfo) : "") + "\n" + (e.stack || "");
  }
  await send();
}
