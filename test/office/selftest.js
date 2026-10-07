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

    // Mixed-font matrix (?mixed=1): in a text box of its own, give part of
    // the text another font in many ways and record what the host reports
    // and whether the pane's warning agrees with it.
    if (p.get("mixed")) {
      const TEXT = "Hello 你好嗎 123\r行路";
      // A few fonts present on Windows, on Mac, or both; a font the machine
      // lacks is simply not applied and shows up as such in the log.
      const FONTS = ["Microsoft JhengHei", "SimSun", "Songti SC", "STKaiti", "Arial", "Times New Roman"];
      const PARTS = { cjk: [6, 2], latin: [0, 5], paragraph2: [14, 2] };
      let id2;
      await PowerPoint.run(async (ctx) => {
        const shape = ctx.presentation.slides.getItemAt(0).shapes.addTextBox(TEXT, { left: 30, top: 360, width: 660, height: 150 });
        shape.textFrame.textRange.font.name = family;
        shape.load("id");
        await ctx.sync();
        id2 = shape.id;
      });
      const tr2 = (ctx) => ctx.presentation.slides.getItemAt(0).shapes.getItem(id2).textFrame.textRange;
      const pick = async (start, len) => {
        await PowerPoint.run(async (ctx) => {
          tr2(ctx).getSubstring(start, len).setSelected();
          await ctx.sync();
        });
        await wait(1200);
      };
      log.matrix = [];
      const check = async (label, setup, selectShape) => {
        const m = { label };
        try {
          await PowerPoint.run(async (ctx) => {
            tr2(ctx).font.name = family; // reset
            await ctx.sync();
          });
          const part = await setup();
          if (!selectShape) await pick(15, 1); // move away so the next selection is a change
          if (selectShape) {
            await PowerPoint.run(async (ctx) => {
              ctx.presentation.slides.getItemAt(0).setSelectedShapes([id2]);
              await ctx.sync();
            });
          } else {
            await pick(0, TEXT.length);
          }
          await wait(2200);
          await PowerPoint.run(async (ctx) => {
            const whole = tr2(ctx);
            whole.font.load("name");
            const sub = part ? whole.getSubstring(part[0], part[1]) : whole;
            sub.font.load("name");
            await ctx.sync();
            m.part = sub.font.name;
            m.whole = whole.font.name;
          });
          m.selectionIs = await new Promise((res) => Office.context.document.getSelectedDataAsync(
            Office.CoercionType.Text, (r) => res(r.status === "succeeded" ? JSON.stringify(r.value) : "ERR")));
          m.applied = !!part && m.part !== family; // did the host take the second font at all?
          m.hostSaysMixed = !m.whole;
          m.warning = !$("mixed").hidden;
          if (selectShape) {
            // Convert is what a user presses with a text box selected.
            $("convert").click();
            await settle();
            m.warningAfterConvert = !$("mixed").hidden;
            m.status = $("status").textContent;
          }
        } catch (e) {
          m.error = e.message;
        }
        log.matrix.push(m);
      };
      // Whole-shape selection (not a text selection). These go first: once
      // text has been selected, setSelectedShapes no longer takes effect.
      for (const font of ["Microsoft JhengHei", "Songti SC", "Arial"]) {
        await check(`${font} / cjk / shape selected`, async () => {
          await PowerPoint.run(async (ctx) => {
            tr2(ctx).getSubstring(6, 2).font.name = font;
            await ctx.sync();
          });
          return [6, 2];
        }, true);
      }
      await check("uniform / shape selected", async () => null, true);
      await check("uniform", async () => null);
      for (const font of FONTS) {
        for (const [name, part] of Object.entries(PARTS)) {
          await check(`${font} / ${name}`, async () => {
            await PowerPoint.run(async (ctx) => {
              tr2(ctx).getSubstring(part[0], part[1]).font.name = font;
              await ctx.sync();
            });
            return part;
          });
        }
      }
      log.done = "mixed";
      return send();
    }

    // Probe mode (?probe=1): log what the selection APIs report while
    // something outside (a person, or SendKeys) moves the caret around.
    if (p.get("probe")) {
      await select("probe", (tr) => tr.getSubstring(3, 2)); // 行攞
      log.probe = [];
      // ?probe=api: no keyboard available, so collapse the selection through
      // the API instead (a zero-length range) and report what comes back.
      const moves = p.get("probe") === "api" ? { 6: 4, 30: 9, 54: 0, 78: 20 } : {};
      let last = "";
      for (let i = 0; i < 100; i++) {
        const o = {};
        if (i in moves) {
          await PowerPoint.run(async (ctx) => {
            box(ctx).getSubstring(moves[i], 0).setSelected();
            await ctx.sync();
          }).catch((e) => (o.moveError = e.message));
        }
        try {
          await PowerPoint.run(async (ctx) => {
            const r = ctx.presentation.getSelectedTextRange();
            r.load("text,start,length");
            await ctx.sync();
            Object.assign(o, { text: r.text, start: r.start, length: r.length });
          });
        } catch (e) {
          o.error = e.message;
        }
        o.common = await new Promise((res) => Office.context.document.getSelectedDataAsync(
          Office.CoercionType.Text, (r) => res(r.status === "succeeded" ? r.value : "ERR " + r.error.message)));
        o.rows = [...document.querySelectorAll("#variants .row")].map((r) => r.firstChild.textContent[0]).join("");
        const key = JSON.stringify(o);
        if (key !== last) log.probe.push({ t: i, ...o });
        last = key;
        await wait(500);
      }
      // Whatever the caret is now, Convert must act on the whole text box.
      $("convert").click();
      await settle();
      await read("probeConvert");
      log.done = "probe";
      return send();
    }

    // 1b. A short multi-character selection lists one row per character;
    // clicking a tile changes just that character, clicking back undoes it.
    try {
      await select("word", (tr) => tr.getSubstring(2, 4)); // 銀行攞錢
      await wait(3500);
      const rows = [...document.querySelectorAll("#variants .row")];
      log.wordRows = rows.map((r) => [...r.children].map((b) => cps(b.textContent) + (b.getAttribute("aria-pressed") === "true" ? "*" : "")));
      const row = rows[1]; // 行
      const was = [...row.children].findIndex((b) => b.getAttribute("aria-pressed") === "true");
      const other = was === 1 ? 2 : 1;
      row.children[other].click();
      await wait(3500);
      await read("wordPick");
      const again = document.querySelectorAll("#variants .row")[1];
      log.wordPickPressed = [...again.children].findIndex((b) => b.getAttribute("aria-pressed") === "true") === other;
      again.children[was].click();
      await wait(3500);
      await read("wordUndo");
      await select("setup2", (tr) => tr.getSubstring(0, RAW.length));
    } catch (e) {
      log.wordError = String(e);
    }

    // 1c. Mixed fonts: give 我去 another font, select the first six
    // characters, and see what the host reports and whether the pane warns.
    log.mixed = [];
    for (const other of ["Arial", "PingFang HK", "Songti SC", "Microsoft JhengHei"]) {
      const m = { other };
      try {
        await PowerPoint.run(async (ctx) => {
          box(ctx).getSubstring(0, 2).font.name = other;
          await ctx.sync();
        });
        await select("mixedAway", (tr) => tr.getSubstring(8, 1)); // so the next one is a change
        await select("mixed", (tr) => tr.getSubstring(0, 6));
        await wait(3000);
        await PowerPoint.run(async (ctx) => {
          const r = ctx.presentation.getSelectedTextRange();
          const part = box(ctx).getSubstring(0, 2);
          r.font.load("name");
          part.font.load("name");
          await ctx.sync();
          m.partReports = JSON.stringify(part.font.name);
          m.selectionReports = JSON.stringify(r.font.name);
        });
        m.warningShown = !$("mixed").hidden;
      } catch (e) {
        m.error = String(e);
      }
      log.mixed.push(m);
    }
    try {
      await PowerPoint.run(async (ctx) => {
        box(ctx).getSubstring(0, 2).font.name = family;
        await ctx.sync();
      });
      await select("setup3", (tr) => tr.getSubstring(0, RAW.length));
      await wait(3000);
      log.mixedClearedAfter = $("mixed").hidden;
    } catch (e) {
      log.mixedError = String(e);
    }

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
