// Core of the add-in: shape with HarfBuzz, then rewrite the text into a
// character sequence that reaches the same glyphs through cmap alone
// (plain char > IVS > PUA), so apps that skip GSUB still render it.
// No DOM / Office dependencies — runs in node for test.mjs.

const IGNORABLE = /\p{Default_Ignorable_Code_Point}/u;
const isPUA = (cp) =>
  (cp >= 0xe000 && cp <= 0xf8ff) || (cp >= 0xf0000 && cp <= 0xffffd) || cp >= 0x100000;
const sameGids = (a, b) => a.length === b.length && a.every((g, i) => g === b[i]);

/** Non-default UVS mappings from a raw cmap table: Map<selector, Map<base, gid>>. */
export function parseUvs(cmap) {
  const uvs = new Map();
  if (!cmap) return uvs;
  const dv = new DataView(cmap.buffer, cmap.byteOffset, cmap.byteLength);
  const u24 = (o) => (dv.getUint8(o) << 16) | dv.getUint16(o + 1);
  for (let i = 0, n = dv.getUint16(2); i < n; i++) {
    const sub = dv.getUint32(4 + i * 8 + 4);
    if (dv.getUint16(sub) !== 14) continue;
    for (let r = 0, nr = dv.getUint32(sub + 6); r < nr; r++) {
      const rec = sub + 10 + r * 11;
      const nonDefault = dv.getUint32(rec + 7);
      if (!nonDefault) continue;
      const map = new Map();
      const base = sub + nonDefault;
      for (let m = 0, nm = dv.getUint32(base); m < nm; m++) {
        map.set(u24(base + 4 + m * 5), dv.getUint16(base + 4 + m * 5 + 3));
      }
      uvs.set(u24(rec), map);
    }
  }
  return uvs;
}

/**
 * Load a font (sfnt bytes) and return { convert, variants }.
 * `hb` is the harfbuzzjs module.
 */
export function loadFont(hb, bytes) {
  const face = new hb.Face(new hb.Blob(bytes));
  const font = new hb.Font(face);
  const uvs = parseUvs(face.referenceTable("cmap"));
  const nominal = (cp) => font.nominalGlyph(cp) ?? 0;

  // gid → best character sequence. Rank: plain char (0) > IVS (1) > PUA (2).
  // IVS degrades to the bare Han char without the font; PUA becomes tofu.
  // Ties prefer NFKC-stable chars (行 over the Kangxi radical ⾏).
  const reverse = new Map();
  const offer = (gid, str, rank) => {
    const key = [rank, str.normalize("NFKC") === str ? 0 : 1, str.codePointAt(0)];
    const cur = reverse.get(gid);
    if (!cur || key < cur.key) reverse.set(gid, { str, key });
  };
  for (const cp of face.collectUnicodes()) {
    const gid = nominal(cp);
    if (gid) offer(gid, String.fromCodePoint(cp), isPUA(cp) ? 2 : 0);
  }
  for (const [vs, map] of uvs) {
    for (const [cp, gid] of map) offer(gid, String.fromCodePoint(cp, vs), isPUA(cp) ? 2 : 1);
  }

  // The generator substitutes a consumed trigger digit (一刀兩斷２) with a
  // cmap-less, zero-advance, empty glyph. It shows nothing, so "no
  // character" is its cmap equivalent: drop it from every glyph sequence.
  const isSwallowed = (gid) => {
    if (reverse.has(gid) || font.glyphHAdvance(gid)) return false;
    const e = font.glyphExtents(gid);
    return !e || (!e.width && !e.height);
  };

  const shape = (text) => {
    const buf = new hb.Buffer();
    buf.addText(text);
    buf.guessSegmentProperties();
    // Drop ZWJ / unused selectors instead of turning them into space glyphs.
    buf.setFlags(hb.BufferFlag.REMOVE_DEFAULT_IGNORABLES);
    hb.shape(font, buf);
    return buf
      .getGlyphInfos()
      .map((g) => ({ gid: g.codepoint, cluster: g.cluster }))
      .filter((g) => !isSwallowed(g.gid));
  };

  /** Glyphs an app WITHOUT GSUB shows: cmap + variation sequences only. */
  const cmapOnly = (text) => {
    const cps = Array.from(text, (c) => c.codePointAt(0));
    const out = [];
    for (let i = 0; i < cps.length; i++) {
      const variant = uvs.get(cps[i + 1])?.get(cps[i]);
      if (variant) out.push(variant), i++;
      else if (!IGNORABLE.test(String.fromCodePoint(cps[i]))) out.push(nominal(cps[i]));
    }
    return out;
  };

  const convertLine = (line, offset, changes, failed) => {
    // Group glyphs by cluster (sorted, so RTL runs work too); each group
    // owns the source text from its cluster start to the next one.
    const groups = new Map();
    for (const g of shape(line)) {
      if (!groups.has(g.cluster)) groups.set(g.cluster, []);
      groups.get(g.cluster).push(g.gid);
    }
    const starts = [...groups.keys()].sort((a, b) => a - b);
    let out = line.slice(0, starts[0] ?? line.length);
    starts.forEach((start, i) => {
      const end = starts[i + 1] ?? line.length;
      const src = line.slice(start, end);
      const gids = groups.get(start);
      if (sameGids(gids, cmapOnly(src))) return (out += src); // already cmap-only
      if (!gids.every((g) => reverse.has(g))) {
        failed.push({ start: offset + start, end: offset + end });
        return (out += src);
      }
      const replacement = gids.map((g) => reverse.get(g).str).join("");
      changes.push({ start: offset + start, end: offset + end, replacement });
      out += replacement;
    });
    return out;
  };

  /**
   * @returns {{
   *   text: string,                 // converted text
   *   changes: {start,end,replacement}[], // UTF-16 offsets into the input
   *   failed: {start,end}[],        // glyphs with no cmap path: left as-is
   *   cmapOk: boolean,              // GSUB-less apps show the shaped glyphs
   *   gsubOk: boolean,              // GSUB-aware apps still show the same glyphs
   * }}
   */
  const convert = (text) => {
    const changes = [];
    const failed = [];
    let out = "";
    let offset = 0;
    let cmapOk = true;
    let gsubOk = true;
    // Shape each line on its own; PowerPoint uses \r (paragraph) and \v (line).
    for (const part of text.split(/([\r\n\v\u2028\u2029]+)/)) {
      if (part && !/^[\r\n\v\u2028\u2029]/.test(part)) {
        const nFailed = failed.length;
        const line = convertLine(part, offset, changes, failed);
        const want = shape(part).map((g) => g.gid);
        // Self-check: re-shape the output; reject silently-different results.
        if (!sameGids(shape(line).map((g) => g.gid), want)) gsubOk = false;
        if (failed.length === nFailed && !sameGids(cmapOnly(line), want)) cmapOk = false;
        out += line;
      } else out += part;
      offset += part.length;
    }
    return { text: out, changes, failed, cmapOk, gsubOk };
  };

  /** Selectable variants of one character: [{ text, gid }], default first. */
  const variants = (char) => {
    const cp = char.codePointAt(0);
    const list = [{ text: String.fromCodePoint(cp), gid: nominal(cp) }];
    for (const [vs, map] of [...uvs].sort((a, b) => a[0] - b[0])) {
      if (map.has(cp)) list.push({ text: String.fromCodePoint(cp, vs), gid: map.get(cp) });
    }
    return list;
  };

  return { convert, variants, shape, cmapOnly };
}

/** Undo: delete IVS selectors and BMP PUA mark carriers (lossy — see README). */
export const stripChanges = (text) =>
  [...text.matchAll(/[\u{E000}-\u{F8FF}\u{E0100}-\u{E01EF}]+/gu)].map((m) => ({
    start: m.index,
    end: m.index + m[0].length,
    replacement: "",
  }));

/** Apply non-overlapping, ascending { start, end, replacement } edits. */
export const applyChanges = (text, changes) =>
  changes.reduceRight((t, c) => t.slice(0, c.start) + c.replacement + t.slice(c.end), text);
