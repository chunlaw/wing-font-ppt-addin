# wing-font-ppt-addin

An Office add-in for [Wing Font](https://github.com/chunlaw/wing-font-generator) annotated fonts. Built for PowerPoint first; it also loads in Word.

PowerPoint does not apply GSUB to Chinese text, so `行２`, `行０ｚａａ１` and word-based readings (`銀行`) all stop working there. The add-in shapes the text with HarfBuzz in its task pane, then rewrites it into characters that render through **cmap alone** (IVS / PUA). The design notes are in [TODO.md](TODO.md) (Chinese).

## Usage

1. Select some text (or a whole text box) and press **Convert selection**.
2. Select a few characters, or just put the cursor in the text, and the pane lists the readings of each character, one row per character; click one to switch. With only a cursor, PowerPoint on Windows reports the word around it and PowerPoint on Mac the character after it (the pane adds the one before).
3. **Restore** deletes the IVS selectors and BMP PUA (U+E000–F8FF) mark carriers in the selection. It does not bring back the `行２` syntax.

If the selection mixes fonts, a warning appears under the Convert button: everything is converted with the one font shown in the Font card.

The pane follows the Office display language: Traditional Chinese for any Chinese locale, English otherwise.

By default the font is the one the selected text uses, downloaded from `https://wing-font.chunlaw.io/fonts/<name>.woff2` and cached. You can instead type another official font name or a URL in the font field, or pick a local font file (it stays on your machine and is never uploaded).

The add-in cannot read the fonts installed on your computer: it converts with its own copy, while PowerPoint renders with the installed one. The two must be the same build, because the numbering of readings can differ between builds. The pane therefore shows the loaded font's version string (for example `Version 2.004; modified by Wing Font 2026-10-05`), which you can compare with the installed font in Font Book or Windows font properties. If the font has no build date, or there is no copy to download, the pane asks you to load the installed font file instead. A downloaded font is checked against the server each time and fetched again when it has been republished. A font file you picked is remembered under its file name; if the official font of the same family has been republished since that file was built, the pane says so and asks you to pick the file again.

> **Font requirement**: the font must be rebuilt with a generator version that adds the bare-glyph IVS and the cmap route for DIY marks, and that fixes `-opt` dropping the IVS table. Older fonts have no cmap route, so the add-in reports those spots and leaves the original text in place.

## Install (sideload)

The pane is a static site. On every push to `main`, [`.github/workflows/pages.yml`](.github/workflows/pages.yml) runs the tests, builds, and deploys to GitHub Pages (set Settings → Pages → Source to **GitHub Actions**). `manifest.xml` assumes the site is at `https://wing-font-ppt.chunlaw.io/`; search and replace that URL if you host it elsewhere.

- **Mac**: copy `manifest.xml` to `~/Library/Containers/com.microsoft.Powerpoint/Data/Documents/wef/` (`com.microsoft.Word` for Word) and restart PowerPoint. **Wing Font** appears on the Home tab.
- **Windows**: put `manifest.xml` in a folder and share that folder (Properties → Sharing → Share). Under File → Options → Trust Center → Trust Center Settings → Trusted Add-in Catalogs, add the folder by its **network path** (`\\COMPUTER\Share`; a local path such as `C:\Users\…` is rejected with a message about `https://`), tick **Show in Menu**, and restart PowerPoint. Then Insert → My Add-ins → Shared Folder.
- **Organisation (Microsoft 365)**: admin center → Settings → Integrated apps → Upload custom apps → upload `manifest.xml`.

## Development

```bash
npm install
```

```bash
npm test
```

```bash
python3 -m http.server 8765
```

`npm install` copies the browser files of harfbuzzjs and wawoff2 from `node_modules` into `vendor/` (not in git). `npm run build` additionally assembles the deployable `dist/`.

Then open <http://localhost:8765/index.html?font=test/WingSmall.woff2>; add `&lang=en` or `&lang=zh` to force a language. Outside Office the pane shows a sandbox text box with `ccmp` turned off, which imitates PowerPoint's lack of GSUB.

| File | What it is |
|---|---|
| `convert.js` | The core: HarfBuzz shaping → reverse cmap lookup → self-check. No DOM or Office dependency |
| `i18n.js` | UI strings: Traditional Chinese and English |
| `app.js` | The pane: font loading and caching, reading and writing the selection through Office.js, the reading picker |
| `build.mjs` | Copies `vendor/`, assembles `dist/` |
| `vendor/` (generated) | [harfbuzzjs](https://github.com/harfbuzz/harfbuzzjs) (shaping) and [wawoff2](https://github.com/fontello/wawoff2) (woff2 decompression); versions are pinned in `package.json` |
| `test/office/` | End-to-end self-test that runs inside real PowerPoint: open the pane with `?selftest=/report` (served by `serve.py`) and it adds a text box to slide 1, drives Convert / the picker / Restore, and posts what it read back |
| `test/WingSmall.*` | Test font (行: VS17 = haang4, VS18 = hang4, VS19 = hong4, the default; DIY marks: zaa1 = U+E000, hong2 = U+E000 + VS17) |
| `test/m0.pptx` | Rendering test deck, generated by the converter; needs the `WingM0v2` font installed. `m0-win.png` and `m0-mac.png` are the results from PowerPoint on Windows and Mac |

## Known limitations

- The ribbon has a single button that opens the pane; Convert and Restore live inside the pane.
- Word, and PowerPoint without PowerPointApi 1.5, go through the common API: the whole selection is replaced at once, so mixed formatting inside it is flattened and the font cannot be detected automatically.
- Word-level composites for non-CJK scripts (Arabic and others) have no cmap route yet.
- Converted text is meant to stay in PowerPoint. Copied out into an app that runs GSUB, two things can go wrong (both accepted, not planned to fix):
  - A small share of words (about 0.1–0.8% of mapping entries, e.g. `波士頓`) get a wrong reading, because pinning one character stops the longer word rule from matching and a shorter one takes over. The add-in warns when it converts such a word.
  - Custom (DIY) annotations do not show in Word for Windows; readings and bare characters do. The add-in does not detect this.
