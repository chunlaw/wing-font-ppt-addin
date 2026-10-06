// UI strings. Traditional Chinese for any zh-* locale, English otherwise.
// Values are strings or functions of the message's arguments.

const STRINGS = {
  zh: {
    convert: "轉換選取文字",
    restore: "還原",
    restoreTip: "刪走選取範圍內嘅讀音同自訂注音",
    pickHeading: "揀讀音",
    pickEmpty: "選取一個字，呢度會列出佢所有讀音。",
    fontHeading: "字型",
    fontEmpty: "未載入字型：轉換時會跟選取文字嘅字型自動載入。",
    changeFont: "更換字型",
    fontField: "官方字型名或 URL（留空＝跟選取文字）",
    fontPlaceholder: "例如 NotoSansHK-NotoCond-lshk",
    pickFile: "揀本機字型檔…",
    sandboxLabel: "測試文字",
    ready: "準備好。",
    loading: (name) => `載入字型 ${name}…`,
    loaded: (name) => `字型 ${name} 已載入。`,
    woff2Failed: "woff2 解壓失敗。",
    repickLocal: (file) => `請重新揀本機字型檔「${file}」。`,
    notFound: (name, status) => `搵唔到字型「${name}」（${status}）。`,
    notLoaded: (name) => `未載入字型「${name}」。`,
    loadLocal: "請撳下面「揀本機字型檔」，載入你電腦裝咗嗰個字型檔。",
    infoFont: (name) => `字型：${name}`,
    infoLocal: (file) => `${file}（本機檔案）`,
    infoVersion: (v) => `版本：${v || "（冇版本資料）"}`,
    checkVersion: "請確認同你電腦裝咗嗰隻版本一樣；唔同就喺「更換字型」揀本機字型檔。",
    noBuildDate: "呢隻字型冇 build 日期，確認唔到同你裝嗰隻係咪同一個版本。",
    versionChanged: (name) => `注意：字型「${name}」同呢份文件上次轉換時用嘅版本唔同，舊文字嘅讀音可能有變。`,
    noFont: "唔知用邊個字型：請喺「更換字型」填字型名，或者揀本機字型檔。",
    noSelection: "請先選取文字或者文字方塊。",
    selfCheckFailed: "自我驗證失敗：轉換結果同字型排版唔一致，冇改動文字。",
    converted: (n) => `已轉換 ${n} 處。`,
    failed: (n, list) => `有 ${n} 處冇 cmap 路徑，保留原文：${list}`,
    gsubWarning: "注意：轉換後喺會做 GSUB 嘅程式（Word、網頁）讀音可能唔同。",
    restored: (n) => `已還原 ${n} 處。`,
    quote: (s) => `「${s}」`,
  },
  en: {
    convert: "Convert selection",
    restore: "Restore",
    restoreTip: "Remove readings and custom annotations from the selection",
    pickHeading: "Pick a reading",
    pickEmpty: "Select a single character to list its readings here.",
    fontHeading: "Font",
    fontEmpty: "No font loaded yet. Converting loads the font of the selected text.",
    changeFont: "Change font",
    fontField: "Official font name or URL (blank = follow the selection)",
    fontPlaceholder: "e.g. NotoSansHK-NotoCond-lshk",
    pickFile: "Choose a local font file…",
    sandboxLabel: "Test text",
    ready: "Ready.",
    loading: (name) => `Loading font ${name}…`,
    loaded: (name) => `Font ${name} loaded.`,
    woff2Failed: "Could not decompress the woff2 file.",
    repickLocal: (file) => `Please choose the local font file “${file}” again.`,
    notFound: (name, status) => `Font “${name}” was not found (${status}).`,
    notLoaded: (name) => `Font “${name}” is not loaded.`,
    loadLocal: "Use “Choose a local font file” below to load the font file installed on this computer.",
    infoFont: (name) => `Font: ${name}`,
    infoLocal: (file) => `${file} (local file)`,
    infoVersion: (v) => `Version: ${v || "(none)"}`,
    checkVersion: "Check that this matches the font installed on this computer. If it differs, choose the local font file under “Change font”.",
    noBuildDate: "This font has no build date, so it cannot be matched against the installed one.",
    versionChanged: (name) => `Note: font “${name}” is a different build from the one last used to convert this document. Readings in already converted text may have changed.`,
    noFont: "No font to use: enter a font name or choose a local font file under “Change font”.",
    noSelection: "Select some text or a text box first.",
    selfCheckFailed: "Self-check failed: the result does not match the font's own shaping. Nothing was changed.",
    converted: (n) => `Converted ${n} ${n === 1 ? "place" : "places"}.`,
    failed: (n, list) => `${n} ${n === 1 ? "place has" : "places have"} no cmap route and were left as typed: ${list}`,
    gsubWarning: "Note: apps that apply GSUB (Word, browsers) may show different readings for the converted text.",
    restored: (n) => `Restored ${n} ${n === 1 ? "place" : "places"}.`,
    quote: (s) => `“${s}” `,
  },
};

let lang = "en";

/** Pick the language from a BCP 47 tag such as Office's displayLanguage. */
export function setLanguage(tag) {
  lang = /^zh\b/i.test(tag || "") ? "zh" : "en";
  document.documentElement.lang = lang === "zh" ? "zh-Hant" : "en";
  for (const el of document.querySelectorAll("[data-i18n]")) el.textContent = t(el.dataset.i18n);
  for (const el of document.querySelectorAll("[data-i18n-attr]")) {
    for (const pair of el.dataset.i18nAttr.split(",")) {
      const [attr, key] = pair.split(":");
      el.setAttribute(attr, t(key));
    }
  }
}

export function t(key, ...args) {
  const s = STRINGS[lang][key];
  return typeof s === "function" ? s(...args) : s;
}
