#!/usr/bin/env node
/**
 * Renders the home card parts that CardService can't style (CYB4-2596 mock):
 *   assets/home/chip-intro-<lang>.png   Chip block: squirrel, 26px heading,
 *                                       light blue rounded box
 *   assets/home/btn-report-<lang>.png   full-width red "Report Email"
 *   assets/home/btn-how-<lang>.png      full-width outlined "How reporting works"
 *   assets/home/footer-<lang>.png       footer row: version left, language box
 *                                       right (footer-single-<lang>.png: no arrow,
 *                                       for accounts with one language)
 *   assets/home/footer.json             the version baked into the footers;
 *                                       npm test fails when package.json moves on
 * for every language in MESSAGES (i18n.gs), using the mock's sizes, colours
 * and Roboto. Re-run after changing any home.* text, then commit the PNGs:
 *
 *   npm run render:home            (needs Google Chrome; CHROME=/path overrides)
 *
 * Rendered 300 CSS px wide at 2x: the Gmail panel shows images ~290 px wide,
 * so text appears at about the mock's size.
 */
"use strict";
const fs = require("fs");
const path = require("path");
const vm = require("vm");
const os = require("os");
const { execFileSync } = require("child_process");

const ROOT = path.resolve(__dirname, "..");
const OUT = path.join(ROOT, "assets", "home");
const WIDTH = 300;
const SCALE = 2;
const CHROME = process.env.CHROME || "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome";

if (!fs.existsSync(CHROME)) {
  console.error(`Chrome not found at ${CHROME}. Set CHROME=/path/to/chrome.`);
  process.exit(1);
}

// i18n.gs only declares data and functions at top level, so it loads alone.
const sandbox = { console: { log() {}, warn() {}, error() {} } };
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync(path.join(ROOT, "i18n.gs"), "utf8") + "\n;this.__MESSAGES = MESSAGES;", sandbox);
const MESSAGES = sandbox.__MESSAGES;
const VERSION = "v " + JSON.parse(fs.readFileSync(path.join(ROOT, "package.json"), "utf8")).version;
// Same label as the add-on's dropdown (i18n.gs localeOptionLabel).
const optionLabel = (lang) => vm.runInContext(`localeOptionLabel(${JSON.stringify(lang)})`, sandbox);

const dataUri = (file, mime) => `data:${mime};base64,${fs.readFileSync(path.join(ROOT, file)).toString("base64")}`;
const FONT = dataUri("assets/fonts/roboto-latin.woff2", "font/woff2");
const CHIP = dataUri("assets/chip.png", "image/png");
const esc = (s) => String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;");
const RTL = new Set(["ar", "he", "fa", "ur"]);

function page(lang, body) {
  return `<!doctype html><html lang="${lang}"><head><meta charset="utf-8"><style>
@font-face { font-family: "Roboto"; src: url(${FONT}) format("woff2"); font-weight: 100 900; }
html, body { margin: 0; background: transparent; width: ${WIDTH}px; overflow: hidden; }
body { width: ${WIDTH}px; font-family: "Roboto", "Geeza Pro", "Noto Sans Arabic", Arial, sans-serif; color: #1B1F3B; -webkit-font-smoothing: antialiased; }
p, h2, h3 { margin: 0; }
#root { display: block; }
</style></head><body><div id="root" dir="${RTL.has(lang) ? "rtl" : "ltr"}">${body}</div></body></html>`;
}

// The mock's Chip block, tightened so the whole home card fits a normal
// Gmail panel without scrolling: smaller squirrel (48x51, mock 76x80) and
// heading (20px, mock 26px), 15px text with tighter line height and gaps.
function chipIntro(m) {
  return `<div style="box-sizing:border-box;padding:12px 14px;border-radius:10px;background:#F1F3FF;display:flex;flex-direction:column;gap:8px">
  <div style="display:flex;align-items:center;gap:12px">
    <img src="${CHIP}" alt="" style="width:48px;height:51px;flex:none;object-fit:contain">
    <h2 style="font-size:20px;line-height:26px;font-weight:700">${esc(m.chipGreeting)}</h2>
  </div>
  <p style="font-size:15px;line-height:21px">${esc(m.chipIntro)}</p>
  <div style="display:flex;flex-direction:column;gap:2px">
    <h3 style="font-size:15px;line-height:20px;font-weight:700">${esc(m.keyPartHeading)}</h3>
    <p style="font-size:15px;line-height:20px">${esc(m.keyPartBody)}</p>
  </div>
  <p style="font-size:15px;line-height:20px">${esc(m.spotOne)}</p>
</div>`;
}

// Buttons: the mock's look, a little shorter (44 and 40px, mock 52 and 48).
const reportButton = (m) =>
  `<div style="box-sizing:border-box;width:100%;height:44px;border-radius:6px;background:#C9402F;color:#FFFFFF;font-size:17px;font-weight:700;display:flex;align-items:center;justify-content:center;padding:0 12px;text-align:center">${esc(m.reportButton)}</div>`;
const howButton = (m) =>
  `<div style="box-sizing:border-box;width:100%;height:40px;border:2px solid #3B4FE4;border-radius:6px;background:#FFFFFF;color:#2F3FBF;font-size:16px;font-weight:700;display:flex;align-items:center;justify-content:center;padding:0 12px;text-align:center">${esc(m.howReportingWorks)}</div>`;

// Footer row from the mock: 13px muted version, 180x40 language box (mock 44).
const footer = (lang, arrow) => () =>
  `<div style="height:40px;display:flex;align-items:center;justify-content:space-between;gap:12px">
  <span style="font-size:13px;color:#4A4F66">${esc(VERSION)}</span>
  <div style="position:relative;box-sizing:border-box;width:180px;height:40px;border:1px solid #8A8FA3;border-radius:6px;background:#FFFFFF;display:flex;align-items:center;padding-inline:12px 34px;font-size:14px;color:#1B1F3B;white-space:nowrap;overflow:hidden">
    ${esc(optionLabel(lang))}
    ${arrow ? `<svg width="12" height="8" viewBox="0 0 12 8" style="position:absolute;inset-inline-end:12px;top:16px"><path d="M1 1.5 L6 6.5 L11 1.5" stroke="#1B1F3B" stroke-width="2" fill="none" stroke-linecap="round" stroke-linejoin="round"/></svg>` : ""}
  </div>
</div>`;

const tmp = fs.mkdtempSync(path.join(os.tmpdir(), "render-home-"));
const chrome = (args) => execFileSync(CHROME, ["--headless=new", "--disable-gpu", "--hide-scrollbars", "--no-first-run", "--allow-file-access-from-files", ...args], { encoding: "utf8", stdio: ["ignore", "pipe", "ignore"] });

function render(html, outFile) {
  const htmlFile = path.join(tmp, "page.html");
  // Pass 1: measure the content height once fonts have loaded.
  fs.writeFileSync(htmlFile, html.replace("</body>",
    `<script>document.fonts.ready.then(()=>{document.title="H"+Math.ceil(document.getElementById("root").getBoundingClientRect().height)})</script></body>`));
  const dom = chrome(["--virtual-time-budget=3000", "--dump-dom", "file://" + htmlFile]);
  const m = dom.match(/<title>H(\d+)<\/title>/);
  if (!m) throw new Error(`could not measure ${outFile}`);
  // Pass 2: screenshot exactly that box, transparent outside it.
  fs.writeFileSync(htmlFile, html);
  chrome([`--window-size=${WIDTH},${m[1]}`, `--force-device-scale-factor=${SCALE}`, "--default-background-color=00000000",
    "--virtual-time-budget=3000", `--screenshot=${outFile}`, "file://" + htmlFile]);
  return Number(m[1]);
}

fs.mkdirSync(OUT, { recursive: true });
for (const lang of Object.keys(MESSAGES)) {
  const m = MESSAGES[lang].home;
  for (const [name, part] of [["chip-intro", chipIntro], ["btn-report", reportButton], ["btn-how", howButton],
                              ["footer", footer(lang, true)], ["footer-single", footer(lang, false)]]) {
    const file = path.join(OUT, `${name}-${lang}.png`);
    const h = render(page(lang, part(m)), file);
    console.log(`${path.relative(ROOT, file)}  ${WIDTH}x${h} @${SCALE}x`);
  }
}
fs.writeFileSync(path.join(OUT, "footer.json"), JSON.stringify({ version: VERSION }, null, 2) + "\n");
console.log(`assets/home/footer.json  ${VERSION}`);
fs.rmSync(tmp, { recursive: true, force: true });
