#!/usr/bin/env node
/**
 * Pre-push checks: run by `npm test` locally and by CI before any push.
 *   1. every *.gs parses as JavaScript
 *   2. Appscript.json is valid JSON with the fields the add-on needs
 *   3. every message catalogue has exactly the keys `en` has, and every
 *      t("...") key used in code exists
 *   4. no host is hard-coded outside env.gs
 *   5. for each environment, every URL prefix the code can fetch is covered
 *      by the manifest the build would produce, and each environment block
 *      has the same shape
 */
"use strict";
const fs = require("fs");
const path = require("path");
const vm = require("vm");

const ROOT = path.resolve(__dirname, "..");
const failures = [];
const fail = (msg) => failures.push(msg);
const read = (f) => fs.readFileSync(path.join(ROOT, f), "utf8");
const sources = fs.readdirSync(ROOT).filter((f) => f.endsWith(".gs")).sort();

// 1. syntax
for (const f of sources) {
  try { new vm.Script(read(f), { filename: f }); }
  catch (err) { fail(`${f}: syntax error: ${err.message}`); }
}

// 2. manifest
let manifest = null;
try {
  manifest = JSON.parse(read("Appscript.json"));
  for (const k of ["oauthScopes", "addOns", "urlFetchWhitelist"]) if (!manifest[k]) fail(`Appscript.json: missing ${k}`);
  if (manifest.addOns && manifest.addOns.common && manifest.addOns.common.homepageTrigger?.runFunction !== "HomePage") fail("Appscript.json: homepageTrigger must run HomePage");
} catch (err) { fail(`Appscript.json: ${err.message}`); }

// shared sandbox with the sources loaded (Apps Script globals stubbed)
const sandbox = {
  console: { log() {}, warn() {}, error() {} },
  PropertiesService: { getScriptProperties: () => ({ getProperty: () => null }), getUserProperties: () => ({ getProperty: () => null, setProperty() {} }) },
  Session: { getActiveUser: () => ({ getEmail: () => "" }), getActiveUserLocale: () => "en" },
  UrlFetchApp: { fetch() { throw new Error("no network in checks"); } },
  CardService: new Proxy({}, { get: () => () => new Proxy(function () {}, { get: (t, p) => (p === "build" ? () => ({}) : () => new Proxy(function () {}, { get: () => () => ({}) })) }) }),
  GmailApp: {}, Logger: { log() {} },
};
vm.createContext(sandbox);
let loaded = true;
for (const f of sources) {
  try { vm.runInContext(read(f), sandbox, { filename: f }); }
  catch (err) { loaded = false; fail(`${f}: failed to load: ${err.message}`); }
}

if (loaded) {
  // 3. catalogues
  const flat = (o, p = "") => Object.entries(o).flatMap(([k, v]) => (v && typeof v === "object" ? flat(v, p + k + ".") : [p + k]));
  const en = new Set(flat(sandbox.MESSAGES.en));
  for (const code of Object.keys(sandbox.MESSAGES)) {
    const keys = new Set(flat(sandbox.MESSAGES[code]));
    for (const k of en) if (!keys.has(k)) fail(`MESSAGES.${code}: missing key ${k}`);
    for (const k of keys) if (!en.has(k)) fail(`MESSAGES.${code}: key ${k} not in en`);
  }
  const used = new Set();
  for (const f of sources) for (const m of read(f).matchAll(/\bt\("([^"]+)"/g)) used.add(m[1]);
  for (const k of used) if (!en.has(k)) fail(`t("${k}") used but not in MESSAGES.en`);

  // 4. no hard-coded hosts outside env.gs (templates with ${...} ids are fine)
  const hostPattern = /https?:\/\/(?!\$\{)[a-z0-9-]+(\.[a-z0-9-]+)+/gi;
  for (const f of sources.filter((s) => s !== "env.gs")) {
    read(f).split("\n").forEach((line, i) => {
      const trimmed = line.trim();
      if (trimmed.startsWith("//") || trimmed.startsWith("*") || trimmed.startsWith("/*")) return;
      for (const m of line.matchAll(hostPattern)) {
        if (/googleapis\.com|w3\.org/.test(m[0])) continue;
        fail(`${f}:${i + 1}: hard-coded host ${m[0]} - move it to env.gs`);
      }
    });
  }

  // 5. environments
  const envs = Object.keys(sandbox.ENVIRONMENTS);
  const shape = (o) => JSON.stringify(Object.keys(o).sort());
  const regionShape = (cfg) => JSON.stringify(Object.keys(cfg.regions).sort());
  for (const name of envs) {
    const cfg = sandbox.ENVIRONMENTS[name];
    if (shape(cfg) !== shape(sandbox.ENVIRONMENTS.prod)) fail(`ENVIRONMENTS.${name}: keys differ from prod`);
    if (regionShape(cfg) !== regionShape(sandbox.ENVIRONMENTS.prod)) fail(`ENVIRONMENTS.${name}: regions differ from prod`);
    for (const [reg, ids] of Object.entries(cfg.regions)) for (const k of ["verifyUrl", "adminUrl", "serviceUrl"]) if (!/^[a-z0-9]{10}$/.test(ids[k] || "")) fail(`ENVIRONMENTS.${name}.regions.${reg}.${k}: not an API Gateway id`);
    // the manifest the build would produce for this env must cover every prefix
    const prefixes = sandbox.environmentFetchPrefixes(name);
    const whitelist = [...prefixes, ...(manifest?.urlFetchWhitelist || []).filter((u) => !/execute-api|cybernut\.ai/.test(u))];
    for (const p of prefixes) if (!whitelist.some((w) => p.startsWith(w))) fail(`${name}: ${p} would not be whitelisted`);
    for (const p of prefixes) if (!p.startsWith("https://")) fail(`${name}: ${p} is not https`);
  }
}

if (failures.length) {
  console.error(`\n${failures.length} check(s) failed:\n` + failures.map((f) => "  - " + f).join("\n"));
  process.exit(1);
}
console.log(`all checks passed: ${sources.length} sources, ${Object.keys(sandbox.MESSAGES || {}).length} catalogues, ${Object.keys(sandbox.ENVIRONMENTS || {}).length} environments`);
