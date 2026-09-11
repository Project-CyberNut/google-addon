#!/usr/bin/env node
/**
 * Build one environment and clasp-push it, failing early with a clear message
 * when the script id or the clasp login is missing.
 *
 *   node scripts/push.js dev|prod
 */
"use strict";
const fs = require("fs");
const os = require("os");
const path = require("path");
const { spawnSync } = require("child_process");

const ROOT = path.resolve(__dirname, "..");
const target = (process.argv[2] || "").toLowerCase();
if (!["dev", "prod"].includes(target)) {
  console.error("usage: node scripts/push.js dev|prod");
  process.exit(2);
}

const build = spawnSync(process.execPath, [path.join(__dirname, "build.js"), target], { stdio: "inherit" });
if (build.status !== 0) process.exit(build.status || 1);

const dist = path.join(ROOT, "dist", target);
const problems = [];

if (!fs.existsSync(path.join(dist, ".clasp.json"))) {
  problems.push(
    `No script id for ${target}. Paste it into clasp.targets.json (from the ${target}\n` +
      `    project's URL, script.google.com/.../projects/<id>/edit) or export SCRIPT_ID_${target.toUpperCase()}=<id>`
  );
}
if (!fs.existsSync(path.join(os.homedir(), ".clasprc.json"))) {
  problems.push("Not logged in to clasp. Run: clasp login");
}
if (problems.length) {
  console.error("\ncannot push:\n  - " + problems.join("\n  - "));
  process.exit(1);
}

console.log(`\npushing dist/${target} -> ${target} Apps Script project`);
const push = spawnSync("clasp", ["push", "--force"], { cwd: dist, stdio: "inherit" });
process.exit(push.status === null ? 1 : push.status);
