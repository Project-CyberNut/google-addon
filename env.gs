/**
 * Environment configuration: the single place where prod and dev differ.
 *
 * The code on main and dev used to be identical apart from URLs and labels,
 * swapped by hand and drifting apart. Every environment-specific value now
 * lives here, keyed by environment, and the rest of the code reads it via
 * `env()`. One code base, one manifest, deployed to both projects.
 *
 * Which environment runs: a project built by scripts/build.js carries
 * BUILD_ENV in build.gs, and that decides. A project pasted in by hand uses
 * the Script Property `ADDON_ENV` (Apps Script editor -> Project Settings ->
 * Script Properties): "prod" or "dev"; unset means prod. `UNIFICATION_ENV`
 * is honoured as a legacy alias.
 *
 * The add-on name in the manifest (`addOns.common.name`) cannot be read from
 * here - the manifest is static - so scripts/build.js writes it per target
 * along with a urlFetchWhitelist generated from environmentFetchPrefixes().
 *
 * API Gateway ids per environment come from the branches that used to hold
 * them: prod from main, dev from the dev branch's code (its manifest
 * disagreed with its code for ap-southeast-1; the code's ids match the
 * portal's dev .env, so those are used).
 */

var ENVIRONMENTS = {
  prod: {
    /** Telemetry (callErrorReportingApi). Route is microsoftaddinactivitynew in both envs. */
    telemetryHost: "https://560ef3pt4j.execute-api.us-east-1.amazonaws.com",
    /** Domain -> AWS region lookup. */
    userRegionHost: "https://44dgkpf1cb.execute-api.us-east-1.amazonaws.com",
    /** Per-region API Gateway ids: verify (admindomainsgoogle, campaignversion), service (eventdispatcher). */
    regions: {
      "us-east-1":      { verifyUrl: "44dgkpf1cb", serviceUrl: "560ef3pt4j" },
      "ap-southeast-1": { verifyUrl: "vsqdkxcc8d", serviceUrl: "vahgicl5qh" },
      "eu-central-1":   { verifyUrl: "telmnzu55i", serviceUrl: "p3shdnpenc" },
    },
    /** Training portal host: v2 report redirect, onboarding link, and the body-link check. */
    trainingHost: "training.cybernut.com",
    /** Legacy report portal host. */
    portalHost: "www.cybernut-k12.com",
    /** Unification platform (language preferences, plugin settings), per region. */
    unification: {
      "us-east-1":      "https://prod-us-east-1.cybernut.ai/api/v1",
      "ap-southeast-1": "https://prod-ap-southeast-1.cybernut.ai/api/v1",
      "eu-central-1":   "https://prod-eu-central-1.cybernut.ai/api/v1",
    },
  },

  dev: {
    telemetryHost: "https://rhqh5ihdvj.execute-api.us-east-1.amazonaws.com",
    userRegionHost: "https://rg0w8yelb6.execute-api.us-east-1.amazonaws.com",
    regions: {
      "us-east-1":      { verifyUrl: "u2o82lbd9f", serviceUrl: "rhqh5ihdvj" },
      "ap-southeast-1": { verifyUrl: "9tp2t9h2o2", serviceUrl: "cllxz8kqk7" },
      "eu-central-1":   { verifyUrl: "9v7i6h5197", serviceUrl: "7jww0knq3g" },
    },
    trainingHost: "dev-training.cybernut.com",
    portalHost: "userportaldev.cybernut-k12.com",
    unification: {
      "us-east-1":      "https://dev-us-east-1.cybernut.ai/api/v1",
      "ap-southeast-1": "https://dev-ap-southeast-1.cybernut.ai/api/v1",
      "eu-central-1":   "https://dev-eu-central-1.cybernut.ai/api/v1",
    },
  },
};

var DEFAULT_ENV = "prod";

/**
 * Home panel assets (CYB4-2596), the same in prod and dev. Each must be empty
 * or https (scripts/check.js). Images are loaded by Gmail, not UrlFetchApp,
 * so they need no whitelist entry.
 *
 *   OVERVIEW_VIDEO_URL           "See an overview of CyberNut" video. Opens
 *                                in a new tab (Gmail can't play it in the
 *                                panel). Empty hides the section.
 *   OVERVIEW_VIDEO_THUMBNAIL_URL Preview above the "Watch" button: a short
 *                                looping GIF, or a still image. Empty shows
 *                                only the button.
 *   CHIP_IMAGE_URL               Chip the squirrel beside "Hi, I'm Chip.".
 *                                Empty shows the greeting without the image.
 */
// TEMP test video (CyberNut "Introducing Advanced Threat Search") and a 30s GIF
// cut from it, served from this public repo (assets/). Before prod: confirm
// the final video and move the GIF to CyberNut's own storage.
var OVERVIEW_VIDEO_URL = "https://www.youtube.com/watch?v=dHMQR2qz5Ow";
var OVERVIEW_VIDEO_THUMBNAIL_URL = "https://raw.githubusercontent.com/Project-CyberNut/google-addon/dev/assets/overview-preview.gif";
var CHIP_IMAGE_URL = "";

/** Memoised per execution. */
var currentEnvMemo = null;

/**
 * "prod" or "dev". A built project (scripts/build.js) carries BUILD_ENV in
 * build.gs and that wins, so a deployment can never run against the wrong
 * backend because of a property. A project pasted in by hand has no
 * build.gs and falls back to the Script Property; unknown or unset -> prod.
 */
function currentEnvName() {
  if (typeof BUILD_ENV === "string" && ENVIRONMENTS[BUILD_ENV]) return BUILD_ENV;
  var name = null;
  try {
    var props = PropertiesService.getScriptProperties();
    name = props.getProperty("ADDON_ENV") || props.getProperty("UNIFICATION_ENV");
  } catch (err) {
    console.log("currentEnvName|script properties unreadable|", err.message);
  }
  name = (name || DEFAULT_ENV).trim().toLowerCase();
  if (!ENVIRONMENTS[name]) {
    console.warn("currentEnvName|unknown ADDON_ENV, using prod|", { value: name });
    return DEFAULT_ENV;
  }
  return name;
}

/** The active environment's configuration. */
function env() {
  if (!currentEnvMemo) {
    var name = currentEnvName();
    currentEnvMemo = ENVIRONMENTS[name];
    console.log("env|", { name: name });
  }
  return currentEnvMemo;
}

/**
 * Every URL prefix the code can fetch, for one environment: what the
 * manifest's urlFetchWhitelist has to contain. Used by debugEnvironment() and
 * by the whitelist check in the repo, so code and manifest cannot drift.
 */
function environmentFetchPrefixes(name) {
  var cfg = ENVIRONMENTS[name];
  var prefixes = [
    cfg.userRegionHost + "/userregion",
    cfg.telemetryHost + "/microsoftaddinactivitynew",
  ];
  Object.keys(cfg.regions).forEach(function (reg) {
    var ids = cfg.regions[reg];
    var host = function (id) { return "https://" + id + ".execute-api." + reg + ".amazonaws.com"; };
    prefixes.push(host(ids.verifyUrl) + "/admindomainsgoogle");
    prefixes.push(host(ids.verifyUrl) + "/campaignversion");
    prefixes.push(host(ids.serviceUrl) + "/eventdispatcher");
  });
  Object.keys(cfg.unification).forEach(function (reg) {
    prefixes.push(cfg.unification[reg] + "/public/");
    prefixes.push(cfg.unification[reg] + "/accounts/settings/plugin");
  });
  return prefixes;
}

/** Run from the Apps Script editor: logs the active environment and every URL it will fetch. */
function debugEnvironment() {
  var name = currentEnvName();
  console.log("environment", name, {
    decidedBy: typeof BUILD_ENV === "string" ? "build.gs (BUILD_ENV)" : "script property / default",
    build: typeof BUILD_VERSION === "string" ? BUILD_VERSION + " " + BUILD_COMMIT + " " + BUILD_TIME : "not a built project",
  });
  console.log("config", ENVIRONMENTS[name]);
  console.log("fetch prefixes", environmentFetchPrefixes(name));
}
