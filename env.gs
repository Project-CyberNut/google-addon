/**
 * Environment configuration: the single place where prod and dev differ.
 *
 * The code on main and dev used to be identical apart from URLs and labels,
 * swapped by hand and drifting apart. Every environment-specific value now
 * lives here, keyed by environment, and the rest of the code reads it via
 * `env()`. One code base, one manifest, deployed to both projects.
 *
 * Which environment runs is decided by the Script Property `ADDON_ENV`
 * (Apps Script editor -> Project Settings -> Script Properties): "prod" or
 * "dev". Unset means prod. `UNIFICATION_ENV` is honoured as a legacy alias
 * for deployments that set it before this file existed.
 *
 * The one value the platform will not let us switch at runtime is the add-on
 * name in Appscript.json (`addOns.common.name`): the manifest is static. The
 * dev Apps Script project keeps its own manifest with "Cybernut Dev" there;
 * everything else in the manifest, including the whitelist, is the same for
 * both because the whitelist carries both environments' hosts.
 *
 * API Gateway ids per environment come from the branches that used to hold
 * them: prod from main, dev from the dev branch's code (its manifest
 * disagreed with its code for ap-southeast-1; the code's ids match the
 * portal's dev .env, so those are used).
 */

var ENVIRONMENTS = {
  prod: {
    /** Heading on every card; the manifest name for this env is "CyberNut Reporting Tool". */
    heading: "Cybernut Reporting Tool",
    /** Telemetry (callErrorReportingApi). Route is microsoftaddinactivitynew in both envs. */
    telemetryHost: "https://560ef3pt4j.execute-api.us-east-1.amazonaws.com",
    /** Domain -> AWS region lookup. */
    userRegionHost: "https://44dgkpf1cb.execute-api.us-east-1.amazonaws.com",
    /** Per-region API Gateway ids: verify (admindomainsgoogle, campaignversion), admin (getemail), service (eventdispatcher). */
    regions: {
      "us-east-1":      { verifyUrl: "44dgkpf1cb", adminUrl: "k3g591je54", serviceUrl: "560ef3pt4j" },
      "ap-southeast-1": { verifyUrl: "vsqdkxcc8d", adminUrl: "b4nzi83qm2", serviceUrl: "vahgicl5qh" },
      "eu-central-1":   { verifyUrl: "telmnzu55i", adminUrl: "dej7cfclm9", serviceUrl: "p3shdnpenc" },
    },
    /** Training portal host: v2 report redirect, onboarding link, and the body-link check. */
    trainingHost: "training.cybernut.com",
    /** Legacy report portal host. */
    portalHost: "www.cybernut-k12.com",
    /** Unification platform (language preference routes), per region. */
    unification: {
      "us-east-1":      "https://prod-us-east-1.cybernut.ai/api/v1",
      "ap-southeast-1": "https://prod-ap-southeast-1.cybernut.ai/api/v1",
      "eu-central-1":   "https://prod-eu-central-1.cybernut.ai/api/v1",
    },
  },

  dev: {
    heading: "Cybernut Reporting Tool (Dev)",
    telemetryHost: "https://rhqh5ihdvj.execute-api.us-east-1.amazonaws.com",
    userRegionHost: "https://rg0w8yelb6.execute-api.us-east-1.amazonaws.com",
    regions: {
      "us-east-1":      { verifyUrl: "u2o82lbd9f", adminUrl: "rg0w8yelb6", serviceUrl: "rhqh5ihdvj" },
      "ap-southeast-1": { verifyUrl: "9tp2t9h2o2", adminUrl: "o1gk4tisc4", serviceUrl: "cllxz8kqk7" },
      "eu-central-1":   { verifyUrl: "9v7i6h5197", adminUrl: "efmvxrr92j", serviceUrl: "7jww0knq3g" },
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

/** Memoised per execution. */
var currentEnvMemo = null;

/** "prod" or "dev", from Script Properties; unknown or unset -> prod. */
function currentEnvName() {
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
    prefixes.push(host(ids.adminUrl) + "/getemail");
    prefixes.push(host(ids.serviceUrl) + "/eventdispatcher");
  });
  Object.keys(cfg.unification).forEach(function (reg) {
    prefixes.push(cfg.unification[reg] + "/public/");
  });
  return prefixes;
}

/** Run from the Apps Script editor: logs the active environment and every URL it will fetch. */
function debugEnvironment() {
  var name = currentEnvName();
  console.log("environment", name, ENVIRONMENTS[name]);
  console.log("fetch prefixes", environmentFetchPrefixes(name));
}
