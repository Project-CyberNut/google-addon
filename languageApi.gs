/**
 * Language preference APIs on the CyberNut unification platform.
 *
 * Mirrors user-portal-micro-learning-v2/src/lib/languageApi.ts so the add-on
 * and the training portal read and write the same stored preference. All
 * three routes are public (no user JWT): they resolve the tenant from the
 * account `domain` (derived from the user's email) and are authenticated by
 * a shared service key sent as `x-service-key`.
 *
 *   GET   /public/accounts/supported-languages?domain=acme.org
 *         -> { success, data: { supportedLanguages: ["en", "es"] } }
 *   GET   /public/users/preferred-language?domain=acme.org&email=jane@acme.org
 *         -> { success, data: { preferredLanguage: "es" | null } }
 *   PATCH /public/users/preferred-language?domain=acme.org&email=jane@acme.org&lang=es
 *         -> 200, or 400 when `lang` is not in the account's supported list
 *
 * Configuration lives in Script Properties (Apps Script editor -> Project
 * Settings -> Script Properties), never in source:
 *   UNIFICATION_SERVICE_KEY  required - the shared service key
 *   UNIFICATION_ENV          optional - "prod" (default) or "dev"
 */

var UNIFICATION_BASES = {
  prod: {
    "us-east-1": "https://prod-us-east-1.cybernut.ai/api/v1",
    "ap-southeast-1": "https://prod-ap-southeast-1.cybernut.ai/api/v1",
    "eu-central-1": "https://prod-eu-central-1.cybernut.ai/api/v1",
  },
  dev: {
    "us-east-1": "https://dev-us-east-1.cybernut.ai/api/v1",
    "ap-southeast-1": "https://dev-ap-southeast-1.cybernut.ai/api/v1",
    "eu-central-1": "https://dev-eu-central-1.cybernut.ai/api/v1",
  },
};

var SUPPORTED_LANGUAGES_PATH = "/public/accounts/supported-languages";
var PREFERRED_LANGUAGE_PATH = "/public/users/preferred-language";

/** Base URL of the unification platform for an AWS region. */
function unificationBase(region) {
  var env = "prod";
  try {
    env = PropertiesService.getScriptProperties().getProperty("UNIFICATION_ENV") || "prod";
  } catch (err) {
    console.log("unificationBase|script properties unreadable|", err.message);
  }
  var bases = UNIFICATION_BASES[env] || UNIFICATION_BASES.prod;
  var base = bases[region] || bases["us-east-1"];
  console.log("unificationBase|", { env: env, region: region, base: base });
  return base;
}

/**
 * Headers for every language route. The key is read from Script Properties so
 * it never lands in the repo. Missing key -> the request is sent without it
 * and the service answers 401, which the callers treat as "unavailable".
 */
function serviceHeaders(callName) {
  var key = null;
  try {
    key = PropertiesService.getScriptProperties().getProperty("UNIFICATION_SERVICE_KEY");
  } catch (err) {
    console.log("serviceHeaders|script properties unreadable|", err.message);
  }
  if (!key) {
    console.warn(
      "serviceHeaders|UNIFICATION_SERVICE_KEY is not set - " + callName +
      " will be rejected. Add it under Project Settings -> Script Properties."
    );
    return { "Content-Type": "application/json" };
  }
  return { "Content-Type": "application/json", "x-service-key": key };
}

/** Account domain for the tenant lookup, e.g. `jane@acme.org` -> `acme.org`. */
function accountDomainFromEmail(email) {
  if (!email) return null;
  var at = email.lastIndexOf("@");
  if (at < 0 || at >= email.length - 1) return null;
  var domain = email.slice(at + 1).trim().toLowerCase();
  return domain || null;
}

/**
 * Query string identifying the user. `email` is sent alone when we have it;
 * `memberId` (campaign_members.id) is the fallback for links with no email.
 */
function identityQuery(identity) {
  var parts = ["domain=" + encodeURIComponent(identity.domain)];
  if (identity.email) {
    parts.push("email=" + encodeURIComponent(identity.email));
  } else if (identity.memberId) {
    parts.push("memberId=" + encodeURIComponent(identity.memberId));
  } else {
    return null;
  }
  return parts.join("&");
}

function parseJsonSafe(text) {
  try {
    return JSON.parse(text);
  } catch (err) {
    return null;
  }
}

/**
 * Language codes the account's admins have switched on.
 * Returns `null` when the list could not be read (caller offers every shipped
 * language); an empty array is a real answer (account configured nothing).
 */
function getSupportedLanguages(domain, region) {
  var url = unificationBase(region) + SUPPORTED_LANGUAGES_PATH +
    "?domain=" + encodeURIComponent(domain);
  console.log("getSupportedLanguages|called|", { url: url });
  try {
    var res = UrlFetchApp.fetch(url, {
      method: "get",
      headers: serviceHeaders("supported-languages"),
      muteHttpExceptions: true,
    });
    var code = res.getResponseCode();
    if (code !== 200) {
      console.warn("getSupportedLanguages|http failure|", { code: code, body: res.getContentText() });
      return null;
    }
    var body = parseJsonSafe(res.getContentText());
    var languages = body && body.data && body.data.supportedLanguages;
    if (!Array.isArray(languages)) {
      console.warn("getSupportedLanguages|no supportedLanguages array in response");
      return null;
    }
    var codes = languages.filter(function (c) { return typeof c === "string"; });
    console.log("getSupportedLanguages|result|", { codes: codes });
    return codes;
  } catch (err) {
    console.error("getSupportedLanguages|failed|", err.message);
    callErrorReportingApi("getSupportedLanguages failed: " + err.message, " ");
    return null;
  }
}

/**
 * The language this user chose before, or `null` when none is stored (the
 * normal first-visit answer) or the route is unavailable.
 */
function getPreferredLanguage(identity) {
  var query = identityQuery(identity);
  if (!query) {
    console.log("getPreferredLanguage|skipped|no email or memberId");
    return null;
  }
  var url = unificationBase(identity.region) + PREFERRED_LANGUAGE_PATH + "?" + query;
  console.log("getPreferredLanguage|called|", { url: url });
  try {
    var res = UrlFetchApp.fetch(url, {
      method: "get",
      headers: serviceHeaders("preferred-language"),
      muteHttpExceptions: true,
    });
    var code = res.getResponseCode();
    if (code !== 200) {
      console.warn("getPreferredLanguage|http failure|", { code: code, body: res.getContentText() });
      return null;
    }
    var body = parseJsonSafe(res.getContentText());
    var language = body && body.data && body.data.preferredLanguage;
    if (typeof language !== "string" || !language) {
      console.log("getPreferredLanguage|result|none stored");
      return null;
    }
    console.log("getPreferredLanguage|result|", { language: language });
    return language;
  } catch (err) {
    console.error("getPreferredLanguage|failed|", err.message);
    callErrorReportingApi("getPreferredLanguage failed: " + err.message, " ");
    return null;
  }
}

/**
 * Persists the user's choice on the account. `false` when it did not land
 * (network, 401 for a missing key, or 400 for a language the account does
 * not allow); the caller still switches the UI, only the memory is lost.
 */
function setPreferredLanguage(identity, lang) {
  var query = identityQuery(identity);
  if (!query) {
    console.log("setPreferredLanguage|skipped|no email or memberId");
    return false;
  }
  var url = unificationBase(identity.region) + PREFERRED_LANGUAGE_PATH + "?" + query +
    "&lang=" + encodeURIComponent(lang);
  console.log("setPreferredLanguage|called|", { url: url });
  try {
    var res = UrlFetchApp.fetch(url, {
      method: "patch",
      headers: serviceHeaders("set-preferred-language"),
      muteHttpExceptions: true,
    });
    var code = res.getResponseCode();
    if (code !== 200) {
      console.warn("setPreferredLanguage|http failure|", { code: code, body: res.getContentText() });
      callErrorReportingApi(
        "setPreferredLanguage returned " + code + " for lang " + lang + ": " + res.getContentText(),
        " "
      );
      return false;
    }
    console.log("setPreferredLanguage|success|", { lang: lang });
    return true;
  } catch (err) {
    console.error("setPreferredLanguage|failed|", err.message);
    callErrorReportingApi("setPreferredLanguage failed: " + err.message, " ");
    return false;
  }
}
