/**
 * Language preference APIs on the CyberNut unification platform.
 *
 * Mirrors the training portal so the add-on and the portal read and write
 * the same stored preference. Both routes are public: no user JWT and no
 * service key. They resolve the tenant from the account `domain` (derived
 * from the user's email).
 *
 *   GET   /public/users/preferred-language?domain=acme.org&email=jane@acme.org
 *         -> { success, data: { preferredLanguage: "es" | null,
 *                               supportedLanguages: ["en", "es"] } }
 *   PATCH /public/users/preferred-language?domain=acme.org&email=jane@acme.org&lang=es
 *         -> 200, or 400 when `lang` is not in the account's supported list
 *
 * Base URLs per environment come from env.gs.
 */

var PREFERRED_LANGUAGE_PATH = "/public/users/preferred-language";

var JSON_HEADERS = { "Content-Type": "application/json" };

/** Base URL of the unification platform for an AWS region, in the active environment. */
function unificationBase(region) {
  var bases = env().unification;
  var base = bases[region] || bases["us-east-1"];
  console.log("unificationBase|", { region: region, base: base });
  return base;
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
 * One call gives everything the selector needs:
 *   { preferredLanguage: string | null, supportedLanguages: string[] }
 * `preferredLanguage` null is the normal first-visit answer. Returns null
 * when the route is unavailable or the response has no supportedLanguages
 * array; the caller then hides the selector.
 */
function getLanguagePreference(identity) {
  var query = identityQuery(identity);
  if (!query) {
    console.log("getLanguagePreference|skipped|no email or memberId");
    return null;
  }
  var url = unificationBase(identity.region) + PREFERRED_LANGUAGE_PATH + "?" + query;
  console.log("getLanguagePreference|called|", { url: url });
  try {
    var res = UrlFetchApp.fetch(url, { method: "get", headers: JSON_HEADERS, muteHttpExceptions: true });
    var code = res.getResponseCode();
    if (code !== 200) {
      // Logged as a message_code, not as the API's English: the selector just
      // hides itself, but the code is what makes a silent hide diagnosable.
      var failure = resolveApiMessage(code, res.getContentText());
      console.warn("getLanguagePreference|http failure|", { status: code, code: failure.code, detail: failure.detail });
      return null;
    }
    var body = parseJsonSafe(res.getContentText());
    var data = body && body.data;
    var supported = data && data.supportedLanguages;
    if (!Array.isArray(supported)) {
      console.warn("getLanguagePreference|no supportedLanguages array in response");
      return null;
    }
    var preferred = data.preferredLanguage;
    var result = {
      preferredLanguage: typeof preferred === "string" && preferred ? preferred : null,
      supportedLanguages: supported.filter(function (c) { return typeof c === "string"; }),
    };
    console.log("getLanguagePreference|result|", result);
    return result;
  } catch (err) {
    console.error("getLanguagePreference|failed|", err.message);
    callErrorReportingApi("getLanguagePreference failed: " + err.message, " ");
    return null;
  }
}

/**
 * Persists the user's choice on the account. `false` when it did not land
 * (network, or 400 for a language the account does not allow); the caller
 * still switches the UI, only the memory is lost.
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
    var res = UrlFetchApp.fetch(url, { method: "patch", headers: JSON_HEADERS, muteHttpExceptions: true });
    var code = res.getResponseCode();
    if (code !== 200) {
      var failure = resolveApiMessage(code, res.getContentText());
      console.warn("setPreferredLanguage|http failure|", { status: code, code: failure.code, detail: failure.detail });
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

/**
 * Run from the Apps Script editor (select it, Run, open the Execution log)
 * to see how the language routes answer for the signed-in account: the
 * environment, the resolved base URL, then status and body of each call.
 * Safe: the only write re-saves the preference that is already stored
 * (or "en").
 */
function debugLanguageApi() {
  var email = Session.getActiveUser().getEmail();
  var domain = accountDomainFromEmail(email);
  console.log("config", { email: email, domain: domain, environment: currentEnvName() });

  var reg = "us-east-1";
  try {
    var res = UrlFetchApp.fetch(env().userRegionHost + "/userregion?domain=" + domain, { muteHttpExceptions: true });
    console.log("userregion", { status: res.getResponseCode(), body: res.getContentText() });
    reg = JSON.parse(res.getContentText()).aws_region || reg;
  } catch (err) {
    console.log("userregion failed", err.message);
  }
  var base = unificationBase(reg);
  console.log("base URL", { region: reg, base: base });

  function call(label, method, url) {
    try {
      var r = UrlFetchApp.fetch(url, { method: method, headers: JSON_HEADERS, muteHttpExceptions: true });
      console.log(label, { method: method, url: url, status: r.getResponseCode(), body: r.getContentText() });
      return r;
    } catch (err) {
      // A whitelist miss surfaces here, as an exception rather than a status.
      console.log(label + " THREW", { method: method, url: url, error: err.message });
      return null;
    }
  }

  var identity = "domain=" + encodeURIComponent(domain) + "&email=" + encodeURIComponent(email);
  var pref = call("preferred-language GET", "get", base + PREFERRED_LANGUAGE_PATH + "?" + identity);
  var current = "en";
  try { current = JSON.parse(pref.getContentText()).data.preferredLanguage || "en"; } catch (err) {}
  call("preferred-language PATCH", "patch", base + PREFERRED_LANGUAGE_PATH + "?" + identity + "&lang=" + encodeURIComponent(current));
}
