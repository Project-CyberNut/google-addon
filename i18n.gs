/**
 * Localization for the CyberNut Gmail add-on.
 *
 * Mirrors user-portal-micro-learning-v2/src/i18n/config.ts and the language
 * selector on the cyb4-1364 branch, so both surfaces behave the same way:
 *
 *   options   = exactly the account's supported languages from the backend
 *               (admins add/remove them there; nothing here needs to change).
 *               Hidden when fewer than two remain, and hidden when the list
 *               cannot be read at all: no guessing on the backend's behalf.
 *   locale    = stored preference on the account
 *               -> last explicit choice on this device
 *               -> Gmail's UI language, if we ship it
 *               -> English
 *   on change = remember locally first (so a failed save never strands the
 *               user), then PATCH the preference to the account
 *
 * Card copy: MESSAGES holds the catalogues we ship. A language the backend
 * offers but we have no catalogue for still appears in the dropdown (name and
 * flag are derived from its code), is saved to the account, and renders the
 * card copy in English until a catalogue is added. Keep every catalogue's
 * keys in sync with `en`; `t()` falls back to English for a missing key.
 */

/**
 * Display overrides, keyed by language code, for when CLDR's derived answer
 * is not the one wanted. Normally empty: the native name comes from
 * Intl.DisplayNames, the flag's country from Intl.Locale#maximize (the same
 * CLDR rule the portal's LanguagePicker uses: en -> US, es -> ES, ar -> EG)
 * and the direction from the locale's text info.
 *   label    native name shown in the dropdown (never translated)
 *   country  ISO 3166 code of the flag to fly; "" for no flag
 *   rtl      true/false
 * Example: `"es": { country: "MX" }`
 */
var LOCALE_OVERRIDES = {};

/** Used only when `Intl.Locale` cannot tell us; CLDR knows the rest. */
var RTL_FALLBACK = ["ar", "he", "fa", "ur"];

/** A language tag we are willing to put in the dropdown: BCP 47 shaped. */
var LANGUAGE_TAG_PATTERN = /^[a-z]{2,3}(-[a-z0-9]{2,8})*$/i;

var DEFAULT_LOCALE = "en";

/** Languages we ship card copy for. */
function shippedLocales() {
  return Object.keys(MESSAGES);
}

/** Cache lifetimes (seconds). CacheService caps entries at 6 hours. */
var LOCALE_CACHE_TTL = 60 * 60;          // stored preference, re-read hourly
var SUPPORTED_CACHE_TTL = 6 * 60 * 60;   // account's supported list
var REGION_CACHE_TTL = 6 * 60 * 60;      // domain -> AWS region

var CACHE_KEY_PREFERRED = "cybernut_lang_preferred";
var CACHE_KEY_SUPPORTED = "cybernut_lang_supported:";
var CACHE_KEY_REGION = "cybernut_lang_region:";
var PROPERTY_KEY_LOCALE = "cybernut_locale";
/** Cached stand-in for "the account has nothing stored" (cache cannot hold null). */
var NONE_SENTINEL = "__none__";

var MESSAGES = {
  en: {
    common: {
      languageSelectLabel: "Select language",
      close: "Close",
    },
    home: {
      tagline: "Suspicious content or sender? Report it for further analysis.",
      reportButton: "Report Email",
      onboardingButton: "Onboarding Tutorial",
    },
    error: {
      generic: "There was an error in completing your action, for escalation / faster resolution you can contact us at support@cybernut.com",
    },
    step1: {
      openEmailPrompt: "Please open the email and look for the button in the top left corner. Click on it to go back and find the report button.",
      waitHeading: "<b>WAIT - Did you accidentally click on something in this email?</b>",
      reassurance: "<b>You will not get in trouble by telling us.</b><br/><br/>By sharing this information, it will help your IT department monitor and catch potential cyber attacks in your school district.<br/><br/>",
      selectPrompt: "Thank you for your cooperation and transparency.<br/><br/><b>Please select from the list below if applicable:</b> ",
      actions: {
        replied: "I replied to the email",
        downloaded: "I downloaded a file",
        openedAttachment: "I opened an attachment",
        visitedLink: "I visited a link",
        enteredPassword: "I entered my password",
        forwarded: "I forwarded the email",
        loggedIn: "I logged into a page",
        none: "None of the above",
      },
    },
    step2: {
      defaultConfirmation: "Thank you, you will hear back from IT if you need to take any further action.",
      movedToTrash: "Email moved to trash. Please refresh your Gmail view.",
    },
    language: {
      updated: "Language updated.",
      savedLocallyOnly: "Language changed, but your preference could not be saved to your account.",
      invalid: "That language is not available.",
    },
  },

  es: {
    common: {
      languageSelectLabel: "Seleccionar idioma",
      close: "Cerrar",
    },
    home: {
      tagline: "¿Contenido o remitente sospechoso? Repórtalo para un análisis más detallado.",
      reportButton: "Reportar correo",
      onboardingButton: "Tutorial de introducción",
    },
    error: {
      generic: "Se produjo un error al completar tu acción. Para escalar el caso o una resolución más rápida, contáctanos en support@cybernut.com",
    },
    step1: {
      openEmailPrompt: "Abre el correo y busca el botón en la esquina superior izquierda. Haz clic en él para volver y encontrar el botón de reportar.",
      waitHeading: "<b>ESPERA: ¿hiciste clic por accidente en algo de este correo?</b>",
      reassurance: "<b>No tendrás problemas por contárnoslo.</b><br/><br/>Al compartir esta información, ayudas a tu departamento de TI a supervisar y detectar posibles ciberataques en tu distrito escolar.<br/><br/>",
      selectPrompt: "Gracias por tu cooperación y transparencia.<br/><br/><b>Selecciona de la lista a continuación si corresponde:</b> ",
      actions: {
        replied: "Respondí al correo",
        downloaded: "Descargué un archivo",
        openedAttachment: "Abrí un archivo adjunto",
        visitedLink: "Visité un enlace",
        enteredPassword: "Ingresé mi contraseña",
        forwarded: "Reenvié el correo",
        loggedIn: "Inicié sesión en una página",
        none: "Ninguna de las anteriores",
      },
    },
    step2: {
      defaultConfirmation: "Gracias. El equipo de TI se comunicará contigo si necesitas tomar alguna otra medida.",
      movedToTrash: "El correo se movió a la papelera. Actualiza tu vista de Gmail.",
    },
    language: {
      updated: "Idioma actualizado.",
      savedLocallyOnly: "El idioma cambió, pero no se pudo guardar tu preferencia en tu cuenta.",
      invalid: "Ese idioma no está disponible.",
    },
  },

  ar: {
    common: {
      languageSelectLabel: "اختيار اللغة",
      close: "إغلاق",
    },
    home: {
      tagline: "محتوى أو مُرسِل مشبوه؟ أبلغ عنه لتحليله بشكل أعمق.",
      reportButton: "الإبلاغ عن البريد",
      onboardingButton: "الدرس التمهيدي",
    },
    error: {
      generic: "حدث خطأ أثناء إكمال الإجراء. للتصعيد أو للحصول على حل أسرع، تواصل معنا عبر support@cybernut.com",
    },
    step1: {
      openEmailPrompt: "يرجى فتح البريد الإلكتروني والبحث عن الزر في الزاوية العلوية اليسرى. انقر عليه للرجوع والعثور على زر الإبلاغ.",
      waitHeading: "<b>انتظر - هل نقرت عن طريق الخطأ على شيء في هذا البريد؟</b>",
      reassurance: "<b>لن تقع في أي مشكلة بإخبارنا.</b><br/><br/>مشاركة هذه المعلومات تساعد قسم تقنية المعلومات لديك على مراقبة الهجمات السيبرانية المحتملة في منطقتك التعليمية واكتشافها.<br/><br/>",
      selectPrompt: "شكرًا لتعاونك وشفافيتك.<br/><br/><b>يرجى الاختيار من القائمة أدناه إن كان ذلك مناسبًا:</b> ",
      actions: {
        replied: "ردّيت على البريد الإلكتروني",
        downloaded: "نزّلت ملفًا",
        openedAttachment: "فتحت مرفقًا",
        visitedLink: "زرت رابطًا",
        enteredPassword: "أدخلت كلمة المرور",
        forwarded: "أعدت توجيه البريد الإلكتروني",
        loggedIn: "سجّلت الدخول إلى صفحة",
        none: "لا شيء مما سبق",
      },
    },
    step2: {
      defaultConfirmation: "شكرًا لك، سيتواصل معك قسم تقنية المعلومات إذا كان عليك اتخاذ أي إجراء إضافي.",
      movedToTrash: "تم نقل البريد الإلكتروني إلى سلة المهملات. يرجى تحديث صفحة Gmail.",
    },
    language: {
      updated: "تم تحديث اللغة.",
      savedLocallyOnly: "تم تغيير اللغة، لكن لم نتمكن من حفظ تفضيلك في حسابك.",
      invalid: "هذه اللغة غير متاحة.",
    },
  },
};

// ---------------------------------------------------------------------------
// Registry helpers
// ---------------------------------------------------------------------------

/** True for a well-formed language tag, whatever the backend chooses to send. */
function isLocale(value) {
  return typeof value === "string" && LANGUAGE_TAG_PATTERN.test(value.trim());
}

/** Canonical form of a tag: trimmed, `_` -> `-`, lowercase. Null when malformed. */
function normalizeLocale(value) {
  if (!value) return null;
  var tag = String(value).trim().replace(/_/g, "-").toLowerCase();
  return isLocale(tag) ? tag : null;
}

/** `es-mx` -> `es`. */
function baseLanguage(code) {
  return String(code).toLowerCase().split("-")[0];
}

/**
 * The option that stands for `code`: an exact match first, else the option
 * with the same base language (device `es-MX` -> option `es`). Null when the
 * account offers nothing for it.
 */
function matchOption(options, code) {
  var normalized = normalizeLocale(code);
  if (!normalized) return null;
  for (var i = 0; i < options.length; i++) {
    if (options[i].toLowerCase() === normalized) return options[i];
  }
  var base = baseLanguage(normalized);
  for (var j = 0; j < options.length; j++) {
    if (baseLanguage(options[j]) === base) return options[j];
  }
  return null;
}

function localeOverride(code) {
  return LOCALE_OVERRIDES[code] || LOCALE_OVERRIDES[baseLanguage(code)] || null;
}

/**
 * Native name of the language, e.g. "Español", "العربية": what the portal
 * shows, so every reader can find their own language whatever the page is
 * currently in. Override wins; otherwise CLDR via Intl.DisplayNames, in the
 * language itself; last resort the bare code.
 */
function localeLabel(code) {
  var override = localeOverride(code);
  if (override && override.label) return override.label;
  try {
    var name = new Intl.DisplayNames([code], { type: "language" }).of(code);
    if (name && name.toLowerCase() !== code.toLowerCase()) {
      // CLDR lowercases some (es -> "español"); a menu wants a capital.
      return name.charAt(0).toLocaleUpperCase(code) + name.slice(1);
    }
  } catch (err) {
    console.log("localeLabel|Intl.DisplayNames unavailable|", { code: code, error: err.message });
  }
  return code;
}

/**
 * Country of the flag to fly. A language is not a country, so this is CLDR's
 * "most likely region" for the tag (`maximize()`), exactly as the portal does
 * it; a tag that already names a region (`pt-BR`) keeps it. Only a two-letter
 * region is a country: `eo` maximizes to `001`, the world, which flies no
 * flag. Empty when unknown.
 */
function localeCountry(code) {
  var override = localeOverride(code);
  if (override && override.country !== undefined) return override.country || "";
  try {
    var region = new Intl.Locale(code).maximize().region;
    if (region && /^[A-Za-z]{2}$/.test(region)) return region.toUpperCase();
  } catch (err) {
    console.log("localeCountry|Intl.Locale unavailable|", { code: code, error: err.message });
  }
  return "";
}

/**
 * Flag emoji for the locale's country, built from the two regional-indicator
 * symbols (`US` -> 🇺🇸). CardService dropdown items are plain text, so an
 * emoji is the only way to show a flag there; Gmail renders it as an image on
 * web, Android and iOS. Windows has no flag glyphs and shows the two letters
 * instead, which still reads fine. Empty when there is no country.
 */
function localeFlag(code) {
  var country = localeCountry(code);
  if (!/^[A-Z]{2}$/.test(country)) return "";
  var base = 0x1f1e6 - 65; // regional indicator A minus "A"
  return String.fromCodePoint(base + country.charCodeAt(0), base + country.charCodeAt(1));
}

/** What a dropdown row shows: flag then native name, e.g. "🇪🇸 Español". */
function localeOptionLabel(code) {
  var flag = localeFlag(code);
  return flag ? flag + " " + localeLabel(code) : localeLabel(code);
}

function localeDirection(code) {
  var override = localeOverride(code);
  if (override && typeof override.rtl === "boolean") return override.rtl ? "rtl" : "ltr";
  try {
    // Newer V8 exposes CLDR's script direction on the locale itself.
    var loc = new Intl.Locale(code);
    var info = typeof loc.getTextInfo === "function" ? loc.getTextInfo() : loc.textInfo;
    if (info && info.direction) return info.direction;
  } catch (err) {
    // fall through to the static list
  }
  return RTL_FALLBACK.indexOf(baseLanguage(code)) !== -1 ? "rtl" : "ltr";
}

/**
 * Run this from the Apps Script editor to see what codes derive to on the
 * real runtime, so a missing Intl feature shows up here rather than in a
 * user's dropdown. Pass the codes to check; defaults to the shipped ones.
 */
function debugLocaleDerivation(codes) {
  (codes || shippedLocales()).forEach(function (code) {
    console.log(code, {
      label: localeLabel(code),
      country: localeCountry(code),
      flag: localeFlag(code),
      direction: localeDirection(code),
      option: localeOptionLabel(code),
      hasCatalogue: catalogueFor(code) !== null,
    });
  });
}

// ---------------------------------------------------------------------------
// Translation
// ---------------------------------------------------------------------------

/** Locale of the current execution. Set by initLocale(); read by t(). */
var currentLocale = DEFAULT_LOCALE;

/** The shipped catalogue for a tag: exact (`pt-br`), then base (`pt`), else null. */
function catalogueFor(code) {
  if (!code) return null;
  var lower = String(code).toLowerCase();
  var keys = Object.keys(MESSAGES);
  for (var i = 0; i < keys.length; i++) {
    if (keys[i].toLowerCase() === lower) return MESSAGES[keys[i]];
  }
  var base = baseLanguage(lower);
  for (var j = 0; j < keys.length; j++) {
    if (keys[j].toLowerCase() === base) return MESSAGES[keys[j]];
  }
  return null;
}

function lookupMessage(catalog, key) {
  var node = catalog;
  var parts = key.split(".");
  for (var i = 0; i < parts.length; i++) {
    if (node === null || typeof node !== "object" || !(parts[i] in node)) return undefined;
    node = node[parts[i]];
  }
  return typeof node === "string" ? node : undefined;
}

/**
 * Translate `key` ("step1.actions.replied") in the current locale.
 * `params` fills `{name}` placeholders. Missing keys fall back to English
 * and then to the key itself, so a gap in a catalogue never breaks a card.
 */
function t(key, params) {
  var catalogue = catalogueFor(currentLocale);
  var text = catalogue ? lookupMessage(catalogue, key) : undefined;
  if (text === undefined) {
    if (catalogue && currentLocale !== DEFAULT_LOCALE) {
      console.warn("t|missing translation|", { locale: currentLocale, key: key });
    }
    text = lookupMessage(MESSAGES[DEFAULT_LOCALE], key);
  }
  if (text === undefined) {
    console.warn("t|missing key|", { key: key });
    return key;
  }
  if (params) {
    Object.keys(params).forEach(function (name) {
      text = text.split("{" + name + "}").join(String(params[name]));
    });
  }
  return text;
}

// ---------------------------------------------------------------------------
// Resolution
// ---------------------------------------------------------------------------

/** Gmail's UI language for this user (`useLocaleFromApp` + script.locale scope). */
function deviceLocale(e) {
  var fromEvent = e && e.commonEventObject && e.commonEventObject.userLocale;
  var normalized = normalizeLocale(fromEvent);
  if (normalized) return normalized;
  try {
    return normalizeLocale(Session.getActiveUserLocale());
  } catch (err) {
    console.log("deviceLocale|Session.getActiveUserLocale failed|", err.message);
    return null;
  }
}

/** Last explicit choice made in this add-on, on this Google account. */
function rememberedLocale() {
  try {
    return normalizeLocale(PropertiesService.getUserProperties().getProperty(PROPERTY_KEY_LOCALE));
  } catch (err) {
    return null;
  }
}

/**
 * Record a choice locally: the durable user property, plus the preference
 * cache so the next card render in this hour sees it without a round trip.
 */
function rememberLocale(locale) {
  try {
    PropertiesService.getUserProperties().setProperty(PROPERTY_KEY_LOCALE, locale);
  } catch (err) {
    console.log("rememberLocale|user property write failed|", err.message);
  }
  try {
    CacheService.getUserCache().put(CACHE_KEY_PREFERRED, locale, LOCALE_CACHE_TTL);
  } catch (err) {
    console.log("rememberLocale|cache write failed|", err.message);
  }
}

/** Domain -> AWS region, cached; wraps the existing region() lookup. */
async function resolveRegionCached(domain) {
  var cache = CacheService.getUserCache();
  var key = CACHE_KEY_REGION + domain;
  var cached = cache.get(key);
  if (cached) return cached;
  var result = await region(domain);
  var reg = (result && result.aws_region) || "us-east-1";
  cache.put(key, reg, REGION_CACHE_TTL);
  return reg;
}

/** Account's supported list, cached per domain. `null` when unreadable. */
function supportedLanguagesCached(domain, reg) {
  var cache = CacheService.getUserCache();
  var key = CACHE_KEY_SUPPORTED + domain;
  var cached = cache.get(key);
  if (cached) {
    var parsed = parseJsonSafe(cached);
    if (Array.isArray(parsed)) return parsed;
  }
  var supported = getSupportedLanguages(domain, reg);
  // Only real answers are cached: an outage should be retried, not remembered.
  if (supported !== null) cache.put(key, JSON.stringify(supported), SUPPORTED_CACHE_TTL);
  return supported;
}

/** Stored preference for this user, cached. `null` when none or unreadable. */
function preferredLanguageCached(identity) {
  var cache = CacheService.getUserCache();
  var cached = cache.get(CACHE_KEY_PREFERRED);
  if (cached === NONE_SENTINEL) return null;
  if (cached) return cached;
  var preferred = getPreferredLanguage(identity);
  cache.put(CACHE_KEY_PREFERRED, preferred || NONE_SENTINEL, LOCALE_CACHE_TTL);
  return preferred;
}

/**
 * What the dropdown offers: exactly the account's supported list, in the
 * backend's order, with the backend's own codes (the PATCH only accepts
 * those). Malformed entries and duplicates are dropped. Unreadable list
 * (network, 401 missing key, 404 unknown domain) or configured-empty ->
 * English only, which hides the dropdown: the backend is the only source of
 * the offer, so when it cannot answer we offer nothing rather than a guess.
 */
function resolveOptions(supported) {
  if (supported === null) {
    console.warn("resolveOptions|supported list unreadable - selector hidden, English used");
    return [DEFAULT_LOCALE];
  }
  var options = [];
  var seen = {};
  supported.forEach(function (code) {
    var trimmed = typeof code === "string" ? code.trim() : "";
    var key = trimmed.toLowerCase();
    if (!isLocale(trimmed) || seen[key]) {
      if (trimmed) console.warn("resolveOptions|dropping entry|", { code: code });
      return;
    }
    seen[key] = true;
    options.push(trimmed);
  });
  var untranslated = options.filter(function (c) { return catalogueFor(c) === null; });
  if (untranslated.length > 0) {
    console.warn("resolveOptions|offered without a shipped catalogue - card copy falls back to English|", {
      untranslated: untranslated,
    });
  }
  return options.length > 0 ? options : [DEFAULT_LOCALE];
}

/** Memoised per execution: every card in one run shares one resolution. */
var languageContextMemo = null;

/**
 * Everything a card needs to speak the user's language and offer a switch:
 *   { locale, options, identity }
 * `identity` is null when the active user's email is unavailable; the
 * dropdown is then hidden and the locale comes from the device.
 *
 * Also sets `currentLocale`, so `t()` works for the rest of the execution.
 */
async function getLanguageContext(e) {
  if (languageContextMemo) return languageContextMemo;

  var email = null;
  try {
    email = Session.getActiveUser().getEmail();
  } catch (err) {
    console.log("getLanguageContext|no active user email|", err.message);
  }
  var domain = accountDomainFromEmail(email);

  var identity = null;
  var supported = null;
  var preferred = null;
  if (domain) {
    var reg = await resolveRegionCached(domain);
    identity = { domain: domain, email: email, region: reg };
    supported = supportedLanguagesCached(domain, reg);
    preferred = preferredLanguageCached(identity);
  } else {
    console.warn("getLanguageContext|no account domain - language selector hidden");
  }

  var options = resolveOptions(supported);
  var remembered = rememberedLocale();
  var device = deviceLocale(e);

  var locale = matchOption(options, preferred) ||
    matchOption(options, remembered) ||
    matchOption(options, device) ||
    matchOption(options, DEFAULT_LOCALE) ||
    options[0];

  currentLocale = locale;
  languageContextMemo = { locale: locale, options: options, identity: identity };
  console.log("getLanguageContext|resolved|", {
    domain: domain,
    supported: supported === null ? "unreadable (selector hidden)" : supported,
    offered: options,
    preferred: preferred || "none stored",
    remembered: remembered || "none",
    device: device || "unsupported",
    locale: locale,
  });
  return languageContextMemo;
}

/** Cheap entry-point hook: resolve the locale so `t()` speaks the right language. */
async function initLocale(e) {
  var ctx = await getLanguageContext(e);
  return ctx.locale;
}

// ---------------------------------------------------------------------------
// UI
// ---------------------------------------------------------------------------

/** Read a form field from either event shape Gmail add-ons produce. */
function readFormInput(e, name) {
  var modern = e && e.commonEventObject && e.commonEventObject.formInputs &&
    e.commonEventObject.formInputs[name];
  if (modern && modern.stringInputs && modern.stringInputs.value && modern.stringInputs.value.length) {
    return modern.stringInputs.value[0];
  }
  var legacy = e && e.formInputs && e.formInputs[name];
  if (legacy && legacy.length) return legacy[0];
  return null;
}

/**
 * The language dropdown, or null when the account leaves the user no choice
 * (one language is no choice at all, so nothing is rendered).
 */
function buildLanguageSelector(ctx) {
  if (!ctx || !ctx.identity || ctx.options.length < 2) return null;
  var dropdown = CardService.newSelectionInput()
    .setType(CardService.SelectionInputType.DROPDOWN)
    .setFieldName("language")
    .setTitle(t("common.languageSelectLabel"))
    .setOnChangeAction(CardService.newAction().setFunctionName("onLanguageChange"));
  ctx.options.forEach(function (code) {
    dropdown.addItem(localeOptionLabel(code), code, code === ctx.locale);
  });
  return dropdown;
}

/**
 * Dropdown handler. Same write order as the portal's saveLanguageChoice():
 * remember locally first (so a rejected or unreachable PATCH never strands
 * the user on a language they just left), then persist on the account, then
 * redraw the home card in the new language.
 */
async function onLanguageChange(e) {
  console.log("onLanguageChange|called!");
  var ctx = await getLanguageContext(e);
  // The dropdown value is the backend's own code; keep it verbatim, since
  // the PATCH accepts only codes from the account's supported list.
  var raw = readFormInput(e, "language");
  var chosen = raw && ctx.options.indexOf(raw) !== -1 ? raw : null;
  console.log("onLanguageChange|", { raw: raw, chosen: chosen, offered: ctx.options });

  if (!chosen) {
    return CardService.newActionResponseBuilder()
      .setNotification(CardService.newNotification().setText(t("language.invalid")))
      .build();
  }

  rememberLocale(chosen);
  currentLocale = chosen;
  languageContextMemo = { locale: chosen, options: ctx.options, identity: ctx.identity };

  var saved = false;
  if (ctx.identity) saved = setPreferredLanguage(ctx.identity, chosen);
  if (!saved) {
    // Keep the local choice authoritative until the account catches up, but
    // don't cache "chosen" as the account's answer for a full hour.
    try { CacheService.getUserCache().remove(CACHE_KEY_PREFERRED); } catch (err) {}
  }

  var card = await buildHomeCard(e);
  return CardService.newActionResponseBuilder()
    .setNavigation(CardService.newNavigation().updateCard(card))
    .setNotification(CardService.newNotification().setText(
      saved ? t("language.updated") : t("language.savedLocallyOnly")
    ))
    .build();
}
