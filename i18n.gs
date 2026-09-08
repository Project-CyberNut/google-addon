/**
 * Localization for the CyberNut Gmail add-on.
 *
 * Mirrors user-portal-micro-learning-v2/src/i18n/config.ts and the language
 * selector on the cyb4-1364 branch, so both surfaces behave the same way:
 *
 *   options   = the account's supported languages, narrowed to the ones we
 *               ship a catalogue for (all shipped languages if the list is
 *               unreadable; hidden entirely when fewer than two remain)
 *   locale    = stored preference on the account
 *               -> last explicit choice on this device
 *               -> Gmail's UI language, if we ship it
 *               -> English
 *   on change = remember locally first (so a failed save never strands the
 *               user), then PATCH the preference to the account
 *
 * Adding a language: add one entry to LOCALES and one catalogue to MESSAGES.
 * Nothing else changes. Keep every catalogue key in sync with `en`; `t()`
 * falls back to English for a missing key and logs it.
 */

/**
 * Registry. `label` is the native name on purpose (never translated).
 * `country` is the flag to fly, as an ISO 3166 code: a language is not a
 * country, so this follows the portal's LanguagePicker (CLDR's most likely
 * region: en -> US, es -> ES, ar -> EG). Leave it empty for no flag.
 */
var LOCALES = [
  { code: "en", label: "English", country: "US", rtl: false },
  { code: "es", label: "Español", country: "ES", rtl: false },
  { code: "ar", label: "العربية", country: "EG", rtl: true },
];

var DEFAULT_LOCALE = "en";

var LOCALE_CODES = LOCALES.map(function (l) { return l.code; });

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

function isLocale(value) {
  return typeof value === "string" && LOCALE_CODES.indexOf(value) !== -1;
}

/** `es-MX` / `en_US` -> `es` / `en`, or null when we ship no catalogue for it. */
function normalizeLocale(value) {
  if (!value) return null;
  var base = String(value).trim().toLowerCase().split(/[-_]/)[0];
  return isLocale(base) ? base : null;
}

function localeEntry(code) {
  for (var i = 0; i < LOCALES.length; i++) {
    if (LOCALES[i].code === code) return LOCALES[i];
  }
  return null;
}

function localeLabel(code) {
  var entry = localeEntry(code);
  return entry ? entry.label : code;
}

/**
 * Flag emoji for the locale's country, built from the two regional-indicator
 * symbols (`US` -> 🇺🇸). CardService dropdown items are plain text, so an
 * emoji is the only way to show a flag there; Gmail renders it as an image on
 * web, Android and iOS. Windows has no flag glyphs and shows the two letters
 * instead, which still reads fine. Empty when the entry has no country.
 */
function localeFlag(code) {
  var entry = localeEntry(code);
  var country = entry && entry.country;
  if (!country || !/^[A-Za-z]{2}$/.test(country)) return "";
  var upper = country.toUpperCase();
  var base = 0x1f1e6 - 65; // regional indicator A minus "A"
  return String.fromCodePoint(base + upper.charCodeAt(0), base + upper.charCodeAt(1));
}

/** What a dropdown row shows: flag then native name, e.g. "🇪🇸 Español". */
function localeOptionLabel(code) {
  var flag = localeFlag(code);
  return flag ? flag + " " + localeLabel(code) : localeLabel(code);
}

function localeDirection(code) {
  for (var i = 0; i < LOCALES.length; i++) {
    if (LOCALES[i].code === code) return LOCALES[i].rtl ? "rtl" : "ltr";
  }
  return "ltr";
}

// ---------------------------------------------------------------------------
// Translation
// ---------------------------------------------------------------------------

/** Locale of the current execution. Set by initLocale(); read by t(). */
var currentLocale = DEFAULT_LOCALE;

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
  var text = lookupMessage(MESSAGES[currentLocale] || {}, key);
  if (text === undefined) {
    if (currentLocale !== DEFAULT_LOCALE) {
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
 * Languages we can put on screen: the account's list narrowed to the ones
 * we ship. Unreadable list -> everything we ship. Configured-empty or nothing
 * we can render -> English only (which hides the dropdown).
 */
function resolveOptions(supported) {
  if (supported === null) return LOCALE_CODES.slice();
  var options = supported.filter(isLocale);
  var untranslated = supported.filter(function (c) { return !isLocale(c); });
  if (untranslated.length > 0) {
    console.warn("resolveOptions|account allows languages the add-on does not ship|", {
      untranslated: untranslated, offered: options,
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

  var locale;
  if (isLocale(preferred) && options.indexOf(preferred) !== -1) locale = preferred;
  else if (remembered && options.indexOf(remembered) !== -1) locale = remembered;
  else if (device && options.indexOf(device) !== -1) locale = device;
  else if (options.indexOf(DEFAULT_LOCALE) !== -1) locale = DEFAULT_LOCALE;
  else locale = options[0];

  currentLocale = locale;
  languageContextMemo = { locale: locale, options: options, identity: identity };
  console.log("getLanguageContext|resolved|", {
    domain: domain,
    supported: supported === null ? "unreadable (offering every shipped language)" : supported,
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
  var chosen = normalizeLocale(readFormInput(e, "language"));
  console.log("onLanguageChange|", { chosen: chosen, offered: ctx.options });

  if (!chosen || ctx.options.indexOf(chosen) === -1) {
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
