/**
 * API message codes and the Condition text behind each one, in every language
 * the add-on ships.
 *
 * The backends (`email-template`, `user-portal-api`,
 * `cybernut-platform-unification`) promise a stable `message_code` on their
 * responses and treat the English beside it as legacy compatibility copy. So
 * this file never shows that English: it resolves the code against the
 * catalogue below and the user reads the sentence in their own language.
 *
 * The sentences are the Condition column of API_MESSAGE_CODE_CATALOG.md, not a
 * rewrite of it — the same wording the backend catalogue uses to describe when
 * a code fires. Field and parameter names (partitionKey, sortKey, messageId,
 * sessionId) stay untranslated in every language: they are identifiers, not
 * words, and translating them would leave the reader with a name that matches
 * nothing in the request they made.
 *
 * Several codes are raised by more than one endpoint with differently worded
 * Conditions. There is one entry per code, generalized to a sentence true of
 * every occurrence, with the per-endpoint variants noted above it so the
 * mapping back to the catalogue stays traceable.
 *
 * Kept out of i18n.gs deliberately: that file is already the locale registry,
 * the resolution chain and the card copy, and this catalogue is four times the
 * size of all of its copy put together.
 */

/**
 * Codes the backends send, plus two the add-on raises itself for failures that
 * never reach a backend (NETWORK_ERROR, UNKNOWN_ERROR — see FALLBACK_MESSAGE_CODE).
 *
 * Keys are in the same order in all three locales so the three read as one
 * table in review; `debugApiCodeCatalogue()` enforces that they stay in step.
 */
var API_CODE_CONDITIONS = {
  en: {
    // --- Request validation -------------------------------------------------
    // Variants: "partitionkey is absent. If both required keys are absent,
    // this code wins." / "This has first precedence among missing inputs." /
    // "Neither sessionId nor partitionKey was supplied."
    MISSING_PARTITION_KEY: "partitionKey is absent, and where sessionId is accepted, neither was supplied.",
    MISSING_SORT_KEY: "partitionkey is present but sortkey is absent.",
    MISSING_MODULE: "partitionkey and sortkey are present but module is absent.",
    // Variants: "messageId is absent or empty." / "A non-demo event has no messageId."
    MISSING_MESSAGE_ID: "messageId is absent or empty.",
    // Variants: "Neither messageId nor the required key-based identifiers were
    // supplied." / "No usable messageId, sessionId, or partitionKey + sortKey."
    MISSING_TEMPLATE_IDENTIFIERS: "No usable combination of messageId, sessionId, or partitionKey + sortKey was supplied.",
    MISSING_DESTINATION_EMAIL: "The partitionKey flow has no destinationEmail.",
    MISSING_USER_KEY: "None of the accepted user-key combinations can be built: partitionKey + email, partitionKey + sortKey, or domain + email.",
    MISSING_CLASSES: "classes is absent, is not an array, or becomes empty after normalization.",
    MISSING_EVENT_SUBJECT: "Neither an action nor an onboarding slide identifier can identify the subject.",
    // Variants: "A non-legacy partition key cannot be parsed as a campaign
    // UUID." / "partitionKey cannot be parsed as a campaign UUID."
    INVALID_PARTITION_KEY: "partitionKey cannot be parsed as a campaign UUID.",
    INVALID_JSON_BODY: "The request body is not valid JSON.",
    INVALID_TUTORIAL_STATUS: "The status is not completed or skipped, or the boolean flags are invalid.",
    // Variants: the endpoint-specific "DTO validation fails, including a
    // missing/invalid action or eventTimestamp", and the platform filter's
    // default for a bare 400 or 422.
    VALIDATION_ERROR: "The request failed validation, including a missing or invalid action or eventTimestamp.",

    // --- Nothing to serve ---------------------------------------------------
    // Variants: "No campaign-user result was found, including after
    // sibling-subcampaign fallback." / "No campaign-user result matched
    // messageId." / "No campaign result matches the submitted messageId." /
    // "The notification, schedule, member, or campaign for messageId cannot be
    // resolved."
    CAMPAIGN_RECORD_NOT_FOUND: "No campaign record could be resolved for the supplied identifiers.",
    // Variants: "No campaign-user row exists for the resolved key." / "The
    // conditional update found no campaign-user row."
    CAMPAIGN_USER_NOT_FOUND: "No campaign-user row exists for the resolved key.",
    // Variants: "No email template/send record could be resolved." / "No
    // eligible onboarding microtraining schedule/template was found."
    TEMPLATE_NOT_FOUND: "No eligible template or send record could be resolved.",
    ACCOUNT_NOT_FOUND: "No account exists for partitionKey.",
    // Variants: "The supplied demo sessionId cannot be resolved." / "sessionId
    // does not identify a demo session."
    DEMO_SESSION_NOT_FOUND: "The supplied sessionId does not identify a demo session.",
    ONBOARDING_SLIDE_NOT_FOUND: "No slide progress exists for the supplied user or demo session.",
    CLASS_ANSWER_NOT_FOUND: "The category catalogue contains no correct answer for {class}.",
    NOT_FOUND: "The requested resource was not found.",

    // --- Refused ------------------------------------------------------------
    UNAUTHORIZED: "The request is not authorized.",
    FORBIDDEN: "Access to this resource is forbidden.",
    CONFLICT: "The request conflicts with the current state of the resource.",
    RATE_LIMITED: "Too many requests were sent in a short period.",
    // Variants: the handler's own "receives a method other than GET or POST",
    // and the platform filter's default for a bare 405.
    METHOD_NOT_ALLOWED: "The HTTP method is not allowed for this endpoint.",
    REQUEST_TIMEOUT: "The request timed out.",

    // --- Campaign wiring ----------------------------------------------------
    CAMPAIGN_TYPE_MISMATCH: "The submitted onboarding campaign type conflicts with the stored campaign.",
    UNSUPPORTED_ONBOARDING_ACTION: "The legacy action {action} is not supported.",

    // --- Success and state --------------------------------------------------
    // Variants: "The result event was published for all matched records." /
    // "The event was newly recorded or an idempotent duplicate was accepted."
    EVENT_RECORDED: "The event was recorded for all matched records.",
    TUTORIAL_STATUS_UPDATED: "Tutorial flags were updated.",
    EMAIL_ALREADY_REPORTED: "A report action already exists for this email.",
    EMAIL_NOT_REPORTED: "No report action exists for this email.",
    TRAINING_COMPLETED: "A training-completed action exists.",
    TRAINING_NOT_COMPLETED: "No training-completed action exists.",
    // Variants: "Campaign-specific playable content was found." / "Playable
    // content for the requested module was found."
    COMPLIANCE_CONTENT_RETURNED: "Playable content was found and returned.",
    // Variants: "No playable content was found after locale
    // fallback/filtering, or the unmatched-campaign fallback partition is
    // empty." / "No playable module content was found after locale fallback."
    COMPLIANCE_CONTENT_EMPTY: "No playable content was found after locale fallback and filtering.",
    COMPLIANCE_CONTENT_UNFILTERED: "No campaign row matched sortkey; the entire content partition was returned instead, and it is not campaign-specific playable content.",

    // --- Last resort --------------------------------------------------------
    SERVICE_MISCONFIGURED: "A required content or campaign table environment variable is missing.",
    // Variants: every backend's "a DynamoDB or other unexpected operation
    // failed", "an unexpected read/update/lookup failure occurred", and the
    // platform filter's default for an unrecognized status.
    INTERNAL_ERROR: "An unexpected server-side operation failed.",

    // --- Raised by the add-on itself, never by a backend ---------------------
    NETWORK_ERROR: "The server could not be reached.",
    UNKNOWN_ERROR: "The server returned a code this version does not recognise."
  },

  es: {
    MISSING_PARTITION_KEY: "partitionKey está ausente y, donde se acepta sessionId, tampoco se proporcionó.",
    MISSING_SORT_KEY: "partitionkey está presente pero falta sortkey.",
    MISSING_MODULE: "partitionkey y sortkey están presentes pero falta module.",
    MISSING_MESSAGE_ID: "messageId está ausente o vacío.",
    MISSING_TEMPLATE_IDENTIFIERS: "No se proporcionó ninguna combinación utilizable de messageId, sessionId o partitionKey + sortKey.",
    MISSING_DESTINATION_EMAIL: "El flujo de partitionKey no tiene destinationEmail.",
    MISSING_USER_KEY: "No se puede construir ninguna de las combinaciones de clave de usuario aceptadas: partitionKey + email, partitionKey + sortKey o domain + email.",
    MISSING_CLASSES: "classes está ausente, no es un arreglo o queda vacío tras la normalización.",
    MISSING_EVENT_SUBJECT: "Ni una acción ni un identificador de diapositiva de incorporación permiten identificar el sujeto.",
    INVALID_PARTITION_KEY: "partitionKey no se puede interpretar como un UUID de campaña.",
    INVALID_JSON_BODY: "El cuerpo de la solicitud no es JSON válido.",
    INVALID_TUTORIAL_STATUS: "El estado no es completed ni skipped, o los indicadores booleanos no son válidos.",
    VALIDATION_ERROR: "La solicitud no superó la validación, incluidos un action o un eventTimestamp ausentes o no válidos.",

    CAMPAIGN_RECORD_NOT_FOUND: "No se pudo resolver ningún registro de campaña con los identificadores proporcionados.",
    CAMPAIGN_USER_NOT_FOUND: "No existe ninguna fila de usuario de campaña para la clave resuelta.",
    TEMPLATE_NOT_FOUND: "No se pudo resolver ninguna plantilla ni registro de envío válido.",
    ACCOUNT_NOT_FOUND: "No existe ninguna cuenta para partitionKey.",
    DEMO_SESSION_NOT_FOUND: "El sessionId proporcionado no identifica ninguna sesión de demostración.",
    ONBOARDING_SLIDE_NOT_FOUND: "No existe progreso de diapositivas para el usuario o la sesión de demostración indicados.",
    CLASS_ANSWER_NOT_FOUND: "El catálogo de categorías no contiene una respuesta correcta para {class}.",
    NOT_FOUND: "No se encontró el recurso solicitado.",

    UNAUTHORIZED: "La solicitud no está autorizada.",
    FORBIDDEN: "El acceso a este recurso está prohibido.",
    CONFLICT: "La solicitud entra en conflicto con el estado actual del recurso.",
    RATE_LIMITED: "Se enviaron demasiadas solicitudes en poco tiempo.",
    METHOD_NOT_ALLOWED: "El método HTTP no está permitido para este endpoint.",
    REQUEST_TIMEOUT: "La solicitud excedió el tiempo de espera.",

    CAMPAIGN_TYPE_MISMATCH: "El tipo de campaña de incorporación enviado no coincide con la campaña almacenada.",
    UNSUPPORTED_ONBOARDING_ACTION: "La acción heredada {action} no es compatible.",

    EVENT_RECORDED: "El evento se registró para todos los registros coincidentes.",
    TUTORIAL_STATUS_UPDATED: "Se actualizaron los indicadores del tutorial.",
    EMAIL_ALREADY_REPORTED: "Ya existe una acción de reporte para este correo.",
    EMAIL_NOT_REPORTED: "No existe ninguna acción de reporte para este correo.",
    TRAINING_COMPLETED: "Existe una acción de formación completada.",
    TRAINING_NOT_COMPLETED: "No existe ninguna acción de formación completada.",
    COMPLIANCE_CONTENT_RETURNED: "Se encontró y devolvió contenido reproducible.",
    COMPLIANCE_CONTENT_EMPTY: "No se encontró contenido reproducible tras la reserva de idioma y el filtrado.",
    COMPLIANCE_CONTENT_UNFILTERED: "Ninguna fila de campaña coincidió con sortkey; en su lugar se devolvió toda la partición de contenido, que no es contenido reproducible específico de la campaña.",

    SERVICE_MISCONFIGURED: "Falta una variable de entorno obligatoria de la tabla de contenido o de campañas.",
    INTERNAL_ERROR: "Falló una operación inesperada en el servidor.",

    NETWORK_ERROR: "No se pudo contactar con el servidor.",
    UNKNOWN_ERROR: "El servidor devolvió un código que esta versión no reconoce."
  },

  ar: {
    MISSING_PARTITION_KEY: "المعامل partitionKey مفقود، وحيث يُقبل sessionId لم يُرسَل أي منهما.",
    MISSING_SORT_KEY: "المعامل partitionkey موجود ولكن sortkey مفقود.",
    MISSING_MODULE: "المعاملان partitionkey و sortkey موجودان ولكن module مفقود.",
    MISSING_MESSAGE_ID: "المعرّف messageId مفقود أو فارغ.",
    MISSING_TEMPLATE_IDENTIFIERS: "لم تُرسَل أي تركيبة صالحة من messageId أو sessionId أو partitionKey + sortKey.",
    MISSING_DESTINATION_EMAIL: "مسار partitionKey لا يحتوي على destinationEmail.",
    MISSING_USER_KEY: "تعذّر تكوين أي من تركيبات مفتاح المستخدم المقبولة: partitionKey + email أو partitionKey + sortKey أو domain + email.",
    MISSING_CLASSES: "العنصر classes مفقود أو ليس مصفوفة أو أصبح فارغاً بعد التنسيق.",
    MISSING_EVENT_SUBJECT: "لا يمكن تحديد الموضوع بواسطة إجراء ولا بمعرّف شريحة تهيئة.",
    INVALID_PARTITION_KEY: "تعذّر تحليل partitionKey كمعرّف UUID لحملة.",
    INVALID_JSON_BODY: "محتوى الطلب ليس بصيغة JSON صالحة.",
    INVALID_TUTORIAL_STATUS: "الحالة ليست completed أو skipped، أو أن القيم المنطقية غير صالحة.",
    VALIDATION_ERROR: "فشل الطلب في التحقق من الصحة، بما في ذلك action أو eventTimestamp مفقود أو غير صالح.",

    CAMPAIGN_RECORD_NOT_FOUND: "تعذّر العثور على سجل حملة مطابق للمعرّفات المُرسَلة.",
    CAMPAIGN_USER_NOT_FOUND: "لا يوجد سجل مستخدم حملة للمفتاح المُستخرَج.",
    TEMPLATE_NOT_FOUND: "تعذّر العثور على قالب أو سجل إرسال مؤهل.",
    ACCOUNT_NOT_FOUND: "لا يوجد حساب لـ partitionKey.",
    DEMO_SESSION_NOT_FOUND: "لا يشير sessionId المُرسَل إلى أي جلسة تجريبية.",
    ONBOARDING_SLIDE_NOT_FOUND: "لا يوجد تقدّم في الشرائح للمستخدم أو الجلسة التجريبية المحددة.",
    CLASS_ANSWER_NOT_FOUND: "لا يحتوي كتالوج الفئات على إجابة صحيحة لـ {class}.",
    NOT_FOUND: "لم يُعثر على المورد المطلوب.",

    UNAUTHORIZED: "الطلب غير مُصرَّح به.",
    FORBIDDEN: "الوصول إلى هذا المورد محظور.",
    CONFLICT: "يتعارض الطلب مع الحالة الحالية للمورد.",
    RATE_LIMITED: "أُرسلت طلبات كثيرة جداً خلال فترة قصيرة.",
    METHOD_NOT_ALLOWED: "طريقة HTTP غير مسموح بها لهذه الواجهة.",
    REQUEST_TIMEOUT: "انتهت مهلة الطلب.",

    CAMPAIGN_TYPE_MISMATCH: "نوع حملة التهيئة المُرسَل لا يتطابق مع الحملة المخزّنة.",
    UNSUPPORTED_ONBOARDING_ACTION: "الإجراء القديم {action} غير مدعوم.",

    EVENT_RECORDED: "تم تسجيل الحدث لجميع السجلات المطابقة.",
    TUTORIAL_STATUS_UPDATED: "تم تحديث مؤشرات الدرس التمهيدي.",
    EMAIL_ALREADY_REPORTED: "يوجد بالفعل إجراء إبلاغ لهذا البريد الإلكتروني.",
    EMAIL_NOT_REPORTED: "لا يوجد إجراء إبلاغ لهذا البريد الإلكتروني.",
    TRAINING_COMPLETED: "يوجد إجراء يفيد بإكمال التدريب.",
    TRAINING_NOT_COMPLETED: "لا يوجد إجراء يفيد بإكمال التدريب.",
    COMPLIANCE_CONTENT_RETURNED: "تم العثور على محتوى قابل للتشغيل وإعادته.",
    COMPLIANCE_CONTENT_EMPTY: "لم يُعثر على محتوى قابل للتشغيل بعد الرجوع إلى اللغة البديلة والتصفية.",
    COMPLIANCE_CONTENT_UNFILTERED: "لم تتطابق أي حملة مع sortkey؛ أُعيد بدلاً من ذلك قسم المحتوى بالكامل، وهو ليس محتوى قابلاً للتشغيل خاصاً بالحملة.",

    SERVICE_MISCONFIGURED: "متغيّر بيئة مطلوب لجدول المحتوى أو الحملات مفقود.",
    INTERNAL_ERROR: "فشلت عملية غير متوقعة على الخادم.",

    NETWORK_ERROR: "تعذّر الوصول إلى الخادم.",
    UNKNOWN_ERROR: "أعاد الخادم رمزاً لا تتعرّف عليه هذه النسخة."
  }
};

/** Shown when a response carries a code this build has never heard of. */
var FALLBACK_MESSAGE_CODE = "UNKNOWN_ERROR";

/** The namespace the catalogue is registered under, so `t()` can reach it. */
var API_CODE_NAMESPACE = "apiCodes";

/**
 * The platform's global exception filter maps a bare HttpException's status to
 * a code. A response that carries no code of its own is read the same way here,
 * so an uncoded 401 still says "not authorized" rather than the generic
 * server-error sentence.
 */
var STATUS_MESSAGE_CODES = {
  400: "VALIDATION_ERROR",
  401: "UNAUTHORIZED",
  403: "FORBIDDEN",
  404: "NOT_FOUND",
  405: "METHOD_NOT_ALLOWED",
  408: "REQUEST_TIMEOUT",
  409: "CONFLICT",
  422: "VALIDATION_ERROR",
  429: "RATE_LIMITED"
};

/** True for a code this build knows how to render. */
function isApiMessageCode(value) {
  return (
    typeof value === "string" &&
    Object.prototype.hasOwnProperty.call(API_CODE_CONDITIONS.en, value)
  );
}

/**
 * Merges the catalogue into `MESSAGES` under `apiCodes`, so every code is
 * reachable through the same `t()` the rest of the add-on uses — one lookup
 * path, one fallback chain, one interpolation rule.
 *
 * Done on first use rather than at load: Apps Script gives no guaranteed
 * evaluation order across .gs files, so `MESSAGES` may not exist yet while
 * this file's top level runs.
 */
function registerApiCodeCatalogue() {
  Object.keys(API_CODE_CONDITIONS).forEach(function (locale) {
    if (!MESSAGES[locale]) return;
    if (!MESSAGES[locale][API_CODE_NAMESPACE]) {
      MESSAGES[locale][API_CODE_NAMESPACE] = API_CODE_CONDITIONS[locale];
    }
  });
}

/**
 * The sentence for `code`, in the locale the current invocation resolved.
 * `params` fills `{class}` / `{action}` placeholders from `message_params`.
 * An unknown code renders the fallback sentence — a raw
 * "UNSUPPORTED_ONBOARDING_ACTION" on a user's card is worse than saying we did
 * not understand the answer.
 */
function apiMessage(code, params) {
  registerApiCodeCatalogue();
  var key = isApiMessageCode(code) ? code : FALLBACK_MESSAGE_CODE;
  return t(API_CODE_NAMESPACE + "." + key, params);
}

/** JSON.parse that answers null instead of throwing, for logging paths. */
function parseJsonSafely(text) {
  if (typeof text !== "string" || !text) return null;
  try {
    return JSON.parse(text);
  } catch (err) {
    return null;
  }
}

function isPlainObject(value) {
  return typeof value === "object" && value !== null && !Array.isArray(value);
}

/** Keeps only the scalar entries of `message_params` — the rest cannot interpolate. */
function readMessageParams(value) {
  if (!isPlainObject(value)) return undefined;
  var params = {};
  var found = false;
  Object.keys(value).forEach(function (name) {
    var entry = value[name];
    var type = typeof entry;
    if (type === "string" || type === "number" || type === "boolean") {
      params[name] = String(entry);
      found = true;
    }
  });
  return found ? params : undefined;
}

/**
 * Pulls the code out of a response body.
 *
 * Canonical fields are top-level `message_code` and `message_params`. The
 * transitional nested, camel-case and meta shapes stay readable during a
 * rolling deployment, but canonical fields always win. Returns null when the
 * body carries no code at all, so the caller can fall back on the status.
 */
function readMessageCode(body) {
  var root = isPlainObject(body) ? body : parseJsonSafely(body);
  if (!isPlainObject(root)) return null;

  var nested = isPlainObject(root.error) ? root.error : null;
  var meta = isPlainObject(root.meta) ? root.meta : null;
  var raw =
    root.message_code ||
    (nested && nested.message_code) ||
    root.errorCode ||
    (meta && meta.code);
  if (typeof raw !== "string") return null;

  var source = root.message_code ? root : nested && nested.message_code ? nested : root;

  return {
    code: isApiMessageCode(raw) ? raw : FALLBACK_MESSAGE_CODE,
    params: readMessageParams(
      root.message_params ||
        (nested && nested.message_params) ||
        source.params ||
        root.errorParams
    ),
    // The backend's own English. Logged, never rendered — but a log line
    // naming the exact sentence the API sent is what makes an unmapped code
    // quick to chase down.
    detail:
      typeof source.message === "string"
        ? source.message
        : typeof root.error === "string"
          ? root.error
          : undefined
  };
}

/** The code a response with no code of its own gets, from its HTTP status. */
function messageCodeForStatus(status) {
  return STATUS_MESSAGE_CODES[Number(status)] || "INTERNAL_ERROR";
}

/**
 * What a completed HTTP call means, as a code: the body's own code if it has
 * one, otherwise the status-mapped default (an uncoded 502, a gateway page).
 */
function resolveApiMessage(statusCode, bodyText) {
  var fromBody = readMessageCode(bodyText);
  if (fromBody) return fromBody;
  return { code: messageCodeForStatus(statusCode), params: undefined, detail: undefined };
}

/**
 * What a thrown failure means, as a code.
 *
 * An error already carrying a resolved code (one this file attached upstream)
 * keeps it. Anything else never reached a server, so it becomes NETWORK_ERROR
 * rather than the generic fallback: "we could not reach the server" and "the
 * server said something we do not understand" are different problems for the
 * person reading the card and for whoever is debugging it.
 */
function resolveApiMessageFromError(error) {
  if (error && isApiMessageCode(error.messageCode)) {
    return { code: error.messageCode, params: error.messageParams, detail: error.message };
  }
  return { code: "NETWORK_ERROR", params: undefined, detail: error && error.message };
}

/**
 * An Error that carries its resolved code, so a catch block far from the fetch
 * can still render the right sentence. `buildErrorCard` reads these two fields.
 */
function apiError(message, resolved) {
  var err = new Error(message);
  if (resolved) {
    err.messageCode = resolved.code;
    err.messageParams = resolved.params;
  }
  return err;
}

/**
 * Prints every code's sentence side by side in all three languages, for
 * reviewing the wording without having to provoke 42 real API failures. Run it
 * from the Apps Script editor and read the execution log.
 *
 * `codes` optionally narrows it, e.g. debugApiCodeMessages(["MISSING_SORT_KEY"]).
 */
function debugApiCodeMessages(codes) {
  var wanted = Array.isArray(codes) && codes.length
    ? codes
    : Object.keys(API_CODE_CONDITIONS.en);
  var locales = Object.keys(API_CODE_CONDITIONS);

  wanted.forEach(function (code) {
    if (!isApiMessageCode(code)) {
      console.log(code + " -> not a code this build knows");
      return;
    }
    locales.forEach(function (locale) {
      console.log(code + " [" + locale + "] " + API_CODE_CONDITIONS[locale][code]);
    });
  });
  return { codes: wanted.length, locales: locales };
}

/**
 * Locale-parity check for the catalogue. This project has no test runner — run
 * this once from the Apps Script editor after touching API_CODE_CONDITIONS.
 * Reports any locale that is missing a code, carries one English does not, or
 * drops an interpolation placeholder the English sentence has.
 */
function debugApiCodeCatalogue() {
  var base = API_CODE_CONDITIONS.en;
  var codes = Object.keys(base);
  var problems = [];

  var placeholders = function (text) {
    return (String(text).match(/\{[a-zA-Z_]+\}/g) || []).sort().join(",");
  };

  Object.keys(API_CODE_CONDITIONS).forEach(function (locale) {
    if (locale === "en") return;
    var catalogue = API_CODE_CONDITIONS[locale];
    codes.forEach(function (code) {
      if (typeof catalogue[code] !== "string" || !catalogue[code]) {
        problems.push(locale + " is missing " + code);
      } else if (placeholders(catalogue[code]) !== placeholders(base[code])) {
        problems.push(locale + " has different placeholders for " + code);
      }
    });
    Object.keys(catalogue).forEach(function (code) {
      if (!Object.prototype.hasOwnProperty.call(base, code)) {
        problems.push(locale + " has " + code + ", which en does not");
      }
    });
  });

  var result = { codes: codes.length, locales: Object.keys(API_CODE_CONDITIONS), problems: problems };
  console.log("debugApiCodeCatalogue|", result);
  return result;
}
