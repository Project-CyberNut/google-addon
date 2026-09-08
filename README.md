# CyberNut Reporting Tool – Gmail add-on

Google Workspace add-on (Apps Script) that puts a **Report Email** card in the
Gmail sidebar. Reports are classified and forwarded through the CyberNut AWS
backends; the card copy is localized and the language choice is stored on the
user's CyberNut account.

## Files

| File | Purpose |
| --- | --- |
| `Appscript.json` | Add-on manifest: scopes, triggers, `urlFetchWhitelist`. Must be named `appsscript.json` inside the Apps Script project. |
| `code.gs` | Report flow: home card, step 1 (what did you click), step 2 (forward to IT). |
| `i18n.gs` | Locale registry, message catalogues, locale resolution, language dropdown. |
| `languageApi.gs` | Unification-platform language routes (supported / preferred language). |

## Setup

1. **Apps Script project.** Create one at script.google.com (or `clasp create`)
   and add the four files above. Rename the manifest to `appsscript.json` and
   enable *Show "appsscript.json" manifest file in editor* in Project Settings.
2. **Gmail advanced service** is declared in the manifest; confirm it is on.
3. **Google Cloud project.** Link a GCP project, configure the OAuth consent
   screen with the scopes listed in the manifest.
4. **Script Properties** (Project Settings → Script Properties):

   | Property | Required | Value |
   | --- | --- | --- |
   | `UNIFICATION_SERVICE_KEY` | yes | Shared service key for the `/public/*` language routes (same key the training portal uses as `UNIFICATION_SERVICE_KEY`). |
   | `UNIFICATION_ENV` | no | `prod` (default) or `dev`. Selects `https://{env}-{region}.cybernut.ai/api/v1`. |

   Without the key the language routes answer 401; the add-on then falls back
   to offering every shipped language and remembering the choice only on this
   Google account.
5. **Deploy as a test add-on**: create a Head deployment, install it under
   Gmail → Settings → Add-ons using the deployment ID, reload Gmail.

## Localization

The behaviour mirrors `user-portal-micro-learning-v2` (`src/i18n/config.ts`,
`src/lib/languageApi.ts`, branch `cyb4-1364`), so a user who picks Spanish in
the add-on opens the training portal in Spanish and vice versa.

**Backend routes** (unification platform, `x-service-key` auth):

```
GET   /public/accounts/supported-languages?domain=acme.org
GET   /public/users/preferred-language?domain=acme.org&email=jane@acme.org
PATCH /public/users/preferred-language?domain=acme.org&email=jane@acme.org&lang=es
```

**Which languages the dropdown offers**: the account's supported list,
narrowed to languages the add-on ships. If the list cannot be read, every
shipped language is offered. If fewer than two remain, the dropdown is hidden.

**Which language a card renders in**, first match wins:

1. Preference stored on the account (`GET preferred-language`)
2. Last choice made in this add-on (Apps Script user property)
3. Gmail's UI language, if shipped (`commonEventObject.userLocale`)
4. English

**On change**: the choice is remembered locally first, then `PATCH`ed to the
account, then the home card is redrawn. A failed PATCH still switches the UI
and shows a "could not save" toast.

**Caching** (CacheService, per user): region 6 h, supported list 6 h,
stored preference 1 h. A change made in the portal reaches the add-on within
an hour; a change made in the add-on is immediate.

### Adding a language

1. Add `{ code, label, country, rtl }` to `LOCALES` in `i18n.gs`. `label` is
   the native name and is deliberately not translated. `country` is the ISO
   3166 code of the flag to show (same choice as the portal's picker: en → US,
   es → ES, ar → EG); it is rendered as an emoji flag in the dropdown.
2. Add a full catalogue under `MESSAGES[code]` with the same keys as `en`.
   A missing key falls back to English and logs a warning.
3. Nothing else changes: the dropdown, resolution and API calls read the
   registry. The account also has to list the code in its supported languages
   for the option to appear.

Checkbox **values** in step 1 stay English (IT receives them in the report
body); only the labels are translated.
