# CyberNut Reporting Tool – Gmail add-on

Google Workspace add-on (Apps Script) that puts a **Report Email** card in the
Gmail sidebar. Reports are classified and forwarded through the CyberNut AWS
backends; the card copy is localized and the language choice is stored on the
user's CyberNut account.

## Files

| File | Purpose |
| --- | --- |
| `Appscript.json` | Add-on manifest: scopes, triggers, `urlFetchWhitelist`. Must be named `appsscript.json` inside the Apps Script project. The whitelist carries both environments' hosts so one manifest serves both projects. |
| `env.gs` | Prod and dev configuration: every URL, API Gateway id and label that differs between environments, selected by `ADDON_ENV`. |
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
   | `UNIFICATION_SERVICE_KEY` | yes | Shared service key for the `/public/*` language routes (same key the training portal uses as `UNIFICATION_SERVICE_KEY`). Dev and prod have different keys. |
   | `ADDON_ENV` | no | `prod` (default) or `dev`. Selects every environment-specific value: API Gateway ids, telemetry host, training and portal hosts, unification base URLs and the card heading. See `env.gs`. `UNIFICATION_ENV` is accepted as a legacy alias. |

   Without the key the language routes answer 401; the add-on then falls back
   to offering every shipped language and remembering the choice only on this
   Google account.
5. **Deploy as a test add-on**: create a Head deployment, install it under
   Gmail → Settings → Add-ons using the deployment ID, reload Gmail.

## Environments

Prod and dev run the same code. Everything that differs lives in `env.gs`
under `ENVIRONMENTS.prod` and `ENVIRONMENTS.dev`, and the Script Property
`ADDON_ENV` picks one. Nothing in `code.gs`, `i18n.gs` or `languageApi.gs`
names a host directly.

Per environment: card heading, telemetry host (route is
`microsoftaddinactivitynew` in both), user-region host, per-region API Gateway
ids (verify / admin / service), training portal host (used for the v2 report
redirect, the onboarding link and the body-link check), legacy report portal
host, and the unification platform base URLs.

The manifest's `urlFetchWhitelist` is the union of both environments'
prefixes, generated from `environmentFetchPrefixes()` in `env.gs`, so the same
manifest deploys to both Apps Script projects. The one manifest field the
platform cannot switch at runtime is `addOns.common.name`: the dev project
sets it to "Cybernut Dev" by hand; everything else is identical.

Run `debugEnvironment()` from the editor to log the active environment and
every URL it will fetch.

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

**Which languages the dropdown offers**: exactly the account's supported
list from `GET /public/accounts/supported-languages`, in the backend's order,
using the backend's codes. Admins add or remove languages there; nothing in
the add-on changes. Name and flag for each option are derived from the code
(`Intl.DisplayNames`, `Intl.Locale#maximize`). If fewer than two remain,
the dropdown is hidden and that one language is used. If the list cannot be
read at all (no service key, unknown domain, network), the dropdown is hidden
and the card stays in English; the backend is the only source of the offer.

**Which language a card renders in**, first match wins:

1. Preference stored on the account (`GET preferred-language`)
2. Last choice made in this add-on (Apps Script user property)
3. Gmail's UI language, if shipped (`commonEventObject.userLocale`)
4. English

**On change**: the choice is remembered locally first, then `PATCH`ed to the
account, then the home card is redrawn. A failed PATCH still switches the UI
and shows a "could not save" toast.

**No caching**: every time the add-on renders a card it calls the region
lookup, supported-languages and preferred-language routes, so a change made
by an admin or in the portal shows up the next time the add-on is opened.
The user's last explicit choice in the add-on is kept as a user property,
used only when the account has no stored preference.

### Translating the card copy for a language

A language the backend offers is selectable and saved to the account even
without a translation here; the card copy then renders in English (a warning
is logged). To translate it, add a catalogue under `MESSAGES[code]` in
`i18n.gs` with the same keys as `en`. Region tags fall back to their base
language (`pt-BR` uses `pt`). A missing key falls back to English.

If CLDR's derived name, flag or direction is not the one wanted, add an entry
to `LOCALE_OVERRIDES`, e.g. `"es": { country: "MX" }`. Run
`debugLocaleDerivation(["en", "fr", "pt-BR"])` in the Apps Script editor to
see what the runtime derives for any codes.

Checkbox **values** in step 1 stay English (IT receives them in the report
body); only the labels are translated.
