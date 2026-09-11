# CyberNut Reporting Tool – Gmail add-on

Google Workspace add-on (Apps Script) that puts a **Report Email** card in the
Gmail sidebar. Reports are classified and forwarded through the CyberNut AWS
backends; the card copy is localized and the language choice is stored on the
user's CyberNut account.

## Files

| File | Purpose |
| --- | --- |
| `Appscript.json` | Manifest template: scopes, triggers, whitelist. The build writes it out as `appsscript.json` per environment with the name and whitelist filled in. If pasting by hand, rename it to `appsscript.json`. |
| `scripts/build.js`, `scripts/check.js` | Build one environment into `dist/<env>/`; pre-push checks. |
| `.github/workflows/deploy.yml` | CI: checks on pull requests; build and `clasp push` on merges to `dev` and `main`. |
| `env.gs` | Prod and dev configuration: every URL, API Gateway id and label that differs between environments, selected by `ADDON_ENV`. |
| `code.gs` | Report flow: home card, step 1 (what did you click), step 2 (forward to IT). |
| `i18n.gs` | Locale registry, message catalogues, locale resolution, language dropdown. |
| `languageApi.gs` | Unification-platform language routes (supported / preferred language). |

## Build and deploy

One code base, two Apps Script projects (dev and prod). `scripts/build.js`
assembles a deployable folder per environment and `clasp` pushes it; GitHub
Actions does both on every merge.

```bash
npm test              # syntax, catalogue parity, no hard-coded hosts, whitelist coverage
npm run build:dev     # -> dist/dev   (build.gs with BUILD_ENV="dev",  manifest named "Cybernut Dev")
npm run build:prod    # -> dist/prod  (build.gs with BUILD_ENV="prod", manifest named "CyberNut Reporting Tool")
npm run push:dev      # build + clasp push to the dev project (needs clasp login + script id)
npm run push:prod
```

The build copies every `.gs` file, writes `build.gs` (environment, version,
commit, time), and generates `appsscript.json` from `Appscript.json` with the
add-on name set and `urlFetchWhitelist` regenerated from `env.gs`. A built
project decides its environment from `build.gs`, so no `ADDON_ENV` property
is needed in either project.

**Branches -> environments**: merging to `dev` pushes to the dev project;
merging to `main` pushes to the prod project. Pull requests run the checks and
both builds without pushing. See `.github/workflows/deploy.yml`.

**One-time setup**

1. Enable the Apps Script API for the deploying Google account at
   script.google.com/home/usersettings.
2. Locally: `clasp login`, then copy `~/.clasprc.json` into the GitHub secret
   `CLASPRC_JSON` (Settings -> Secrets and variables -> Actions). Use an
   account with edit access to both Apps Script projects, ideally a service
   account-like shared account rather than a personal one.
3. Repository variables: `SCRIPT_ID_DEV` and `SCRIPT_ID_PROD` (the id in each
   project's URL, `script.google.com/.../projects/<id>/edit`). Optional:
   `DEPLOYMENT_ID_PROD` so a push to main also moves the published deployment
   to the new code; without it, pushes update the head deployment only.
4. Create two GitHub environments named `dev` and `prod` (Settings ->
   Environments). Add required reviewers to `prod` if you want a manual
   approval before production pushes.
5. In each Apps Script project, set the Script Property
   `UNIFICATION_SERVICE_KEY` once (dev and prod keys differ). It is the only
   thing the pipeline cannot set.

For local pushes, copy `clasp.targets.example.json` to `clasp.targets.json`
(git-ignored) and fill in the script ids, or export `SCRIPT_ID_DEV` /
`SCRIPT_ID_PROD`.

## Setup (manual, without the pipeline)

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

The build generates each environment's `urlFetchWhitelist` from
`environmentFetchPrefixes()` in `env.gs`, so code and manifest cannot drift,
and sets `addOns.common.name` per environment ("CyberNut Reporting Tool" /
"Cybernut Dev"). The repo's `Appscript.json` is the template.

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
