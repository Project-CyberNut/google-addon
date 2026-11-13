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

The build merges every `.gs` file into one `Code.gs` (build constants first,
then `env.gs`, then the rest; Apps Script files share one global scope, so
this is behaviour-identical and easier to paste and review), and generates
`appsscript.json` from `Appscript.json` with the add-on name set and
`urlFetchWhitelist` regenerated from `env.gs`. A built project decides its
environment from the baked-in `BUILD_ENV`, so no `ADDON_ENV` property is
needed in either project. Pass `--separate` to keep one file per source.

**If you paste by hand**: build, then paste `dist/<env>/Code.gs` and
`dist/<env>/appsscript.json` into the Apps Script project. Two files, done.

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
   `DEPLOYMENT_ID_DEV` / `DEPLOYMENT_ID_PROD` so a push also moves that
   environment's published deployment to the new code; without one, the push
   updates the head deployment only (which is what a Gmail test deployment
   installed from the editor already uses).
4. Create two GitHub environments named `dev` and `prod` (Settings ->
   Environments). Add required reviewers to `prod` if you want a manual
   approval before production pushes.
5. Nothing else: the language routes need no key, and a built project
   carries its environment, so no Script Properties are required.

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
   | `ADDON_ENV` | no | `prod` (default) or `dev`. Selects every environment-specific value: API Gateway ids, telemetry host, training and portal hosts and unification base URLs. See `env.gs`. `UNIFICATION_ENV` is accepted as a legacy alias. |

5. **Deploy as a test add-on**: create a Head deployment, install it under
   Gmail → Settings → Add-ons using the deployment ID, reload Gmail.

## Environments

Prod and dev run the same code. Everything that differs lives in `env.gs`
under `ENVIRONMENTS.prod` and `ENVIRONMENTS.dev`, and the Script Property
`ADDON_ENV` picks one. Nothing in `code.gs`, `i18n.gs` or `languageApi.gs`
names a host directly.

Per environment: telemetry host (route is
`microsoftaddinactivitynew` in both), user-region host, per-region API Gateway
ids (verify / service), training portal host (used for the v2 report
redirect, the onboarding link and the body-link check), legacy report portal
host, and the unification platform base URLs.

The build generates each environment's `urlFetchWhitelist` from
`environmentFetchPrefixes()` in `env.gs`, so code and manifest cannot drift,
and sets `addOns.common.name` per environment ("CyberNut Reporting Tool" /
"Cybernut Dev"). The repo's `Appscript.json` is the template.

Run `debugEnvironment()` from the editor to log the active environment and
every URL it will fetch.

## Home panel (CYB4-2596)

The home card (`buildHomeCard()` in `code.gs`) follows the "Chip Reporting
Panel" mock attached to CYB4-2596, built only from native Gmail card
widgets (no rendered images), top to bottom:

1. Chip block as text: the blue squirrel (`CHIP_IMAGE_URL`) as a small icon
   beside a bold "Hi, I'm Chip.", then the program copy (`home.*` keys in
   `i18n.gs`).
2. "See an overview of CyberNut": a thumbnail with a play button. Clicking it
   plays the video as a GIF in the panel (see **Video** below).
3. **Report Email** (only when an email is open; filled, the mock's red
   `#C9402F`) and **How reporting works** (outlined; opens the training
   onboarding page, as "Onboarding Tutorial" did). Both are Gmail buttons.
4. Pinned footer, always at the bottom of the panel: the version (`v 2.5.12`,
   a disabled grey borderless button) beside the language button
   ("【EN】 English ⌄"). A fixed footer only takes Gmail buttons (no dropdown,
   text or image), so clicking the button (`showLanguagePicker()`) opens the
   language page: a pushed card with Gmail's back arrow and the real language
   dropdown. Choosing a language saves it (`onLanguageChange()`) and returns
   to the home card redrawn in that language. Gmail keeps both footer buttons
   together on the right; pushing the version to the left corner with padding
   wrapped the footer onto two rows (v 2.5.9 - v 2.5.11), so it stays on the
   right. With one language offered, the button is disabled. **No `Columns`:** on
   the dev add-on Gmail drew a `Columns` widget as a blank row.

**Spam.** For a message in Spam the home card has no video but the same
pinned footer. `buildHomeCard()` detects Spam itself, so a language change in
Spam keeps this layout. Caveat (#40): Gmail has left an invisible layer over
the bottom of the panel in Spam that swallowed fixed-footer clicks. If the
language button does nothing on a Spam message, that's the cause; the fix
is to move the language button back into the card for Spam (a
`DecoratedText` with the language button as its end button).

Items 1-3 are one card section, so Gmail draws no dividers between them, as
in the mock.

**Assets** live in `env.gs` as top-level constants, the same in prod and dev.
`""` hides the part, and `npm test` rejects anything that is not https:

| Constant | Empty means |
| --- | --- |
| `OVERVIEW_VIDEO_URL` | The whole video section is hidden. |
| `OVERVIEW_VIDEO_THUMBNAIL_URL` | The GIF shows straight away (or only the "Watch the full video" button, with no GIF either). |
| `OVERVIEW_VIDEO_GIF_URL` | Clicking the thumbnail opens the full video in a new tab instead of playing in the panel. |
| `CHIP_IMAGE_URL` | "Hi, I'm Chip." shows without the squirrel. |

The video, thumbnail, GIF and squirrel currently load from the `dev` branch
of this public repo (`assets/`), and the video is a temporary test video.
Before prod: confirm the final video and move the images to CyberNut's own
storage.

**Gmail limits.** Add-ons only render Google's card widgets, so these parts
of the mock are not matched:
- **Styling:** no font sizes (the 26px heading is bold body text), no
  background fills (no light blue Chip box), no full-width buttons or border
  colours. The blue header bar title is Gmail's own (it shows the add-on
  name).
- **Language dropdown in the footer:** a fixed footer only takes buttons, so
  the footer has a language button and the dropdown is on the language page.
- **Video:** it cannot play inside Gmail. Cards have no video widget or
  iframe, and Gmail opened links in a new tab even with `OpenAs.OVERLAY` (and
  with `OnClose.RELOAD`, tried in #47). So the video plays as a silent looping
  GIF in the panel:
  1. At rest the panel shows `OVERVIEW_VIDEO_THUMBNAIL_URL`, a still with a
     play button.
  2. Clicking it runs `playOverviewPreview()`, which redraws the home card
     with `OVERVIEW_VIDEO_GIF_URL` in its place (about a second, one server
     round trip). Changing language returns to the thumbnail.
  3. "Watch the full video", and clicking the GIF, open `OVERVIEW_VIDEO_URL`
     in a new tab, with sound.

  A video of 30 s or less can be shown whole this way (no sound, no pause).
  To make the GIF: up to 30 s, about 360 px wide, 8 fps, under ~1 MB (30 s of
  screen recording came to ~470 KB), the same shape as the thumbnail, and on
  a public https URL (Apps Script only receives the `.gs` files). With ffmpeg:
  `ffmpeg -i video.mp4 -t 30 -vf "fps=8,scale=360:-2:flags=lanczos,split[a][b];[a]palettegen=max_colors=96:stats_mode=diff[p];[b][p]paletteuse=dither=bayer:bayer_scale=4:diff_mode=rectangle" -loop 0 overview-preview.gif`

## Report flow: account settings (step 2)

When the user submits a report, step 2 reads the account's add-on settings
from the unification platform (public, no key):

```
GET {unification}/accounts/settings/plugin?domain=acme.org
    -> { data: { forwardSuspiciousEmail: "a@acme.org,b@acme.org",
                 isReportedEmailMoveToTrash: false,
                 reportedEmailConfirmationMessage: "Thank you, ..." } }
```

| Field | Used for |
| --- | --- |
| `forwardSuspiciousEmail` | `email` in the eventdispatcher payload (who the report is forwarded to). |
| `isReportedEmailMoveToTrash` | The **only** thing that decides whether the reported email is moved to trash. Only a literal `true` trashes. |
| `reportedEmailConfirmationMessage` | The confirmation card text. Empty -> the localized `step2.defaultConfirmation`. Not localized by the backend. |

If the call fails (non-200, no `data`), the report is **not** sent and the
error card is shown. This replaced the AWS `getemail` route and, in step 2
only, the `trashEmail` answer from `admindomainsgoogle`. Step 1 still calls
`admindomainsgoogle` to route simulation emails. Runbook:
[docs/runbooks/report-forwarding-and-trash.md](docs/runbooks/report-forwarding-and-trash.md).

## Localization

The behaviour mirrors `user-portal-micro-learning-v2` (`src/i18n/config.ts`,
`src/lib/languageApi.ts`, branch `cyb4-1364`), so a user who picks Spanish in
the add-on opens the training portal in Spanish and vice versa.

**Backend routes** (unification platform, public, no key):

```
GET   /public/users/preferences?domain=acme.org&email=jane@acme.org
      -> { data: { preferredLanguage: "es" | null, supportedLanguages: ["en","es"] } }
PATCH /public/users/preferences?domain=acme.org&email=jane@acme.org&lang=es
```

**Which languages the dropdown offers**: exactly the account's supported
list from the preferences GET, in the backend's order,
using the backend's codes. Admins add or remove languages there; nothing in
the add-on changes. Each option shows the code in capitals and the native
name derived from it (`Intl.DisplayNames`), e.g. `【ES】 Español`. If fewer than two remain,
the dropdown is hidden and that one language is used. If the list cannot be
read at all (unknown domain, network), the dropdown is hidden
and the card stays in English; the backend is the only source of the offer.

**Which language a card renders in**, first match wins:

1. Preference stored on the account (same GET)
2. Last choice made in this add-on (Apps Script user property)
3. Gmail's UI language, if shipped (`commonEventObject.userLocale`)
4. English

**On change**: the choice is remembered locally first, then `PATCH`ed to the
account, then the home card is redrawn. A failed PATCH still switches the UI
and shows a "could not save" toast.

**No caching**: every time the add-on renders a card it calls the region
lookup and the preferences GET, so a change made
by an admin or in the portal shows up the next time the add-on is opened.
The user's last explicit choice in the add-on is kept as a user property,
used only when the account has no stored preference.

### Translating the card copy for a language

A language the backend offers is selectable and saved to the account even
without a translation here; the card copy then renders in English (a warning
is logged). To translate it, add a catalogue under `MESSAGES[code]` in
`i18n.gs` with the same keys as `en`. Region tags fall back to their base
language (`pt-BR` uses `pt`). A missing key falls back to English.

If CLDR's derived name or direction is not the one wanted, add an entry
to `LOCALE_OVERRIDES`, e.g. `"pt": { label: "Português" }`. Run
`debugLocaleDerivation(["en", "fr", "pt-BR"])` in the Apps Script editor to
see what the runtime derives for any codes.

Checkbox **values** in step 1 stay English (IT receives them in the report
body); only the labels are translated.
