# Gmail add-on: report forwarding, trash and confirmation message

Covers step 2 of the report flow (`handleStep2` in `code.gs`), which runs after
the user picks what they did with the email and presses submit.

## How it works

1. `region(domain)` resolves the account's AWS region.
2. `getPluginSettings(domain, region)` calls
   `GET {unification}/accounts/settings/plugin?domain=<user's domain>`.
   The base URL is in `env.gs` under `unification` (prod: `prod-<region>.cybernut.ai/api/v1`,
   dev: `dev-<region>.cybernut.ai/api/v1`). No auth.
3. The report goes to `eventdispatcher` with `email = forwardSuspiciousEmail`.
4. The confirmation card shows `reportedEmailConfirmationMessage`, or the
   built-in localized text if that is empty.
5. If `isReportedEmailMoveToTrash === true`, the email is moved to trash.
   Nothing else decides this: the step 2 call to `admindomainsgoogle` was removed.

Every step is logged to telemetry (`microsoftaddinactivitynew`) with the
add-on version, e.g. `forward suspicious email a@x.org,b@x.org moveToTrash false`.

## Symptoms and fixes

### Reported email was (or was not) moved to trash unexpectedly
- **Root cause:** the account's `isReportedEmailMoveToTrash` setting. Only a
  JSON boolean `true` trashes; `"true"` (string), `null` or a missing field do not.
- **Fix:** change the setting for the account in the admin portal (plugin settings).
  No add-on change needed.
- **Verify:** `curl "https://prod-<region>.cybernut.ai/api/v1/accounts/settings/plugin?domain=<domain>"`
  and check `data.isReportedEmailMoveToTrash`. Then report a test email and
  look for `moveToTrash true|false` in telemetry.

### Report did not reach IT / went to the wrong people
- **Root cause:** `forwardSuspiciousEmail` on the account is wrong or empty.
  The add-on sends it as-is (comma-separated list). An empty value is not
  blocked by the add-on.
- **Fix:** update `forwardSuspiciousEmail` in the account's plugin settings.
- **Verify:** curl as above, then check the `forward suspicious email ...`
  telemetry line for the next report.

### User sees the generic error card after pressing submit
- **Root cause:** the plugin settings call failed (non-200, no `data`, or the
  URL is not in the manifest's `urlFetchWhitelist`). The report is not sent.
- **Fix:** curl the URL above. A 4xx/5xx is a backend issue (the `message_code`
  is in the Apps Script execution log under `getPluginSettings|coded failure`).
  A "URL not allowed" exception means the deployed manifest is stale: rebuild
  and push (`npm run push:prod`) so `environmentFetchPrefixes()` regenerates it.
- **Verify:** `debugEnvironment()` in the Apps Script editor lists
  `.../accounts/settings/plugin` among the fetch prefixes; a test report succeeds.

### Confirmation text is in English for a Spanish/Arabic user
- **Root cause:** `reportedEmailConfirmationMessage` is one string per account;
  the backend does not translate it (the old `getemail` route took `?lang=`).
- **Fix:** clear the custom message to fall back to the localized default, or
  ask the backend to return a per-language message.
