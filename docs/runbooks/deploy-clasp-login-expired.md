# Deploy fails: clasp login expired (`invalid_rapt`)

Covers the **Push** step of the `Apps Script` GitHub workflow
(`.github/workflows/deploy.yml`), which runs `clasp push --force` to send the
built add-on (`dist/dev` or `dist/prod`) to Google Apps Script.

## How it works

1. The workflow writes the `CLASPRC_JSON` repository secret to
   `~/.clasprc.json`. That file is a saved Google login (an OAuth refresh
   token) for the account that owns the Apps Script projects.
2. `clasp push` uses that login to upload the code.
3. If `DEPLOYMENT_ID_DEV` / `DEPLOYMENT_ID_PROD` is set, `clasp deploy` then
   moves the published deployment to the new code.

Nothing in the repo stores the login; it lives only in the `CLASPRC_JSON`
secret (repository level, not per environment).

## Symptom

The **Push** step fails with:

```
Run clasp push --force
{"error":"invalid_grant","error_description":"reauth related error (invalid_rapt)", ...}
Error: Process completed with exit code 1.
```

The **Check** job passes; only the push fails. Every run (dev, prod, manual)
fails the same way until the secret is refreshed. The code is not the cause.

## Root cause

The Google Workspace has **Google Cloud session control** turned on: it makes
users sign in again every N hours/days. When that window runs out, Google
rejects the saved refresh token in `CLASPRC_JSON` with `invalid_rapt`
("reauthentication required"). The secret set on 2026-11-16 expired around
2026-11-19, so the window is a few days.

## Fix

On a Mac with Node 20+, signed in to GitHub (`gh auth status`):

```sh
npm install -g @google/clasp@3
clasp logout
clasp login            # browser opens: sign in with the account that owns
                       # the add-on's Apps Script projects, allow access
gh secret set CLASPRC_JSON < ~/.clasprc.json   # run in this repo
```

Then on GitHub: Actions → the failed run → **Re-run failed jobs**.

Do not paste `~/.clasprc.json` anywhere else (chat, tickets): it is a live
login to the account.

## Verify

- `gh secret list` shows `CLASPRC_JSON` with today's date.
- The re-run's **Push** step ends without the error, and **Update deployment**
  runs (if a deployment ID is set).
- In Gmail, the add-on footer shows the version from `package.json`
  (e.g. `v 2.5.18`), and a new `HomePage` run in the project's Executions log
  shows `callErrorReportingApi|called| version: v <new version>`.

## Stop it from recurring

Refreshing the secret only lasts until the next session-control window. The
lasting fix needs a Google Workspace admin, one of:

- Admin console → Security → Access and data control → **Google Cloud session
  control**: set reauthentication to "Never require reauthentication", or
  exempt the account clasp uses (e.g. put it in an org unit / group with a
  longer or no reauth policy).
- Use a dedicated deploy account in such an org unit, and log in to clasp as
  that account.

Until one of these is done, repeat **Fix** whenever the error comes back.
