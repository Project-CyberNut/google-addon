# Gmail add-on: home panel footer (language button)

Covers the pinned footer on the home card (`homeFixedFooter()` in `code.gs`).
It holds one button, the language button ("【EN】 English ⌄"), pinned to the
bottom of the panel (Spam included). Clicking it opens the language page
(`showLanguagePicker()`), a pushed card with the language dropdown. Picking a
language runs `onLanguageChange()`, which saves it and returns to the home
card in that language.

The add-on version is not in the footer. It's a small grey line at the end of
the card (`buildHomeCard()`).

## Why the version left the footer

A Gmail fixed footer only takes buttons. The version used to be a second,
disabled button (Gmail greys disabled buttons), pushed left with invisible
padding. Gmail wrapped the footer onto two rows whenever that padding got
too wide, and the wrap point can only be found by trial in Gmail: v 2.5.9,
v 2.5.10 and v 2.5.11 all changed the padding, and #57 reverted #56. In
v 2.5.12 the footer holds only the language button, so there's nothing to
wrap or tune.

The report cards' footer still uses that trick (`VERSION_FOOTER_PADDING`,
measured in Gmail). Leave it alone unless it wraps.

## Symptoms and fixes

### Language button is greyed out (no ⌄) and does nothing
- **Root cause:** the add-on offers only one language for this user, so
  there's nothing to pick. `getLanguageContext()` gives one option when it
  can't read the user's email/domain, when the language preference call
  fails or returns nothing, or when the account supports one language.
- **Fix:** open the Apps Script project's **Executions** page, find the
  latest `HomePage` run and read the `getLanguageContext|resolved|` line:
  `supported` is what the account's language API returned ("unreadable"
  means the call failed), `offered` is what the button can switch to. Fix
  the account's supported languages, or the API call, accordingly.
- **Verify:** reopen the add-on; the button shows ⌄ and opens the language
  page.

### Language button does nothing on a message in Spam
- **Root cause:** in Spam, Gmail has left an invisible layer over the bottom
  of the panel that swallows clicks on the fixed footer (#40). The footer is
  pinned in Spam since CYB4-2596 (v 2.5.9).
- **Fix:** for Spam only, put the language button inside the card instead of
  the fixed footer: in `buildHomeCard()`, when `isSpamContext(e)` is true,
  add a last section with a `DecoratedText` whose end button is
  `homeLanguageButton(ctx)`, and skip `setFixedFooter`.
- **Verify:** open a message in Spam, click the language button, and check the
  language page opens; pick a language and check the home card comes back in
  it.

### Picking a language doesn't change the card
- **Root cause:** `onLanguageChange()` only accepts a code from the account's
  supported list; anything else shows "language.invalid". If the save to the
  account fails, it still switches locally and shows "savedLocallyOnly".
- **Fix:** check the toast text and the `onLanguageChange|` log line (`raw`,
  `chosen`, `offered`) in **Executions**.
- **Verify:** pick a language; the home card redraws in it with
  "Language updated.".
