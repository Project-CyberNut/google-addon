# Gmail add-on: home panel footer (language button)

Covers the pinned footer on the home card (`homeFixedFooter()` in `code.gs`),
pinned to the bottom of the panel (Spam included). It holds two buttons: the
version on the left (a disabled, grey, borderless button) and the language
button ("【EN】 English ⌄") on the right. Clicking the language button opens
the language page (`showLanguagePicker()`), a pushed card with the language
dropdown. Picking a language runs `onLanguageChange()`, which saves it and
returns to the home card in that language.

## The version's padding (don't widen it)

Gmail keeps both footer buttons together on the right; CardService has no
alignment setting. The version is pushed left by invisible padding after it
(`HOME_FOOTER_PADDING`: no-break spaces + a thin space, per language). Gmail
wraps the footer onto two rows when the language label plus that padding
goes over ~140px, and the wrap point can only be found by trial in Gmail.
History: #53 (en 6, es 5, ar 10, + 1 thin) was one row; v 2.5.9, v 2.5.10
and v 2.5.11 raised it and wrapped (#57 reverted #56); v 2.5.12 moved the
version into the card; v 2.5.13 put it back with the #53 values.

The report cards' footer uses the same trick (`VERSION_FOOTER_PADDING`).

## Symptoms and fixes

### Footer is on two rows (version above the language button)
- **Root cause:** the version's padding plus the language label is too wide
  for one row in Gmail, for that language.
- **Fix:** take one nbsp off that language in `HOME_FOOTER_PADDING`
  (`code.gs`). Never add spaces to push the version further left.
- **Verify:** reopen the add-on in that language; the footer is one row.

### Language dropdown has only one language
- **Symptom:** the language button opens the language page, but the dropdown
  lists only the current language. (Until v 2.5.12 the button was disabled
  instead, and Gmail greyed it out, which looked broken.)
- **Root cause:** the add-on offers only one language for this user.
  `getLanguageContext()` gives one option when it can't read the user's
  email/domain, when the language preference call fails or returns nothing,
  or when the account supports one language.
- **Fix:** open the Apps Script project's **Executions** page and read:
  - `showLanguagePicker|single language|`: `hasIdentity: false` means Gmail
    gave no email/domain; `offered` is the single language.
  - `getLanguageContext|resolved|` (in the same or the latest `HomePage`
    run): `supported` is what the account's language API returned
    ("unreadable" means the call failed), `offered` is what the button can
    switch to.
  Fix the account's supported languages, or the API call, accordingly.
- **Verify:** reopen the add-on, click the button; the dropdown lists the
  account's languages.

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
