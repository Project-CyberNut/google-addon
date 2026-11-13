# Gmail add-on: home panel footer (version and language button)

Covers the pinned footer on the home card (`homeFixedFooter()` in `code.gs`):
the add-on version and the language button ("【EN】 English ⌄") side by side
on one row, pinned to the bottom of the panel (Spam included). Clicking the
language button opens the language page (`showLanguagePicker()`), a pushed
card with the language dropdown. Picking a language runs
`onLanguageChange()`, which saves it and returns to the home card in that
language.

## How it works

A Gmail fixed footer only takes two buttons and always puts them together on
the right; CardService has no alignment setting. The version is a disabled
borderless button with a little invisible padding after its text
(`HOME_FOOTER_PADDING`, no-break spaces ~3.9px and thin spaces ~2.8px), which
moves it a bit to the left of the language button. `npm test` fails if a
message catalogue has no `HOME_FOOTER_PADDING` entry.

Known-good values (one row on the dev add-on): en 6 + 1 thin, es 5 + 1 thin,
ar 10 + 1 thin. The version then sits about 74px from the card's left edge,
next to the language button. It can't be pushed to the far left: more padding
wraps the footer (next section).

The same trick is used on the report cards' footer (`VERSION_FOOTER_PADDING`).

## Symptoms and fixes

### Version sits above the language button (footer is two rows)
- **Root cause:** too much padding in `HOME_FOOTER_PADDING`. Gmail wraps the
  two footer buttons onto separate rows when the version button gets too wide.
  v 2.5.9 (14-15 nbsp) and v 2.5.10 (17-22 nbsp, with a shorter
  "English ⌄" label) both did this; v 2.5.11 went back to the values above.
- **Fix:** take `nbsp` off that language's entry until the footer is one row.
  Don't add padding to push the version further left.
- **Verify:** push to the Cybernut Dev add-on, open the home card in that
  language, and check the version and the language button are on one row.

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

### New language, or a language label changed
- **Root cause:** the padding is tuned to each language's button label width.
- **Fix:** add the language's entry in `HOME_FOOTER_PADDING`, starting from
  the closest language with a similar label width (a wider label needs less
  padding), run `npm test`, then check on the dev add-on for one row.

### Language button does nothing on a message in Spam
- **Root cause:** in Spam, Gmail has left an invisible layer over the bottom
  of the panel that swallows clicks on the fixed footer (#40). The footer is
  pinned in Spam since CYB4-2596 (v 2.5.9).
- **Fix:** for Spam only, put the version and language button inside the card
  instead of the fixed footer: in `buildHomeCard()`, when `isSpamContext(e)`
  is true, add a last section with a `DecoratedText` (version text) whose end
  button is `homeLanguageButton(ctx)`, and skip `setFixedFooter`.
- **Verify:** open a message in Spam, click the language button, and check the
  language page opens; pick a language and check the home card comes back in
  it.
