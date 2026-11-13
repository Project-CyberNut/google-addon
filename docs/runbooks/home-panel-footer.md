# Gmail add-on: home panel footer (version and language button)

Covers the pinned footer on the home card (`homeFixedFooter()` in `code.gs`):
the add-on version on the left and the language button ("English ⌄", the
language name only) on the right. Clicking the language button opens the language page
(`showLanguagePicker()`), a pushed card with the language dropdown. Picking a
language runs `onLanguageChange()`, which saves it and returns to the home
card in that language.

## How it works

A Gmail fixed footer only takes two buttons and always puts them together on
the right; CardService has no alignment setting. So the version is a
disabled borderless button, and invisible padding after its text
(`HOME_FOOTER_PADDING`, no-break spaces ~3.9px and thin spaces ~2.8px)
widens it leftwards until the text sits at the left edge. The padding is per
language because the language button's label width differs per language.
`npm test` fails if a message catalogue has no `HOME_FOOTER_PADDING` entry.

The same trick is used on the report cards' footer (`VERSION_FOOTER_PADDING`).

## Symptoms and fixes

### Version sits in the middle of the footer, not at the left
- **Root cause:** too little padding for that language, or a language label
  too wide to leave room for it. The label plus the padding must stay under
  ~140px or the footer wraps (next section). In v 2.5.8 the label was
  "【EN】 English ⌄" (~114px), which left only ~25px of padding, so the
  version stayed ~74px from the left; v 2.5.10 shows the name only.
- **Fix:** add one no-break space (`nbsp`) to that language's entry in
  `HOME_FOOTER_PADDING` per ~4px it should move left. Use `thin` for the last
  couple of px.
- **Verify:** push to the Cybernut Dev add-on, open the home card in that
  language, and check the version starts about 39px from the card's left
  edge (where the report cards' footer puts it).

### Version wraps onto its own line, or the footer is two rows
- **Root cause:** too much padding: the language label plus the padding is
  over ~140px, so Gmail puts the two footer buttons on separate rows. v 2.5.9
  did this ("【EN】 English ⌄" plus 15 + 1 thin).
- **Fix:** take one `nbsp` off that language's entry until it fits on one line.
- **Verify:** as above, one row with the version at the left.

### New language, or a language label changed
- **Root cause:** the padding is tuned to each language's button label width.
- **Fix:** add or adjust the language's entry in `HOME_FOOTER_PADDING`
  (start from the closest language with a similar label width), run
  `npm test`, then tune on the dev add-on as above.

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
