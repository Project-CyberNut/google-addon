# Gmail add-on: "network or server error" when reporting in Split Pane

Covers how the report flow (`handleStep1` → `handleStep2` in `code.gs`) finds
the email being reported. Ticket: CYB2-1457.

## Symptom
A user reports an email (often a phishing test) and, after pressing the final
**Report** button, gets an error card about a network or server error. It
happens only with Gmail's **Split Pane** on (Settings → Reading pane → "Right
of inbox" or "Below inbox"). With Split Pane off, opening the email and
reporting works.

## Root cause
Not a network or server problem. Gmail does not always resend the message
context on a button click: in Split Pane the click on step 1's Report button
can reach `handleStep2` **without `e.messageMetadata`**. Before v 2.5.8 on main (v 2.5.16 on dev)
`handleStep2` read `e.messageMetadata.messageId` with no check, so it threw
`TypeError: Cannot read properties of undefined (reading 'messageId')`.
The error classifier (`resolveApiMessageFromError` in `apiMessages.gs`)
labels any error that isn't a backend answer as `NETWORK_ERROR`, so the crash
showed as a network/server error. The report was never sent.

## Fix (v 2.5.8 on main, v 2.5.16 on dev)
- `eventMessageId(e)` reads the id from, in order: the button's action
  parameters (`e.parameters.messageId`), `e.messageMetadata.messageId`,
  `e.gmail.messageId`.
- Every report button passes the id forward (`reportAction(name, id)`): the
  home card's Report Email → `handleStep1`, and step 1's Report → `handleStep2`.
  Action parameters survive any pane mode.
- If there's still no id, `handleStep2` shows the "open the email" prompt
  instead of crashing, and logs `handleStep2 no messageId (CYB2-1457)` to
  telemetry.

If it happens again:
1. Apps Script → **Executions** → the failed `handleStep2` run. The
   `handleStep2|messageId:` line shows the id and where it came from
   (`fromParameters`, `hasMessageMetadata`).
2. `null` with `fromParameters: false` means a button that leads to
   `handleStep2` was built without `reportAction(...)`. Find it with
   `grep -n 'setFunctionName("handleStep' code.gs` (should be none).

## Verify
1. Gmail → Settings → Reading pane → "Right of inbox".
2. Select a test phishing email so it opens in the pane.
3. Open the add-on → Report Email → pick an option → Report.
4. Expect the confirmation card ("Thank you, …"), not an error, and the
   report to arrive at the forwarding address.
5. Repeat with Reading pane off (full conversation view): same result.
