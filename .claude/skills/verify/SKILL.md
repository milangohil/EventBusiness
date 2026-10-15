---
name: verify
description: Run the Event Business web app (apps-script/Index.html) in Chromium against the real Code.gs to verify a change end to end.
---

# Verifying the Event Business app

The app is a Google Apps Script web app: `apps-script/Index.html` (UI) calls `apps-script/Code.gs` through `google.script.run`.
Opening Index.html directly runs **demo mode**, which uses an in-page `DemoServer` instead of Code.gs, so it does not exercise the server.

## Real UI + real Code.gs (preferred)
`gas.js` loads Code.gs in a Node VM with a fake SpreadsheetApp that mimics Sheets type coercion
(ISO dates -> Date cells, numeric text -> numbers, unless the cell format is `@`) and fails any Date leaking out of an API return.
`harness.js` opens Index.html in Playwright with a `google.script.run` / `google.script.history` shim wired to that backend.

```js
const { chromium } = require('/opt/node22/lib/node_modules/playwright');
const { createBackend } = require('./.claude/skills/verify/gas');
const { openApp } = require('./.claude/skills/verify/harness');
const be = createBackend('apps-script/Code.gs');          // runs setup(); sheets in be.sheets.<Tab>._rows
const b = await chromium.launch();
const A = await openApp(b, be, 'rensi@example.com');       // each openApp = one signed-in staff member
// drive with clicks: '.tabbar [data-act=new]', '#save-btn', '[data-act=menu]', '[data-act=ks-approve]' ...
// A.errors collects page errors; A.calls lists server functions called
```
Run scripts from the repo root with `node your-script.js`. Chromium is preinstalled; do not run `playwright install`.

## Flows worth driving
- New booking via form -> check `be.sheets.Records._rows` (Event Date/Phone must stay strings).
- Two sessions: stale session books the same hall/slot -> server error toast; both edit one record -> "Someone else changed".
- Kitchen: supplier -> dish (type English, Hindi fills via LanguageApp fake) -> menu -> kitchen sheet review
  (How?, red flag blocks Mark checked, Approve confirm) -> locked view with Order links -> change guests -> Re-check banner.
- Phone width 390px: `document.documentElement.scrollWidth` should stay 390.

## Gotchas
- `S`, `go(view, id)`, `byId`, `kitchenFor` are globals in the page; handy for jumping to a screen.
- Full-page screenshots show the sticky header mid-page; that is a capture artifact.
- Demo mode only: `[data-act=demo-seed]` loads sample data.
