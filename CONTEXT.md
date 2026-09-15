

## FAQ page + callback log (15 Sep 2026)

**Surface:** the offer page is opened from an **in-app banner** (not a CleverTap in-app).
The banner points at `https://vikaswiom.github.io/wiom-csp-750-pages/flow1.html?cspId=<ID>`.
⚠️ Upstream `vikashpd.github.io/csp-750-pages` has **none** of this work and Vikas cannot push
there — if a surface points at upstream, nothing below exists on it. That was the cause of
"no events": the banner was on the upstream URL.

**`faq.html`** — companion sawaal-jawaab page, reachable from a **sticky** `जरूरी सवालों के जवाब`
pill left of the language toggle on every flow page. Sticky uses `-webkit-sticky` + a JS
feature-detect that falls back to `position:fixed` (`html.nosticky`) for old WebViews.
Carries `cspId` + `lang` across; back arrow returns via `history.back()`.
One answer open at a time; the language toggle re-opens only what is currently open.

**Callback log — every "मुझे कॉल करके समझाएं" tap:**
- ONE generic event from every CTA (offer page choice box, FAQ answer CTAs, FAQ help card):
  `?flow=P750CALL&screen=<page>_<lang>&uid=<cspId>` — no per-CTA variants.
- Endpoint: `https://script.google.com/macros/s/AKfycbzpJ5R_TtfAeyjb04775eoFFwTvHza0nCoX_Oa5knETWmzZQs0-JpGQJPjIIO1dnrEn/exec`
  (Apps Script source = `callback-sheet.gs` in this repo, single `doGet`, writes a `Callbacks`
  tab: date / time IST / csp_id / page / lang / status / requests / notes, same-day dedup).
- Fired via `sheetBeacon()` on **two channels** — `fetch(no-cors)` + `new Image()` — because some
  in-app WebViews drop the Apps Script 302 for an `<img>`.
- Tapping a call CTA no longer navigates; it confirms in place
  (`रिक्वेस्ट मिल गई — टीम जल्द ही आपको कॉल करेगी।`). Only the tapped CTA is replaced.
- `beacon-test.html` = device diagnostic: open on the phone, tap, screenshot. Reports UA,
  parsed cspId, fetch/Image availability and whether each channel's request left the device.

**Apps Script deployment gotchas that cost several rounds:**
- Pasting only the part shown in chat leaves the helpers undefined (`ReferenceError: json_ /
  logCallback_ / gif_ is not defined`) — hence the single-function rewrite.
- **Save does not update the live URL.** Deploy → Manage deployments → ✏️ → *New version*.
  "New deployment" mints a DIFFERENT /exec URL (that is what finally happened here).
- The `/dev` URL runs the latest saved code for the owner only; `/exec` runs the deployed version.
  Seeing `ok` on `/dev` while `/exec` errors is the classic mismatch.
- Access must be **Anyone**; anything else redirects to accounts.google.com and silently drops rows.
- Sheets stores the date cell as a `Date`, so a `String(cell) === 'yyyy-MM-dd'` dedup never matches
  (fixed in the repo copy; re-paste when convenient).

**ES5:** all six flow pages are now transpiled (Babel, targets chrome 38). They had 11 arrow
functions + 16 const/let, which an old Android WebView rejects as a SyntaxError — that kills the
whole script, so every button goes inert while the HTML still renders. `faq.html` was already ES5.
