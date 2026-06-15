# 15 — AppSource / Partner Center Publish Readiness

> **Status:** Readiness assessment + checklist. The project is currently sideloaded; this doc maps the gap to a Microsoft Partner Center (AppSource) listing.
> **Audience:** Any LLM/engineer driving the publish effort. Each gap has current state, what's needed, file anchors, and acceptance criteria.
> **Goal:** Publish SheetClaw to **Microsoft Partner Center → Marketplace (AppSource)** as an Office (Excel) add-in.
> **Caveat:** Microsoft's marketplace policies and validation rules change. Treat the specifics below as a starting checklist and **re-verify against the current Partner Center / "Office Add-ins validation policies" docs before submission.**

---

## TL;DR

SheetClaw is closer to publishable than the older planning docs implied. The **architecture is AppSource-compatible by design** — backend-less, BYOK, web access off by default and disclosed. The real gaps are **listing/account logistics, Office-on-the-web compatibility, manifest validation, and a licensing decision** — not a rewrite.

---

## What's already ready

| Area | Evidence | Notes |
|---|---|---|
| Backend-less / static hosting | Hosted on GitHub Pages (`https://cwtf.github.io/SheetClaw/…`), `manifest.xml` | AppSource fully supports static-hosted add-ins. No backend required. |
| Credentials encrypted at rest | `src/auth/secureStore.ts` (AES-GCM-256, non-extractable key in IndexedDB) | Not plaintext. Residual XSS-decrypt risk only (see Security below). |
| BYOK model | Settings UI, `src/store/slices/config.ts` | "Your key, your traffic" — clean privacy story for a marketplace add-in. |
| Web access off by default + disclosed | `appConfig.webAccess.provider: 'none'` default; `public/privacy.html` "Optional web search and URL fetching" | Reader-proxy and search egress are disclosed. |
| No telemetry | `public/privacy.html` "Data not collected"; no analytics SDK in tree | True; keep it true. |
| Manifest basics | `manifest.xml` — `ProviderName` (Icon Learning & Development Sdn Bhd), HTTPS icons, SupportUrl, `ReadWriteDocument` | Valid XML add-in manifest shape. Needs validation + a few additions (below). |

---

## Gaps & action items

Grouped by area. Ordered roughly by how likely each is to block certification.

### A. Office-on-the-web compatibility (most likely to surprise)

**Why it matters.** AppSource validation tests the add-in on **Excel on the web** in addition to Windows desktop. Anything assuming a local/native environment will fail there.

- **Localhost provider (Ollama) and any sidecar do not work on Excel on the web.** The web origin is `https://…`, so calls to `http://localhost:11434` are blocked (mixed content / CORS). `src/store/slices/config.ts` defaults `activeProvider: 'ollama'`. On the web this yields an immediate failure for a fresh user.
  - **Needed:** detect the host environment (`Office.context` / platform) and degrade gracefully — don't default a web user into an unreachable local provider; show a clear "this provider needs Excel desktop" message instead of a raw fetch error.
  - **Acceptance:** a clean Excel-on-the-web session loads, shows no uncaught errors, and a user with a cloud BYOK key can complete a basic read→write flow.
- **Reader-proxy / external fetches under the web origin.** Confirm `web_search`/`fetch_url` behave (or fail cleanly) on the web origin. The base-URL bug in [Doc 14](14-agent-speed-optimization.md) Task 1 should be fixed first — a broken core tool will read as a defect to a reviewer.
  - **Acceptance:** with Search enabled and a key, a search→fetch→write flow works on Excel on the web, or fails with an actionable message (never a blank/hung pane).

### B. Manifest & hosting

- **Validate the manifest** with the current tooling: `npx office-addin-manifest validate manifest.xml`.
  - **Acceptance:** validator passes with no errors.
- **Declare a minimum requirement set.** `manifest.xml` declares no `<Requirements>`. Add the minimum `ExcelApi` requirement set the code actually uses so Office only offers the add-in where supported.
  - **Acceptance:** `<Requirements>` present; add-in installs only on supporting hosts.
- **Stable production host.** `cwtf.github.io` is fine for AppSource, but the SourceLocation, privacy, and support URLs must be **live and stable at submission time** and stay up. Consider a custom domain if branding/longevity matters.
  - **Acceptance:** `SourceLocation`, `SupportUrl`, privacy URL all return 200 with valid content.
- **Version bump discipline.** `Version` is `1.0.0.0`; each resubmission needs an incremented version.

### C. Listing assets & Partner Center account

- **Partner Center account** with a verified company (the manifest already names a Sdn Bhd entity — registration/verification + payout/tax profile if monetizing).
- **Marketplace offer assets:** logo (store sizes), screenshots, short + long description, search keywords, category, supported Office products & versions, and supported markets/locales.
- **Required URLs in the listing:** privacy policy, support/help, and EULA (use Microsoft's Standard Application License Terms or supply your own).
  - **Acceptance:** offer draft in Partner Center passes the listing-completeness checks.

### D. Privacy policy & support contact

- **Retire "personal-use" wording** in `public/privacy.html` (line ~30 says "personal-use Microsoft Excel task pane add-in").
  - **Acceptance:** wording reflects a published add-in.
- **State encryption-at-rest.** The policy says credentials are stored in `localStorage` but not that they're AES-GCM encrypted (`secureStore.ts`). Add a sentence — it strengthens the privacy story.
- **Formal support channel.** Contact is currently "open a GitHub issue" (`privacy.html` bottom). AppSource expects a working support link/contact; a GitHub issues page can qualify, but a monitored support email tied to the provider entity is safer.
  - **Acceptance:** support URL/contact in manifest + listing + privacy policy are consistent and monitored.
- **Confirm egress completeness.** Verify every external destination is disclosed: configured LLM provider(s), configured search provider, and the reader proxy (`r.jina.ai`) — see `src/web/fetch.ts` and `src/web/providers/`.

### E. Security posture

- **Encryption at rest: done** (`src/auth/secureStore.ts`). Keep the threat-model comment honest.
- **Residual: XSS-can-decrypt + JWK fallback.** Same-origin XSS can call decrypt; the no-IndexedDB JWK fallback colocates key and ciphertext. Mitigate by minimizing injection surface:
  - Add a **Content-Security-Policy** to the hosted `taskpane.html` (no inline-script execution from untrusted sources, restrict connect-src to expected API hosts where feasible given BYOK custom endpoints).
  - Run a dependency audit (`npm audit`) and keep Office.js/React patched.
  - **Acceptance:** CSP present; `npm audit` shows no unaddressed high/critical; documented decision that OS-vault is out of scope for a backend-less add-in.

### F. Licensing decision (needs a human call)

- The repo ships under **PolyForm Noncommercial License 1.0.0** (`LICENSE.md`), and `README.md` states commercial use needs a separate commercial license (contact `christopher.wong@iconlearning.com.my`).
- The author/company holds the rights, so **publishing is the licensor exercising their own commercial rights** — not a conflict. But:
  - Decide the **EULA presented to AppSource users** (Microsoft Standard Terms vs custom). The public PolyForm license governs the source, not necessarily the marketplace distribution terms.
  - If listing for a fee or with commercial intent, make sure the public license posture and the marketplace terms are intentionally consistent.
  - **Acceptance:** explicit decision recorded on (a) marketplace EULA, (b) whether the public OSS license changes.

### G. AI-use disclosure & content policy

- Marketplace policies for AI-powered add-ins expect clear disclosure that the add-in sends user content to third-party AI providers. This is already true and disclosed in `privacy.html`; ensure the **listing description** also states it plainly (BYOK, data goes to the user's chosen provider, no SheetClaw server).
  - **Acceptance:** listing copy discloses AI use and the BYOK data flow.

---

## Suggested sequence

1. **Fix the core web tool bug** ([Doc 14](14-agent-speed-optimization.md) Task 1) — a broken `web_search` reads as a defect to a reviewer.
2. **Office-on-the-web pass (A)** — graceful degradation for local-only providers; verify the core flow on the web.
3. **Manifest validation + requirement set (B).**
4. **Privacy/support polish (D)** and **CSP/audit (E).**
5. **Licensing decision (F).**
6. **Build the Partner Center offer + assets (C)** and submit.

## Verification before submission

- `npx office-addin-manifest validate manifest.xml` passes.
- Manual run on **Excel on the web** and **Excel on Windows desktop**: install from a test catalog, complete a read→write flow with a cloud BYOK key, confirm no uncaught errors and that local-only providers degrade gracefully.
- Privacy + support + EULA URLs return 200.
- `npm audit` clean of unaddressed high/critical; `npm test` green.
- Re-read the current **Office Add-ins validation policies** and Partner Center submission checklist — confirm nothing in this doc has gone stale.
