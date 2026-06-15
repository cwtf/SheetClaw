# Document 5 — Auth and OAuth Flow Specification

## 5.1 Provider auth capability matrix

| Provider | OAuth/PKCE | API key | Notes |
|----------|-----------|---------|-------|
| Ollama | — | — | Local, no auth. |
| OpenAI | No **[D4 resolved]** | ✓ | `auth.openai.com` OIDC has PKCE but scopes are user-identity only — no API-access scopes, no public client for API provisioning. **API-key-only.** |
| Anthropic | — (API key) | ✓ | Direct browser calls supported via `anthropic-dangerous-direct-browser-access: true`. CORS confirmed. |
| OpenRouter | Optional (Phase 8) | ✓ | `https://openrouter.ai/api/v1`. `access-control-allow-origin: *` (CORS confirmed). OAuth/PKCE key provisioning available — primary target for §5.2. Uses Generic adapter. |
| DeepSeek | — (API key) | ✓ | `https://api.deepseek.com/v1` (OpenAI-compatible). CORS confirmed (`access-control-allow-origin: <origin>` echo). Uses Generic adapter. |
| Generic | varies | ✓ | Groq, LM Studio, self-hosted, etc. use API keys. CORS varies by endpoint. |

> **[D4 resolved]:** No usable PKCE for OpenAI API access. Ship API-key-only for all providers in Phase 1–7. §5.2 OAuth/PKCE applies to **OpenRouter only** in Phase 8.

## 5.2 OAuth 2.0 + PKCE flow (OpenRouter)

Actors: **task pane** (public client), **provider authorization server**, optional **sidecar** (loopback HTTPS) for token exchange [D1].

### Step-by-step
1. **Generate PKCE pair** in the browser: `code_verifier` = 43–128 char high-entropy random; `code_challenge` = BASE64URL(SHA-256(code_verifier)). Generate `state` (CSRF) and `nonce`. Store `{verifier, state}` transiently (sessionStorage / in-memory in AuthManager) — never persisted.
2. **Construct auth URL:**
   `GET {authorizeEndpoint}?response_type=code&client_id={CLIENT_ID}&redirect_uri={REDIRECT_URI}&scope={SCOPES}&state={state}&code_challenge={challenge}&code_challenge_method=S256`
   `REDIRECT_URI` = a same-origin page served by the add-in, e.g. `https://localhost:<port>/oauth-redirect.html`.
3. **Open the auth window** using the **Office Dialog API** — `Office.context.ui.displayDialogAsync(authUrl, {height, width, promptBeforeOpen:false})`. This is the WebView-safe way to host a popup (see §5.3). The dialog navigates the provider's login.
4. **User authenticates and consents** in the dialog. Provider redirects the dialog to `REDIRECT_URI?code=...&state=...`.
5. **Redirect page captures the code.** `oauth-redirect.html` (same origin as the dialog) reads the query string and calls `Office.context.ui.messageParent(JSON.stringify({code, state}))` to send the code back to the task pane, then closes.
6. **Task pane receives the message** via the `displayDialogAsync` callback's `DialogMessageReceived` event. AuthManager validates `state` matches the stored value → else `OAuthStateMismatch` (abort).
7. **Token exchange** — POST to `{tokenEndpoint}` with `grant_type=authorization_code, code, redirect_uri, client_id, code_verifier`.
   - **If the token endpoint sends CORS headers** → exchange directly from the browser via `fetch`.
   - **If it does not** (common) → POST to the **sidecar** loopback proxy, which forwards to the provider and returns the token JSON. **[D1]**
8. **Receive tokens** `{access_token, refresh_token?, expires_in, token_type}`. Compute `expiresAt = now + expires_in - safetyMargin(60s)`.
9. **Store tokens** under `xl.auth.<provider>` (encryption/vault per [D3], §5.6). Transition auth state → `authenticated`.
10. **Use tokens**: adapters call `AuthManager.getToken(provider)`; it returns the access token, transparently refreshing if expired.

### Refresh strategy
- `getToken()` checks `expiresAt`; if within margin, runs refresh: POST `grant_type=refresh_token` (same CORS/sidecar decision as exchange).
- On refresh success → update tokens in place. On failure → state `token-expired`, surface re-auth prompt; if refresh token invalid → `unauthenticated`.
- A background timer pre-refreshes ~1 min before expiry to avoid mid-run stalls; if a run is active, refresh is awaited inline before the next LLM call.

## 5.3 Edge WebView constraints and handling

Task panes run in **WebView2 (Edge Chromium)**. Constraints and mitigations:

1. **`window.open` popups are unreliable / may be blocked or open detached** inside the WebView host. **Mitigation:** use the **Office Dialog API** (`displayDialogAsync`) for all OAuth windows — it is the supported, host-managed popup and survives WebView restrictions.
2. **Cross-origin `postMessage` from a popup to the pane is blocked**; the Dialog API's `messageParent`/`DialogMessageReceived` channel is the sanctioned cross-window message path. The redirect page must be **same-origin** as the add-in so it can call `messageParent`.
3. **Loopback / localhost exemptions**: WebView2 must be allowed to reach `https://localhost:<port>` and `http://localhost:11434` (Ollama). On Windows, AppContainer loopback may need an exemption (`CheckNetIsolation`) — and **this can reset after Windows updates** (Risk R9). Mitigation: a documented one-line exemption command + a startup connectivity self-test that surfaces a clear fix message.
4. **Third-party-cookie / storage partitioning** can interfere with provider login sessions inside the dialog. Mitigation: rely on the auth-code redirect (not cookies in the pane); keep the dialog flow short-lived.
5. **Mixed content**: everything must be HTTPS (except the explicit Ollama loopback http allowance). The dev cert (§Doc 1) must be trusted by the machine cert store, and **trust can break after Windows updates** (Risk R8) — startup self-test detects cert/trust failure and links to the re-trust command.

## 5.4 API-key fallback flow

1. **Entry**: SettingsPanel shows a masked password input per provider. Paste key.
2. **Masking**: input type=password with a reveal toggle; stored value never rendered in full after save (show `sk-…last4`).
3. **Validation**: `testConnection(provider)` performs a minimal authenticated call (`listModels` or a 1-token completion) → `valid | invalid | network-error`. Result shown inline.
4. **Storage**: `xl.auth.<provider>` blob `{ apiKeyMasked:"sk-…1234", apiKeyCipher:<obfuscated/encrypted> , authMode:'apikey', state:'authenticated' }` ([D3]).
5. **Use**: adapters fetch the key via `AuthManager.getToken`; OpenAI/Generic → `Authorization: Bearer`, Anthropic → `x-api-key`.

## 5.5 Auth state machine (per provider)

States: `unauthenticated`, `authenticating`, `authenticated`, `token-expired`, `error`, `validating` (api-key test).

```
                 setApiKey / startOAuth
 unauthenticated ───────────────────────────► authenticating
        ▲   ▲                                      │
        │   │                          success     │  failure (denied/exchange/invalid)
        │   │                                      ▼
        │   │   signOut                       authenticated ──────────────┐
        │   └──────────────────────────────────  │  ▲                     │ token near expiry
        │                                          │  │ refresh success    ▼
        │                                          │  └──────────────  (auto-refresh)
        │  signOut / refresh-token-invalid         │
        │◄─────────────────────────────────────────┤ token expired & no refresh
        │                                          ▼
        │                                     token-expired ──re-auth──► authenticating
        │                                          │
        └──────────────── error ◄──────────────────┘  (network/exchange/refresh failure)
                            │  retry
                            └──────► authenticating
```

Transition table:

| From | Event | To |
|------|-------|----|
| unauthenticated | startOAuth | authenticating |
| unauthenticated | setApiKey | validating |
| validating | valid | authenticated |
| validating | invalid/network | error |
| authenticating | code+exchange ok | authenticated |
| authenticating | denied/state-mismatch/exchange-fail | error |
| authenticated | token near expiry | (auto-refresh) → authenticated \| token-expired |
| authenticated | signOut | unauthenticated |
| token-expired | re-auth | authenticating |
| token-expired | refresh ok | authenticated |
| error | retry | authenticating/validating |
| any | signOut | unauthenticated |

## 5.6 How AuthManager exposes state

- Holds `Record<provider, AuthState>` in the store; emits `auth:state-changed(provider, state)`.
- Selectors: `isProviderReady(provider)` (true only in `authenticated`), `activeProviderReady`.
- UI: SettingsPanel shows per-provider badge (green/amber/red); a **blocking banner** appears in ChatPanel only when the *active* provider is not `authenticated`, with a CTA to sign in / enter key.
- AgentLoop pre-flights `getToken(activeProvider)` before each run; an auth failure there short-circuits the run with an actionable error rather than a provider 401.

## 5.7 Security considerations (BYOK browser add-in; AppSource-targeted)

1. **No confidential client** — never embed a client secret; PKCE public-client only. A secret in the published bundle is readable by anyone who downloads it.
2. **Secrets at rest [D3] — resolved (implemented).** Secrets are sealed with AES-GCM-256 before they touch localStorage, using a non-extractable key in IndexedDB (`src/auth/secureStore.ts`). This defeats plaintext disk/storage scans.
   - *Residual risk:* same-origin XSS can still call decrypt; the JWK fallback path (no IndexedDB) stores key and ciphertext in the same store, which is weaker. Fully closing this requires an OS credential vault (Windows Credential Manager / DPAPI) via a sidecar — out of scope for a backend-less add-in. Tracked in [Doc 15](15-appsource-readiness.md).
   - *Posture for AppSource:* BYOK ("your key, your traffic"), encrypted at rest, disclosed in the privacy policy — acceptable for marketplace distribution.
3. **`anthropic-dangerous-direct-browser-access`** exposes the Anthropic key to the page; inherent to a BYOK browser add-in. Disclose that the key lives in the browser; the encryption-at-rest above is the mitigation.
4. **Redirect URI pinning** — only accept the exact same-origin redirect; validate `state` to prevent CSRF/code injection.
5. **Token scope minimization** — request least-privilege scopes.
6. **Loopback-only network exposure** — if a sidecar runs, bind to `127.0.0.1` only, require a per-session shared token in requests so other local processes can't drive it, and use its own trusted loopback cert.
7. **No telemetry / no third-party exfiltration** — usage data and snapshots stay local.
8. **Clear-on-signout** — wipe tokens, refresh timers, and cached models on sign-out.
