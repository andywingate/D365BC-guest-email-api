# D365BC Current User Email API - Project Plan

## Problem Statement

Business Central's built-in email connectors do not work for Entra B2B guest users:

- **Current User** (built-in) - works only for accounts native to the BC host tenant. Guest users have a cross-tenant identity; BC cannot obtain a `Mail.Send` token for them.
- **Microsoft 365 / SMTP** - sends from a shared mailbox or service account, not the individual user's address.

The result: guest users in BC either cannot send email, or their email arrives from the wrong address.

The same problem applies to any user (guest or member) in a BC environment hosted by a partner or shared services organisation where users sign in from their own company accounts.

---

## Goals

- Every user - guest or member - sends email from BC using their own work address
- Admins configure once; no per-user admin action required after initial setup
- Users complete a one-time OAuth consent; all subsequent sends are automatic
- All BC send paths covered - compose dialog, customer statements, scheduled reports, background jobs, ISV extensions
- Open source, per-tenant extension (not AppSource)

---

## Decisions and Assumptions

| Decision | Choice | Rationale |
|---|---|---|
| Account model | Single fixed-GUID account (`Current User Email API`) | BC's `Email Scenario` table maps a scenario to one `Account Id`. A single account set as default means scenarios are configured once and every send resolves to the correct user's token at runtime via `UserSecurityId()`. Same pattern as BC's built-in Current User connector. |
| OAuth flow | Authorization Code + PKCE (plain method) | Best practice for delegated flows. Plain method avoids SHA-256 byte manipulation in AL. |
| Token storage (Phase 1) | `IsolatedStorage` with `DataScope::User` | Per-user, private. Phase 1 stores access and refresh tokens via plain `Set()`. Phase 3 replaces this entirely - see below. |
| Token storage (Phase 3) | In-memory only - no persistent token storage | Access and refresh tokens live in memory for the duration of the RestClient instance. When the session ends or the user logs out, tokens are gone and the user re-authenticates. This eliminates the IsolatedStorage attack vector for tokens entirely. With SSO via the built-in OAuth2 module, re-authentication should be seamless. Based on AJ Kauffmann's RestClientOAuth architecture. |
| Client secret storage | `IsolatedStorage` with `DataScope::Company`, `SetEncrypted()` | Never in a table field. Admin sets it via a masked field; value is never read back to the UI. Phase 1 uses plain `Set()`. Phase 3 switches to `SetEncrypted()` (client secret is under the 215-char limit). |
| Guest detection | `#EXT#` in `User."Authentication Email"` | Entra ID always injects `#EXT#` into the UPN of every B2B guest. Available without additional Graph permissions. Not used for routing (single-account model handles all users) but available for informational use. |
| HTTP client (Phase 1) | Native AL `HttpClient` | No external dependencies. Replaced in Phase 3. |
| HTTP client (Phase 3) | System Application `Rest Client` module | Handles headers, content types, error responses, retry logic. Eliminates ~100 lines of HTTP boilerplate. Also serves as the in-memory token holder - the RestClient instance holds the access token for its lifetime, refreshing automatically. |
| Redirect URI (Phase 1) | `https://businesscentral.dynamics.com/OAuthLanding.htm` | BC's standard OAuth landing page. Phase 1 uses a custom popup control add-in to open the consent URL and detect the redirect. Replaced in Phase 3. |
| OAuth flow (Phase 3) | System Application `OAuth2` codeunit (501) | Built-in OAuth redirect implementation. Handles the popup, PKCE, and code exchange natively. Eliminates the custom control add-in, JavaScript file, and manual PKCE code entirely. Per AJ Kauffmann's recommendation. |
| RestClientOAuth library | Architecture model for Phase 3 | Arend-Jan Kauffmann's library (MIT licensed) keeps tokens in memory only - no persistent storage. Phase 3 adopts this approach: tokens live in memory for the session, and re-authentication via SSO is seamless. The advanced OAuth app in the same repo demonstrates a self-hosted SSO landing page for true single sign-on. |

---

## Azure App Registration Requirements

The app registration lives in the **host tenant**. It uses delegated permissions so each user's consent applies only to their own mailbox.

| Setting | Value |
|---|---|
| Supported account types | Accounts in any organizational directory (Multitenant) |
| Redirect URI | `https://businesscentral.dynamics.com/OAuthLanding.htm` |
| API Permission | `Mail.Send` (delegated) - Microsoft Graph |
| Client secret | Required for token exchange - stored in `IsolatedStorage`, never in a table |

---

## Phase 1 - Core Send Capability *(complete)*

### Delivered

- Central admin setup table and card page (`W365 Email Setup`) for app registration details
- Per-user OAuth consent flow - Authorization Code + PKCE via popup control add-in
- Token storage, silent refresh, and expiry handling per user
- `POST /v1.0/me/sendMail` via Microsoft Graph with delegated token
- `GET /me` call at consent time to retrieve and store the user's real home email address
- `Email Connector v4` implementation - single fixed-GUID account (`Current User Email API`)
- Admin token status list page (`W365 User Token List`) with consent action and token clear
- Permission set (`W365 Guest Email`)
- Full documentation - README, QUICKSTART, TESTING

### AL Objects

| ID | Object | Type | Purpose |
|---|---|---|---|
| 50100 | `W365 Email Setup` | Table | App registration details - App ID, Tenant ID, Redirect URI |
| 50101 | `W365 User Email Token` | Table | Per-user home email address (Phase 3: consent status and token expiry fields removed) |
| 50103 | `W365 Email Setup Card` | Page | Admin setup card |
| 50104 | `W365 User Token List` | Page | Admin list - all users and connection status |
| 50105 | `W365 Graph Mail Mgt` | Codeunit | `POST /v1.0/me/sendMail` and `GET /me` Graph calls |
| 50106 | `W365 OAuth Mgt` | Codeunit | Auth URL builder, code exchange, token refresh (Phase 3: major rewrite to use OAuth2 module) |
| 50108 | `W365 OAuth Consent` | Page | User consent flow (Phase 3: **deleted** - auth handled inline at send time) |
| 50109 | `W365 Guest Email` | PermissionSet | Access to all W365 objects |
| 50110 | `W365 Guest Email Connector` | Codeunit | `Email Connector`, `Email Connector v4`, `Default Email Rate Limit` implementations |
| 50100 | `W365 Guest Email Connector` | Enum Extension | Extends `Email Connector` enum |
| 50100 | `W365 OAuth Popup` | ControlAddin | Popup window for PKCE consent flow (Phase 3: **deleted** - replaced by OAuth2 module) |

---

## Phase 2 - Email Scenarios Integration *(complete - delivered in Phase 1)*

The Email Scenarios problem was solved by the single fixed-account model. A single `Email Connector v4` account is registered once, set as the system default once, and BC's Email Scenarios can be mapped to it once. Every send resolves to the correct user's Graph token at runtime via `UserSecurityId()`. No per-user scenario management required.

---

## Phase 3 *(next)*

Phase 3 covers two workstreams: feedback changes from code review, and multi-home-tenancy support.

### 3a - Architecture Overhaul (Code Review)

Based on review and follow-up guidance from Arend-Jan Kauffmann. This is a significant architecture change - the token model moves from persistent storage to in-memory only, and the separate consent page is eliminated entirely.

#### Zero-Touch User Experience

The Phase 1 approach requires users to visit a separate "Connect Current User Email API" consent page before they can send email. Phase 3 removes that step entirely:

1. Admin configures the Entra app registration in W365 Email Setup and enables the connector as default
2. Users just use BC normally - compose emails, send statements, print reports
3. At first send, the OAuth2 module triggers SSO or consent inline - the user signs in, approves, and the email sends in one flow
4. Subsequent sends in the same session use the in-memory token - no prompt
5. New sessions - SSO kicks in silently, no popup, no user action

No consent page, no "connect my email" button, no extra step. Users never need to know the connector exists.

#### In-Memory Token Model (No Persistent Token Storage)

The Phase 1 approach stores access tokens, refresh tokens, and expiry timestamps in IsolatedStorage. AJ's feedback: **don't store tokens at all.**

His RestClientOAuth library keeps tokens in memory for the lifetime of the RestClient instance. When the instance is destroyed or the user logs out, tokens are gone and the user must re-authenticate. This eliminates the IsolatedStorage attack vector entirely - there is nothing for a malicious PTE with a cloned app GUID to steal.

Key points from AJ:
- Every time you acquire an access token, you also get a new refresh token that replaces the previous one with a new expiry date (refresh token cycling)
- The RestClient instance holds both tokens in memory and refreshes automatically in the background
- When the session ends, tokens are gone - the user re-authenticates
- With the built-in OAuth2 module's SSO support, re-authentication should be seamless (no popup, no user action)

Implementation:

1. Remove all IsolatedStorage keys for tokens: `W365_AT`, `W365_RT`, `W365_EXP`, `W365_CV`, `W365_STATE`
2. Only `W365_CS` (client secret, `DataScope::Company`) remains in IsolatedStorage, switched to `SetEncrypted()`
3. Token acquisition and refresh handled by RestClient + OAuth2 module at send time
4. If the user has no valid session, the OAuth2 module triggers SSO or prompts for consent inline
5. The `W365 User Email Token` table stores the user's home email address (for display in Email Accounts) after first successful auth - populated from `GET /me` at first send. No token expiry tracking.

Files affected: `W365OAuthMgt.Codeunit.al` (major rewrite), `W365GraphMailMgt.Codeunit.al`, `W365GuestEmailConnector.Codeunit.al`

#### Send Path - No Silent Failures

**Emails must never silently fail due to an expired or missing token.**

- **Interactive context** (compose dialog, manual send, print-and-email): the OAuth2 module triggers SSO or consent inline before sending. The user sees a sign-in prompt if needed, completes it, and the email sends. No error, no lost email.
- **Background context** (scheduled reports, job queue, ISV sends): if no valid in-memory token exists and the context is non-interactive, the email must be queued in BC's outbox with a clear status message ("Re-authentication required") rather than failing silently. The user picks it up on their next interactive session.

#### GetAccounts Display Email

The Email Accounts page needs to show a "From" address before the user has ever authenticated via Graph. Phase 3 approach:

- Before first send: show the user's BC `Authentication Email` from the `User` table as the display address
- After first successful send: call `GET /me` to retrieve the real Graph home email address, store it in `W365 User Email Token`, and use that going forward
- This handles the chicken-and-egg problem without requiring a separate consent step

With SSO via the built-in OAuth2 module, re-authentication should be seamless in most cases. If SSO is not viable for cross-tenant guest users, investigate AJ's advanced OAuth app pattern (self-hosted SSO landing page) as a fallback.

#### SecretText and NonDebuggable

All token and secret parameters must use `SecretText` and be marked `[NonDebuggable]`. Both are available on BC27 (runtime 16.0). With the in-memory model, this primarily affects the RestClient token handling and the client secret storage path.

Files affected: `W365OAuthMgt.Codeunit.al`, `W365GraphMailMgt.Codeunit.al`

#### Client Secret Encryption

Switch `W365_CS` from `IsolatedStorage.Set()` to `SetEncrypted()`. Client secret is under the 215-character limit. Fix the misleading code comments that claim `Set()` encrypts at rest - it does not.

Files affected: `W365OAuthMgt.Codeunit.al`

#### Replace HttpClient with RestClient

The System Application `Rest Client` module (available BC21+) handles headers, content types, error responses, and retry logic. The current `W365 Graph Mail Mgt` has ~100 lines of HTTP boilerplate that RestClient eliminates. The RestClient instance also serves as the in-memory token holder.

Files affected: `W365GraphMailMgt.Codeunit.al`

#### Replace Custom Control Add-in with System Application OAuth2

The built-in OAuth2 library has a built-in OAuth redirect implementation that eliminates the custom control add-in entirely. Per AJ's direct guidance: "look into the OAuth library at the built-in OAuth redirect implementation."

This removes:

- `W365 OAuth Popup` control add-in definition (deleted)
- `W365OAuthPopup.js` JavaScript file (deleted)
- `W365 OAuth Consent` page (deleted - no separate consent step needed)
- Manual PKCE verifier generation in `W365 OAuth Mgt`
- Manual code exchange HTTP call in `W365 OAuth Mgt`

Auth is triggered from the connector's `Send()` path via the OAuth2 codeunit. The OAuth2 module handles the popup, PKCE, code exchange, and redirect natively. Users experience it as a sign-in prompt at first send time.

Files affected: `W365OAuthMgt.Codeunit.al` (major rewrite), `W365OAuthConsent.Page.al` (deleted), `W365OAuthPopup.ControlAddin.al` (deleted), `W365OAuthPopup.js` (deleted)

### 3b - Multi-Home-Tenancy Support

The current `W365 Email Setup` is a singleton - one App ID and one Host Tenant ID. This covers environments where all users authenticate against a single Entra app registration.

In environments where guests come from multiple home tenancies that each require a separate app registration (e.g. due to client IT policy), Phase 3b extends the setup layer:

- Change `W365 Email Setup` from singleton to row-based, keyed by home tenant domain or tenant ID
- At token exchange time, detect the current user's home domain from their `#EXT#` UPN and look up the matching setup row
- Admin UI updated to a list + card pattern to manage multiple registrations
- IsolatedStorage key strategy for per-registration client secrets (keyed by App ID)

### 3c - Email Attachment Support

The current `Send()` implementation passes subject and body only - attachments from the `Email Message` codeunit are not forwarded to Graph. This is a critical gap for BC where almost every email send includes a file (invoices, statements, reports, delivery notes).

Graph supports two attachment strategies:

| Attachment Size | Method | API Pattern |
|---|---|---|
| Under ~3MB | Inline base64 in `sendMail` JSON | Single `POST /me/sendMail` with `attachments` array in the message body. Base64 encoding inflates by ~33%, so the practical file limit is ~3MB to stay under the 4MB JSON body limit. |
| 3MB - 150MB | Upload session | `POST /me/messages` (create draft) then `POST /me/messages/{id}/attachments/createUploadSession` then `PUT` byte ranges in chunks, then `POST /me/messages/{id}/send`. |

Implementation:

1. Read attachments from `Email Message` codeunit using `GetAttachments()`
2. For each attachment, check size
3. If all attachments fit inline (total base64 < ~3MB), use the current `sendMail` path with an `attachments` array added to the JSON
4. If any attachment exceeds ~3MB, switch to the draft + upload session pattern for that message
5. Target: support attachments up to 30MB comfortably (well within Graph's 150MB upload session limit)
6. Requires `Mail.ReadWrite` delegated permission in addition to `Mail.Send` for the draft/upload path - update the Entra app registration requirements and QUICKSTART accordingly

Files affected: `W365GraphMailMgt.Codeunit.al`, `W365GuestEmailConnector.Codeunit.al`

---

## Security

- Tokens never appear in error messages, captions, or telemetry dimensions
- **Tokens are never persisted** (Phase 3) - access and refresh tokens live in memory only for the duration of the RestClient instance. This eliminates the IsolatedStorage cross-app attack vector entirely.
- Client secret stored in `IsolatedStorage` with `SetEncrypted()` (Phase 3) - never in a table field, never read back to the UI after saving
- All token-handling procedures use `SecretText` parameters and `[NonDebuggable]` attribute
- OAuth flow handled by the System Application `OAuth2` module - no custom JavaScript or control add-in
- OAuth scope locked to `Mail.Send` only (plus `Mail.ReadWrite` for attachment upload sessions in Phase 3c)
- Token refresh is automatic via the RestClient instance - every refresh returns a new refresh token with a new expiry (refresh token cycling)
- Consent is strictly per-user - the OAuth2 module scopes tokens to the authenticated user
- Re-authentication on session expiry is expected to be seamless via SSO

---

## Parking Lot

- Multi-language / translation - not in scope for per-tenant extension
