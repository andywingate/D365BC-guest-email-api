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
| Token storage | `IsolatedStorage` with `DataScope::User` | Per-user, private. Phase 1 stores access and refresh tokens via plain `Set()`. Phase 3 will fix: only persist the refresh token, use `SetEncrypted()` for client secret (under 215-char limit), use `SecretText` parameters throughout, and stop persisting short-lived access tokens. |
| Client secret storage | `IsolatedStorage` with `DataScope::Company` | Never in a table field. Admin sets it via a masked field; value is never read back to the UI. Phase 3 will switch to `SetEncrypted()`. |
| Guest detection | `#EXT#` in `User."Authentication Email"` | Entra ID always injects `#EXT#` into the UPN of every B2B guest. Available without additional Graph permissions. Not used for routing (single-account model handles all users) but available for informational use. |
| HTTP client | Native AL `HttpClient` | No external dependencies. Phase 3 will migrate to System Application `Rest Client` module to reduce boilerplate. |
| Redirect URI | `https://businesscentral.dynamics.com/OAuthLanding.htm` | BC's standard OAuth landing page. The popup control add-in opens the consent URL and detects the redirect automatically. Phase 3 will evaluate replacing with System Application `OAuth2` module. |
| RestClientOAuth library | Not adopted as a dependency | Arend-Jan Kauffmann's library follows the same Auth Code + PKCE pattern and is MIT licensed, but is in-memory only. This app requires persistent tokens since the consent session and the send context are separate. Architecture patterns were informed by that work. Phase 3 will evaluate incorporating code or adopting the System Application OAuth2 module. |

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
| 50101 | `W365 User Email Token` | Table | Per-user consent status, token expiry, home email address |
| 50103 | `W365 Email Setup Card` | Page | Admin setup card |
| 50104 | `W365 User Token List` | Page | Admin list - all users, token status, consent and clear actions |
| 50105 | `W365 Graph Mail Mgt` | Codeunit | `POST /v1.0/me/sendMail` and `GET /me` Graph calls |
| 50106 | `W365 OAuth Mgt` | Codeunit | Auth URL builder, code exchange, token refresh |
| 50108 | `W365 OAuth Consent` | Page | User consent flow - opens PKCE popup, stores token on return |
| 50109 | `W365 Guest Email` | PermissionSet | Access to all W365 objects |
| 50110 | `W365 Guest Email Connector` | Codeunit | `Email Connector`, `Email Connector v4`, `Default Email Rate Limit` implementations |
| 50100 | `W365 Guest Email Connector` | Enum Extension | Extends `Email Connector` enum |
| 50100 | `W365 OAuth Popup` | ControlAddin | Popup window for PKCE consent flow |

---

## Phase 2 - Email Scenarios Integration *(complete - delivered in Phase 1)*

The Email Scenarios problem was solved by the single fixed-account model. A single `Email Connector v4` account is registered once, set as the system default once, and BC's Email Scenarios can be mapped to it once. Every send resolves to the correct user's Graph token at runtime via `UserSecurityId()`. No per-user scenario management required.

---

## Phase 3 *(next)*

Phase 3 covers two workstreams: feedback changes from code review, and multi-home-tenancy support.

### 3a - Feedback Changes (Code Review)

Based on review by Arend-Jan Kauffmann. All items are valid and accepted.

#### Security Hardening

**SecretText and NonDebuggable** - `W365 OAuth Mgt` currently passes access tokens, refresh tokens, and client secret as plain `Text` parameters. All token/secret-handling procedures should use `SecretText` and be marked `[NonDebuggable]` to prevent debugger exposure. Both are available on BC27 (runtime 16.0).

Files affected: `W365OAuthMgt.Codeunit.al`, `W365GraphMailMgt.Codeunit.al`

**IsolatedStorage encryption** - The code uses `IsolatedStorage.Set()` (plain text, not encrypted). `SetEncrypted()` is limited to 215 characters, which rules it out for access tokens (~1200-2000 chars from Entra) and refresh tokens. Client secret is typically under 215 chars and should use `SetEncrypted()`. Fix the misleading code comments that claim `IsolatedStorage.Set()` encrypts at rest - it does not.

Files affected: `W365OAuthMgt.Codeunit.al`

**IsolatedStorage cross-app risk** - IsolatedStorage is scoped by app ID. A second PTE published with the same app GUID could read these values and steal tokens. Investigate whether this is reproducible on current BC versions. If confirmed, document as a known PTE deployment risk and consider mitigations (e.g. prefixed keys, runtime caller validation).

#### Token Storage Simplification

**Stop persisting access tokens** - Only the refresh token needs persistent storage. The access token is valid for ~60 minutes and should not be stored. The code verifier and state are already cleaned up after exchange (correct). Remove `W365_AT` and `W365_EXP` keys from IsolatedStorage. Acquire a fresh access token from the refresh token on each send operation. The access token should live only as a `SecretText` variable for the duration of a single operation.

Files affected: `W365OAuthMgt.Codeunit.al`

#### Replace HttpClient with RestClient

The System Application `Rest Client` module (available BC21+) handles headers, content types, error responses, and retry logic. The current `W365 Graph Mail Mgt` has ~100 lines of HTTP boilerplate that RestClient eliminates. Adopt RestClient for all Graph API calls.

Files affected: `W365GraphMailMgt.Codeunit.al`

#### Replace Custom Control Add-in with System Application OAuth2

The current implementation uses a custom control add-in (`W365 OAuth Popup`) with a JavaScript file to open a popup, monitor for the redirect URL, and pass the auth code back to AL. The AL side then manually builds the PKCE challenge, exchanges the code for tokens, and stores them.

BC's System Application provides the `OAuth2` codeunit (501) with procedures like `AcquireAuthorizationCodeWithCacheByTokenCache` that handle the entire OAuth popup, PKCE, and code exchange flow natively in the BC web client. Adopting this would eliminate:

- `W365 OAuth Popup` control add-in and its JavaScript file
- Manual PKCE verifier generation in `W365 OAuth Mgt`
- Manual code exchange HTTP call in `W365 OAuth Mgt`
- The `W365 OAuth Consent` page popup wiring

Investigate whether the System Application OAuth2 module supports the specific requirements (delegated `Mail.Send`, multi-tenant app registration, guest user consent). If it does, this is a significant simplification. If not (e.g. due to token cache ownership or scope limitations), incorporate patterns from AJ's RestClientOAuth library (MIT licensed) to simplify the manual flow while keeping `SecretText` and `[NonDebuggable]`.

Files affected: `W365OAuthMgt.Codeunit.al`, `W365OAuthConsent.Page.al`, `W365OAuthPopup.ControlAddin.al`, `W365OAuthPopup.js`

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
- Client secret stored in `IsolatedStorage` only - never read back to the UI after saving
- Client secret should use `SetEncrypted()` (under 215-char limit); access/refresh tokens exceed this limit and cannot use `SetEncrypted()`
- All token-handling procedures should use `SecretText` parameters and `[NonDebuggable]` attribute
- Redirect URI validated against the stored value on every callback - mismatch is rejected
- OAuth scope locked to `Mail.Send` only
- Token refresh is automatic and transparent to the user
- Consent is strictly per-user - tokens are `DataScope::User` and cannot be accessed by other users or background tasks running as a different identity
- IsolatedStorage cross-app risk: a PTE with the same app GUID could theoretically read stored values - to be investigated and mitigated

---

## Parking Lot

- Multi-language / translation - not in scope for per-tenant extension
