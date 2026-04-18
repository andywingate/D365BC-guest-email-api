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
| Token storage | `IsolatedStorage` with `DataScope::User` | Per-user, private, encrypted at rest. Tokens are read back as `Text` for HTTP headers - `SecretText` wrapping adds no meaningful protection here since tokens originate as HTTP response text. |
| Client secret storage | `IsolatedStorage` with `DataScope::Company` | Never in a table field. Admin sets it via a masked field; value is never read back to the UI. |
| Guest detection | `#EXT#` in `User."Authentication Email"` | Entra ID always injects `#EXT#` into the UPN of every B2B guest. Available without additional Graph permissions. Not used for routing (single-account model handles all users) but available for informational use. |
| HTTP client | Native AL `HttpClient` | No external dependencies. |
| Redirect URI | `https://login.microsoftonline.com/common/oauth2/nativeclient` | User pastes full redirect URL into BC to complete the exchange. Simple for PoC. |
| RestClientOAuth library | Not adopted as a dependency | Arend-Jan Kauffmann's library follows the same Auth Code + PKCE pattern and is MIT licensed, but is in-memory only. This app requires persistent tokens since the consent session and the send context are separate. Architecture patterns were informed by that work. |

---

## Azure App Registration Requirements

The app registration lives in the **host tenant**. It uses delegated permissions so each user's consent applies only to their own mailbox.

| Setting | Value |
|---|---|
| Supported account types | Accounts in any organizational directory (Multitenant) |
| Redirect URI | `https://login.microsoftonline.com/common/oauth2/nativeclient` |
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

## Phase 3 - Multi-Home-Tenancy Support *(next)*

The current `W365 Email Setup` is a singleton - one App ID and one Host Tenant ID. This covers environments where all users authenticate against a single Entra app registration.

In environments where guests come from multiple home tenancies that each require a separate app registration (e.g. due to client IT policy), Phase 3 extends the setup layer:

- Change `W365 Email Setup` from singleton to row-based, keyed by home tenant domain or tenant ID
- At token exchange time, detect the current user's home domain from their `#EXT#` UPN and look up the matching setup row
- Admin UI updated to a list + card pattern to manage multiple registrations
- IsolatedStorage key strategy for per-registration client secrets (keyed by App ID)

---

## Security

- Tokens never appear in error messages, captions, or telemetry dimensions
- Client secret stored in `IsolatedStorage` only - never read back to the UI after saving
- Redirect URI validated against the stored value on every callback - mismatch is rejected
- OAuth scope locked to `Mail.Send` only
- Token refresh is automatic and transparent to the user
- Consent is strictly per-user - tokens are `DataScope::User` and cannot be accessed by other users or background tasks running as a different identity

---

## Parking Lot

- Attachment handling - Graph `sendMail` supports base64 inline attachments up to 4MB and upload sessions for larger files; size threshold strategy deferred to a future phase
- PKCE S256 upgrade - current implementation uses plain method; S256 is more correct but requires SHA-256 byte handling in AL
- Multi-language / translation - not in scope for per-tenant extension
