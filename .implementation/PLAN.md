# D365BC Guest Email API - Project Plan

## Problem Statement

Business Central natively supports a "Current User" email account, which sends email using the signed-in user's identity via the host tenant. This works for member accounts in that tenant.

Guest accounts (external users invited via Azure AD B2B) are signed into BC under the host tenant, but their actual email identity lives in their home tenant. When BC tries to send email "as" them using the host tenant's current user mechanism, it either fails or sends from the wrong address.

This app solves that by routing outbound email for guest users through Microsoft Graph API with a delegated `Mail.Send` token scoped to each guest user's home tenancy identity.

---

## Goals

- Guest users send email from BC appearing to come from their real home-tenancy address (e.g. `user@theircompany.com`)
- No change in experience for member accounts - they continue using BC's native current user email
- No user-visible complexity after the one-time OAuth consent step
- Open source, per-tenant extension (not AppSource)

---

## Account Type Detection - Member vs Guest

BC does not natively expose whether a user is a member or guest account via a simple flag. However, Entra ID always injects `#EXT#` into the User Principal Name of every B2B guest invite. This appears in the BC `User."Authentication Email"` field.

Example guest UPN: `user_theircompany.com#EXT#@yourtenant.onmicrosoft.com`

**Decision**: Automatic detection via `#EXT#` in Authentication Email. No manual flag, no User Setup extension, no additional Graph permissions required. This is simpler, more reliable, and requires zero admin effort per user.

---

## Phase 1 - Proof of Concept (Priority: Ship Fast)

### Scope

- Azure App Registration setup guide (README / wiki - not automated)
- Central admin setup table and page for the app registration details
- Per-user OAuth consent flow (Authorization Code + PKCE)
- Token storage (encrypted, per-user)
- Token refresh on expiry
- Sending a single email via `POST /me/sendMail` Graph endpoint
- Manual `W365 Use Graph Email` flag on User Setup
- A test action (temporary) on a page to trigger a test email send - validates the full flow

### Out of Scope for Phase 1

- Email Scenarios integration (Phase 2)
- Automatic guest detection
- Attachment handling beyond what Graph sendMail supports inline
- CC / BCC handling beyond basic pass-through

---

## Phase 2 - Email Scenarios Integration

BC's Email Scenarios feature allows routing different types of system emails (sales documents, reminders, etc.) to different email accounts. Phase 2 will:

- Register a new Email Account Connector implementing the `Email Connector v4` interface (BC27+; v3 is obsolete from BC28)
- Expose the Graph-based send as a named account type within Email Scenarios
- Allow admins to assign guest-user Graph accounts to specific scenarios
- Handle the consent / token status check within the connector flow

---

## Azure App Registration Requirements

The app registration lives in the **host tenant** (the BC tenant). It uses delegated permissions so each user's consent applies only to their own mailbox.

| Setting | Value |
|---|---|
| Supported account types | Accounts in any organizational directory (Multitenant) |
| Redirect URI | BC OAuthLanding page URL or a minimal external landing page |
| API Permission | `Mail.Send` (delegated) - Microsoft Graph |
| Client secret | Required for token exchange (stored in IsolatedStorage, never in a table) |

A setup guide (markdown) will be included in the repo covering how to register the app and where to get the client ID, tenant ID, and client secret.

---

## AL Objects - Phase 1

| ID | Object | Type | Purpose |
|---|---|---|---|
| 50100 | `W365 Email Setup` | Table | Client ID, tenant ID, redirect URI (admin config) |
| 50101 | `W365 User Email Token` | Table | Per-user access token (encrypted), refresh token (encrypted), expiry datetime, user ID |
| 50103 | `W365 Email Setup Card` | Page | Admin page - configure app registration details |
| 50104 | `W365 User Token List` | Page | Admin list - all users, token status, trigger consent action |
| 50105 | `W365 Graph Mail Mgt` | Codeunit | HttpClient calls to Graph - sendMail, parse errors |
| 50106 | `W365 OAuth Mgt` | Codeunit | Build auth URL, exchange code for token, refresh token |
| 50107 | `W365 Email Subscriber` | Codeunit | Guest detection via `#EXT#` check; Phase 2 routes email to Graph automatically |
| 50108 | `W365 OAuth Consent` | Page (Card) | User-facing consent flow - user opens auth URL in browser, pastes full redirect URL back, code is exchanged for token |
| 50109 | `W365 Guest Email` | PermissionSet | Grants users access to all W365 tables, pages, and codeunits |

---

## Development Sequence

### Step 1 - Scaffold and Cleanup *(complete)*
- Remove `HelloWorld.al`
- Update `app.json` metadata (name, publisher, description)
- Create folder structure: `src/Tables/`, `src/Pages/`, `src/Codeunits/`

### Step 2 - Data Layer *(complete)*
- Build `W365 Email Setup` table and card page
- Build `W365 User Email Token` table with encrypted field handling

### Step 3 - OAuth Flow *(complete)*
- Build `W365 OAuth Mgt` codeunit - auth URL builder and code exchange
- Build `W365 OAuth Consent` Card page - user pastes redirect URL back to complete exchange

### Step 4 - Graph Send *(complete)*
- Build `W365 Graph Mail Mgt` codeunit - send email, handle 401/429, surface errors via ErrorInfo

### Step 5 - Integration
- Build `W365 Email Subscriber` event subscriber
- Wire up flag check: if `W365 Use Graph Email` is true on User Setup, route to Graph; otherwise pass through to native BC

### Step 6 - Admin UI
- Build `W365 User Token List` page with token status column and "Trigger Consent" action
- Add test-send action (temporary, Phase 1 only)

### Step 7 - Docs and Setup Guide
- Azure App Registration step-by-step guide (SETUP.md)
- Update README with architecture summary, setup steps, and limitations

---

## Key Technical Decisions

| Decision | Choice | Rationale |
|---|---|---|
| OAuth flow | Authorization Code + PKCE (plain method for Phase 1) | Best practice for delegated flows. Plain method avoids SHA-256 byte manipulation in AL; Phase 2 upgrades to S256. |
| RestClientOAuth library | Not used as a dependency | Arend-Jan Kauffmann's library (github.com/ajkauffmann/RestClientOAuth) handles Auth Code + PKCE well and is MIT licensed. However it is in-memory only by design - tokens are lost when the AL object lifecycle resets. Our use case requires persistent tokens (email send fires in server context with no interactive prompt). We follow the same architecture patterns but implement our own persistent layer. |
| Token storage | IsolatedStorage (Text, encrypted at rest) | Tokens are read back as Text for HTTP headers. SecretText wrapping provides minimal additional protection since tokens originate as HTTP response Text. IsolatedStorage encrypts at rest regardless of overload. |
| Client secret storage | IsolatedStorage (DataScope::Company) | Never in a table field. Admin sets it via a masked field on the setup card. The value is never read back to the UI. |
| Member vs guest routing | Automatic via `#EXT#` in `User."Authentication Email"` | Entra ID always injects `#EXT#` into the UPN of every B2B guest. No manual flag, no extra Graph permissions, no User Setup extension needed. |
| HTTP client | Native AL HttpClient | No external dependencies. |
| Error surfacing | Error() with clear messages; Phase 2 adds ErrorInfo NavigationAction | Consistent with BC UX patterns. |
| Redirect URI (Phase 1) | https://login.microsoftonline.com/common/oauth2/nativeclient | User pastes full redirect URL into BC. Simple for PoC; Phase 2 adds proper callback page or control add-in. |

---

## Security Considerations

- Tokens never appear in error messages, captions, or telemetry dimensions
- Client secret stored via `IsolatedStorage` only - setup page writes it but never reads it back to the UI
- Redirect URI validated against stored value on every callback - mismatch is rejected
- Scope locked to `Mail.Send` only
- Token refresh is automatic and transparent to the user
- Consent is strictly per-user - no cross-user token sharing

---

## Phase 3 - Multi-Home-Tenancy Support

Phase 1 and 2 assume all guest users belong to a **single home tenancy** and share one Entra app registration in the host tenant.

In environments where guests come from **multiple different home tenancies** (e.g. guests from Contoso Ltd and guests from Fabrikam Ltd), the current architecture has a limitation: the single app registration and its delegated `Mail.Send` token works across tenants because the app is registered as multi-tenant. However, if stricter per-tenancy app registration isolation is required (e.g. by a client's IT policy), a separate Entra app registration per home tenancy would be needed.

Phase 3 scope:
- Extend `W365 Email Setup` to support multiple rows keyed by home tenant domain or tenant ID
- Match the current user's `#EXT#` UPN home domain to the correct app registration at token exchange time
- Admin UI to manage multiple app registrations
- Determine whether per-tenancy client secrets are feasible within IsolatedStorage constraints

---

## Open Questions / Parking Lot

- Redirect URI approach: **Resolved** - Phase 1 uses `https://login.microsoftonline.com/common/oauth2/nativeclient` with manual paste-back on the `W365 OAuth Consent` page. Phase 2 can upgrade to a control add-in for SSO-friendly UX.
- Attachment handling: Graph `sendMail` supports base64 inline attachments up to 4MB and upload sessions for larger files - decide on the size threshold approach for Phase 2
- Multi-language / translation considerations for future AppSource path (not in scope now)

---

## Phase 1 Success Criteria

- A guest user can log into BC, trigger an email send (via test action), and receive that email in their inbox from their home-tenancy address
- A member user on the same BC tenant is unaffected - their emails continue via native BC
- No tokens are visible in any UI or log
- The OAuth consent flow completes end-to-end in a sandbox environment
