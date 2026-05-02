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

## Phase 3 *(in progress)*

Phase 3 is split into five sprints. Each sprint is independently testable, leaves the app shippable, and can be handed to a sub-agent as a self-contained brief.

### Architecture Decisions for Phase 3

| Decision | Choice | Rationale |
|---|---|---|
| OAuth library | `Rest Client OAuth` by AJ Kauffmann (MIT) added as a hard dependency | Provides in-memory `SecretText`-only token handling, PKCE + state, automatic refresh, built-in redirect page, multi-tenant authority. Confirmed by AJ's `SecurityConsiderations.md` and `Architecture.md`: "Access and refresh tokens are held in `SecretText` variables only", "In-memory token storage: tokens tied to object lifetime; no persistent cache". Replicating this in our app would duplicate ~20 codeunits of well-tested code. Customers install two PTEs - acceptable. |
| Token persistence | None - in-memory only via the library's `Rest Client` instance lifetime | Removes IsolatedStorage as an attack vector for tokens entirely. Re-authentication is seamless via SSO when the same browser session is still signed in to Entra. |
| App Registration model | Multiple registrations stored as table rows, keyed by a short Code, with optional Domain Filter for current-user routing | Lets a partner-hosted environment serve users from many home tenancies, each with its own Entra app registration. One registration can also be tagged as default/fallback. |
| App Registration UI | Reusable `App Registrations` list + card pages, drilled into from any Email Account that uses Microsoft Graph | Same management surface for the Current User connector and the Shared Mailbox connector (and any future Graph-based connector). |
| Setup entry point | All setup happens through the standard `Email Account` page (drill-in actions on the account card) - no separate top-level setup page | Aligns with how BC's other email connectors expose their settings. Admins manage everything from one familiar place. |
| Connector types | Two connectors share the same OAuth + App Registration plumbing: `Current User Email API` and `Shared Mailbox Email API` | Same Graph stack, two different `/sendMail` endpoints (`/me/sendMail` vs `/users/{mailbox}/sendMail`). |
| Caption convention | All page, table, field, action and enum captions use plain English - no `W365` prefix anywhere user-visible | The `W365` prefix is retained as the AL object name prefix and the `Wingate365` publisher name is enough branding. |

### Dependency Details

Add to `app.json`:

```json
"dependencies": [
    {
        "id": "19642efb-0a6e-4738-afc4-025f37856f4f",
        "publisher": "AJ Kauffmann",
        "name": "Rest Client OAuth",
        "version": "1.0.0.0"
    }
]
```

Library object range: 50300 - 50349. Our object range: 50100 - 50149. No conflict.

Library codeunit names used (all suffixed `KFM`): `OAuth Client Application KFM`, `Microsoft Entra ID KFM`, `Auth. Code Grant Flow KFM`, `Http Authentication OAuth2 KFM`, plus interfaces `OAuth Authority KFM` and `OAuth Authorization Flow KFM`.

The library targets platform 26.0; we run on platform 27, so we satisfy the dependency. The library uses BC's System Application `Rest Client` module under the hood.

---

### Sprint 3.1 - Caption Cleanup and Library Dependency *(small, low risk)*

Goal: prepare the codebase for the architectural rework. No behaviour change.

**Tasks:**
1. Strip `W365` from every `Caption =` value across all AL objects (pages, tables, fields, actions, enums, permission sets). Object names stay unchanged - only captions change.
2. Add the `Rest Client OAuth` dependency to `app.json`.
3. Copy the AJ library `.app` file into `.alpackages/` so the workspace compiles.
4. Bump app version to `2.0.0.0` (breaking change in setup data model coming in Sprint 3.2).
5. Verify the app still compiles, deploys, and sends an email end-to-end with the existing Phase 1 stack.

**Deliverable:** Phase 1 functionality unchanged, captions clean, library available for next sprint.

---

### Sprint 3.2 - App Registration Data Model *(foundational)*

Goal: replace the singleton `W365 Email Setup` with a multi-row `App Registration` table, with a reusable list and card page set.

**New AL objects:**

| ID | Object | Type | Purpose |
|---|---|---|---|
| 50111 | `W365 App Registration` | Table | One row per Entra app registration. Fields: Code (PK), Description, App ID (Guid), Tenant ID (Text - blank or `common` = multi-tenant), Domain Filter (Text - optional, e.g. `contoso.com`), Is Default (Boolean, only one) |
| 50112 | `W365 App Registrations` | Page (List) | Reusable list - launched by drill-in from any Email Account that uses Graph |
| 50113 | `W365 App Registration Card` | Page (Card) | Edit a single registration, set the client secret (masked, write-only) |

**Removed AL objects:**

- `W365 Email Setup` table (50100) and `W365 Email Setup Card` page (50103) - replaced by the App Registration table and pages.

**IsolatedStorage key strategy:**

- Per-registration client secret keyed as `W365_CS_{AppId}` (`DataScope::Company`, `SetEncrypted()`).
- `Set Client Secret` action on the card page writes the value; the field is never read back to the UI.

**Migration:** Phase 1 data is wiped on upgrade (early Phase, not yet in production). No upgrade codeunit needed.

**Deliverable:** Admins can create, edit, and delete multiple App Registrations from a list page. No connector wiring changes yet.

---

### Sprint 3.3 - Current User Connector Rewrite Using AJ's Library *(major architectural change)*

Goal: replace the manual PKCE / control add-in / IsolatedStorage token model with AJ's `Rest Client OAuth` stack.

**Setup pattern at send time** (per AJ's `GettingStarted.md` step 2-5, with our extras):

```al
// 1. Detect the user's home domain from their Authentication Email
//    e.g. user@contoso.com -> 'contoso.com'
//    Guest user format alice_contoso.com#EXT#@hosttenant.onmicrosoft.com -> 'contoso.com'
HomeDomain := DetectHomeDomain(UserId());

// 2. Look up the App Registration for that domain (with fallback to Is Default = true)
AppReg.SetRange("Domain Filter", HomeDomain);
if not AppReg.FindFirst() then begin
    AppReg.SetRange("Domain Filter");
    AppReg.SetRange("Is Default", true);
    AppReg.FindFirst();
end;

// 3. Build OAuth Client Application from the registration
OAuthClientApplication.SetClientId(Format(AppReg."App ID"));
OAuthClientApplication.SetClientSecret(GetClientSecret(AppReg."App ID"));
OAuthClientApplication.AddScope('https://graph.microsoft.com/Mail.Send');

// 4. Authority - use the registration's Tenant ID (or 'common')
MicrosoftEntraID.SetTenantID(GetAuthorityTenant(AppReg));

// 5. Auth Code Grant Flow - PromptInteraction::None for SSO-first
AuthCodeGrantFlow.SetAuthority(MicrosoftEntraID);
AuthCodeGrantFlow.SetPromptInteraction(Enum::"Prompt Interaction"::None);

// 6. Http Authentication + Rest Client - tokens live in this instance only
HttpAuthenticationOAuth2.Initialize(OAuthClientApplication, AuthCodeGrantFlow);
RestClient.Initialize(HttpClientHandler, HttpAuthenticationOAuth2);

// 7. Send via Graph - library handles auth header, refresh, retry
RestClient.PostAsJson('https://graph.microsoft.com/v1.0/me/sendMail', SendMailJson);
```

**Hold the Rest Client across sends in one session:**

A new `SingleInstance = true` codeunit `W365 Graph Session` (50114) holds initialized `Rest Client` instances keyed by App Registration code. First send in a session triggers SSO; subsequent sends reuse the in-memory tokens.

**New AL objects:**

| ID | Object | Type | Purpose |
|---|---|---|---|
| 50114 | `W365 Graph Session` | Codeunit (`SingleInstance = true`) | Holds initialized `Rest Client` instances per App Registration for the BC session lifetime |

**Major rewrites:**

- `W365 OAuth Mgt` (50106) - reduced to: domain detection, App Registration lookup, client secret retrieval. All PKCE, code exchange, token storage code deleted.
- `W365 Graph Mail Mgt` (50105) - swap manual `HttpClient` for the library's `Rest Client`. About 100 lines of HTTP boilerplate removed.
- `W365 Guest Email Connector` (50110) - `Send()` now triggers auth inline via `W365 Graph Session`. `RegisterAccount()` no longer opens a separate consent page (it just confirms the connector is enabled).

**Deletions:**

- `W365 OAuth Consent` page (50108) - no separate consent step
- `W365 OAuth Popup` control add-in (50100) and `W365OAuthPopup.js` - replaced by library's built-in redirect page
- `W365 Consent Status` enum (50100 EnumExt is unrelated) - no longer needed
- All token-related fields on `W365 User Email Token` (Consent Status, Token Expiry, etc.) - keep only User Name (PK) + Home Email cache

**IsolatedStorage cleanup:** delete keys `W365_AT`, `W365_RT`, `W365_EXP`, `W365_CV`, `W365_STATE` on upgrade.

**Background-context safety:** in `Send()`, if `IsBackground()` and no in-memory token, return an error that BC's Email module turns into an Outbox entry with status "Re-authentication required" rather than failing silently.

**Drill-in from Email Account:** add a `ShowAccountInformation` implementation that opens a small card showing the resolved App Registration for the current user plus an `App Registrations` action that opens the full list (Sprint 3.2 page).

**Fallback if SSO popup is blocked from `Send()`:** the test page committed to `phase-3` (`W365 Auth Test`, page 50111) validates this before we start the rewrite. If the embedded redirect popup does not fire from `Send()` context, fall back to AJ's `RestClientOAuthAdvancedRedirectURI` library (separate dependency) which uses a control add-in for true SSO.

**Deliverable:** Current User connector works end-to-end via AJ's library. Zero tokens stored in IsolatedStorage. Multi-domain routing works.

---

### Sprint 3.4 - Shared Mailbox Connector *(new feature)*

Goal: add a second Email Connector that sends from a shared mailbox via Graph, reusing the same App Registration plumbing.

**Use case:** "Sales team" shared mailbox. Anyone with delegated access (member or guest) can send from BC as the shared mailbox.

**New AL objects:**

| ID | Object | Type | Purpose |
|---|---|---|---|
| 50115 | `W365 Shared Mailbox Account` | Table | One row per configured shared mailbox. Fields: Code (PK), Display Name, Mailbox Email/UPN, App Registration Code (FK -> `W365 App Registration`), Description |
| 50116 | `W365 Shared Mailbox Accounts` | Page (List) | Internal list of configured shared mailboxes |
| 50117 | `W365 Shared Mailbox Card` | Page (Card) | Edit a single shared mailbox, drill-in to its App Registration |
| 50118 | `W365 Shared Mailbox Connector` | Codeunit (Email Connector v4) | `Send()` calls `POST /v1.0/users/{mailbox}/sendMail` |
| 50101 | `W365 Shared Mailbox Connector` | Enum Extension | Adds value to `Email Connector` enum |

**Required Graph permissions** on the chosen App Registration:

- `Mail.Send.Shared` (delegated) - permits sending as a shared mailbox the user has access to

**`GetAccounts()`:** returns one entry per row in `W365 Shared Mailbox Account`. Each entry shows the mailbox display name and email. Multiple shared mailboxes can be set up in parallel and assigned to different Email Scenarios.

**`Send()` flow:**

1. Look up the row by AccountId (which equals the row's SystemId).
2. Resolve its App Registration (Sprint 3.2 table).
3. Acquire token via the same `W365 Graph Session` codeunit (Sprint 3.3) - the in-memory pattern is shared.
4. POST to `https://graph.microsoft.com/v1.0/users/{mailboxEmail}/sendMail`.

**Drill-in:** `ShowAccountInformation` opens the Shared Mailbox Card for that account.

**Deliverable:** A second working connector that admins can register multiple instances of, each pointing at a shared mailbox.

---

### Sprint 3.5 - Email Attachments and Multi-Recipient *(send completeness)*

Goal: forward attachments and all recipient addresses from `Email Message` to Graph. Affects both connectors (Current User and Shared Mailbox).

**Attachment strategy:**

| Total Size | Method | Graph Pattern |
|---|---|---|
| Under ~3 MB | Inline base64 in `sendMail` JSON | One POST to `/sendMail` with `attachments` array |
| 3 MB - 150 MB | Upload session | Create draft (`POST /messages`), `createUploadSession`, PUT byte ranges, `POST /messages/{id}/send` |

**Tasks:**

1. Iterate `EmailMessage.GetAttachments()`, sum size, choose strategy.
2. Inline path: extend the JSON builder with an `attachments[]` array of `fileAttachment` objects.
3. Upload session path: implement the four-call flow with chunked PUTs (chunk size 4 MB).
4. Pass all `To`, `Cc`, `Bcc` recipients (currently only first `To` is used).
5. For shared mailbox: same logic but against `/users/{mailbox}/messages` instead of `/me/messages`.
6. Update Entra App Registration setup notes - `Mail.ReadWrite` (delegated) needed in addition to `Mail.Send` for the upload session path.
7. Target: comfortably support 30 MB attachments.

**Deliverable:** All BC email scenarios (statements, invoices, reports) work with attachments through both connectors.

---

### Out of Scope for Phase 3

- Reply, RetrieveEmails, MarkAsRead, GetEmailFolders (`Email Connector v4` interface stubs remain `not supported`)
- AppSource publication (per-tenant only)
- Multi-language translations

---

## Sub-Agent Hand-Off Brief Template

For each sprint above, a sub-agent brief should include:

1. **Scope** - the sprint's goal in one sentence.
2. **Reference reading** - this `PLAN.md` Phase 3 section, the relevant existing AL files, AJ's library docs (`c:\Git\RestClientOAuth\RestClientOAuth\docs\GettingStarted.md`, `Architecture.md`, `SecurityConsiderations.md`).
3. **AL object inventory** - new objects to create, existing objects to modify, objects to delete (all listed above).
4. **Acceptance test** - what to manually run in a sandbox to confirm done.
5. **Branch** - work on `phase-3`, commit with `phase-3.{sprint number}: <change>` prefix, do not merge to main until all sprints are done.

---

## Security

- Tokens are never persisted - access and refresh tokens live in memory only inside AJ's library `Rest Client` instance for the BC session
- Every refresh returns a new refresh token with a new expiry (refresh token cycling)
- Client secrets stored in `IsolatedStorage` with `SetEncrypted()`, `DataScope::Company`, keyed per App Registration
- Client secrets never written into a table field, never read back to the UI
- All token and secret parameters use `SecretText` and are marked `[NonDebuggable]`
- OAuth flow handled entirely by AJ's library - no custom JavaScript or control add-in in our code (unless the SSO test in `W365 Auth Test` shows the built-in redirect is blocked from `Send()` context, in which case we add the `RestClientOAuthAdvancedRedirectURI` dependency)
- OAuth scopes locked to `Mail.Send` (Current User), `Mail.Send.Shared` (Shared Mailbox), plus `Mail.ReadWrite` for the attachment upload session path
- Consent is strictly per-user - the library scopes tokens to the authenticated user
- Re-authentication on session expiry is seamless via SSO; if not interactive, the email is queued in the BC Outbox with a clear status

---

## Parking Lot

- Multi-language / translation - not in scope for per-tenant extension
- AppSource publication - per-tenant only
- Reply / inbox retrieval - send-only connector by design