---
description: "Project-specific AL settings for this BC Guest Email app. Universal AL coding standards are defined in the user-level al-coding-standards.instructions.md."
applyTo: "**/*.al"
---

# AL Project Settings - BC Guest Email (Per-User Email via Microsoft Graph)

Universal AL coding standards apply to this project. The following project-specific rules override or extend them.

## Project Purpose

Open source Business Central app that enables BC tenant guest users to send emails from their own home-tenancy email address (e.g. `user@theircompany.com`) rather than being blocked or forced to use the host tenant identity.

Admins configure a central setup. Each guest user completes a one-time OAuth consent step to grant the app a delegated Microsoft Graph `Mail.Send` permission token. After that, emails sent from BC use that user's identity automatically with no further action required.

## Project Context

- **BC Version**: BC27 (runtime 16.0, application 27.0.0.0)
- **Feature flags**: `NoImplicitWith` enabled in `app.json`
- **ID Range**: 50100-50149 - all new objects must use IDs within this range
- **Publisher**: Default Publisher
- **Object prefix**: `W365` - all custom object names, field names, and enum values must begin with `W365` (e.g., `W365 Email Setup`, `W365 User Token`)
- **Open source**: no AppSource submission planned - per-tenant extension rules apply

## Key Domain Objects

These are the expected core objects. New objects must follow this design:

| Object | Type | Purpose |
|---|---|---|
| `W365 Email Setup` | Table | Central admin configuration - Azure App Registration client ID, tenant ID, redirect URI |
| `W365 User Email Token` | Table | Per-user OAuth token store - user ID, access token, refresh token, expiry |
| `W365 Email Setup Card` | Page | Admin page to configure the Azure App Registration |
| `W365 User Token List` | Page | Admin view of all users and their token status |
| `W365 Graph Mail Mgt` | Codeunit | Microsoft Graph API calls - send email using delegated token |
| `W365 OAuth Mgt` | Codeunit | OAuth 2.0 flow - build auth URL, exchange code for token, refresh token |
| `W365 Email Subscriber` | Codeunit | Event subscriber hooking into BC email sending pipeline |

## Security Rules

These rules are mandatory given the app handles OAuth tokens and user credentials:

- **Never log or display OAuth tokens** - access tokens and refresh tokens must never appear in messages, error text, captions, or telemetry custom dimensions
- **Store tokens encrypted** - use `IsolatedStorage` with `DataScope::User` or encrypt values with `EncryptionKey` before storing in a table field; never store tokens as plain `Text` in a table
- **Never hardcode client secrets** - the Azure App Registration client secret must come from `IsolatedStorage` or an Azure Key Vault reference, never from code or setup tables in plain text
- **Validate redirect URIs** - the redirect URI used in OAuth flows must be validated against the value stored in `W365 Email Setup`; reject any mismatch
- **Scope least privilege** - the OAuth scope must be `Mail.Send` only; do not request broader Mail or Directory scopes
- **Token refresh** - always check token expiry before calling Graph; refresh automatically if expired before falling through to re-consent
- **User consent is per-user** - never share or copy a token from one user to another

## Microsoft Graph API Patterns

- Use `HttpClient` and `HttpRequestMessage` / `HttpResponseMessage` for all Graph calls - do not use third-party HTTP libraries
- Graph endpoint for sending mail: `POST https://graph.microsoft.com/v1.0/me/sendMail`
- Always set `Authorization: Bearer <accessToken>` header
- Always set `Content-Type: application/json` header
- Parse Graph error responses (`error.code`, `error.message`) and surface them via `ErrorInfo` with a Show-it action pointing to the token setup page where appropriate
- Handle `401 Unauthorized` by attempting a token refresh before surfacing an error to the user
- Handle Graph throttling (`429 Too Many Requests`) - respect the `Retry-After` header; do not retry in a tight loop

## OAuth 2.0 Flow Pattern

The one-time user consent flow uses the Authorization Code flow with PKCE:

1. Admin has pre-configured `W365 Email Setup` with the Azure App Registration details
2. When a user first sends an email (or from a dedicated action), `W365 OAuth Mgt` builds an authorization URL and opens it in a new browser tab via `Hyperlink()`
3. After consent, Azure AD redirects to the registered redirect URI with an auth code
4. A BC page (or external minimal web page) receives the code and calls back into BC via an API page or direct URL to complete the exchange
5. `W365 OAuth Mgt` exchanges the code for access + refresh tokens and stores them in `W365 User Email Token` (encrypted)
6. Subsequent email sends use the stored token; refresh automatically on expiry

## Event Integration

- Hook into BC's email sending pipeline using `OnBeforeSendEmail` or equivalent events from codeunit `Email` (base app)
- The subscriber should check whether the current user has a valid token in `W365 User Email Token`; if not, prompt for consent rather than failing silently
- Do not replace BC's native email for system/admin emails - only intercept outbound user-context emails
