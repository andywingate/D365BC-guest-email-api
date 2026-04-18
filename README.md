# D365BC-guest-email-api

An AL extension for Microsoft Dynamics 365 Business Central that allows Microsoft Entra guest users to send email from their home-tenancy address via the Microsoft Graph API (`mail.Send`), bypassing the limitation where the built-in "current user" email account does not work for guest/cross-tenant identities.

## Overview

Guest accounts logging into Business Central from a different Entra tenant cannot use BC's native current-user email connector. This extension implements an OAuth 2.0 Authorization Code + PKCE flow to obtain a delegated `Mail.Send` token for each user, persisted in IsolatedStorage, and routes outbound email through the Microsoft Graph API (`POST /v1.0/me/sendMail`) for users flagged as guests in User Setup.

## Architecture

- **OAuth token acquisition** — Authorization Code + PKCE via a consent page in BC
- **Token persistence** — `IsolatedStorage` (per-user scope) so tokens survive server restarts
- **Email routing** — `W365 Use Graph Email` flag on User Setup routes traffic to Graph instead of native SMTP
- **Admin setup** — Entra App ID, Tenant ID, and Redirect URI stored in a dedicated setup table; client secret stored separately in `IsolatedStorage` (company scope)

## Acknowledgements

Architecture patterns for OAuth flows in AL were informed by the excellent reference implementation by **Arend-Jan Kauffmann**:
[ajkauffmann/RestClientOAuth](https://github.com/ajkauffmann/RestClientOAuth)

> Note: RestClientOAuth was evaluated as a direct dependency but not adopted, as it stores tokens in-memory only. This extension implements its own persistent token layer using `IsolatedStorage` to support server-side email sending where the consent page session and the send context are separate.

Thanks AJ for the inspiration and for making your work publicly available.
