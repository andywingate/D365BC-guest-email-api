# PLAN.md - Critical Review

Reviewer: GitHub Copilot (Claude Opus 4.6)
Date: 2025-04-18
Scope: Full review of PLAN.md against implemented code, AL coding standards, BC28 platform capabilities, and security requirements.
BC Code Intel analysis: comprehensive + security validation run (9 files, 0 critical/high issues).

## Overall Assessment

**Verdict: APPROVE WITH CHANGES**

The plan is well-structured, security-conscious, and makes sensible architectural decisions for a Phase 1 PoC. The implemented code is clean, follows AL conventions, and the BC Code Intel analysis returned zero critical or high-severity issues. The plan can be adopted after addressing the issues below.

## Issues Requiring Changes Before Adoption

### 1. CRITICAL - UrlEncode does not encode `%` before other characters

**File:** [W365OAuthMgt.Codeunit.al](src/Codeunits/W365OAuthMgt.Codeunit.al)
**Problem:** The custom `UrlEncode` procedure encodes `%` in its case statement, but the encoding runs sequentially character-by-character, so if a `%` appears in the input it will be encoded. However, the case statement lists `%` after characters like `#`, `@`, `&`, etc. Since the procedure works character-by-character (not find-and-replace on the whole string), this is actually correct. **Correction: on closer inspection this is NOT a bug** - the char-by-char loop means each character is only processed once.

**However**, the UrlEncode procedure is missing several characters that are unsafe in OAuth parameter values:
- `!`, `*`, `'`, `(`, `)` are RFC 3986 reserved sub-delimiters and should be percent-encoded in query parameter values
- This is low-risk for Phase 1 (these characters are unlikely in client IDs, redirect URIs, or fixed scope strings), but should be noted as a Phase 2 improvement.

**Recommendation:** Add a TODO comment noting incomplete RFC 3986 coverage. For Phase 2, consider using `System.Uri` from the .NET interop layer (available from BC runtime 10+) or the `Uri` codeunit from the System Application if available.

### 2. HIGH - State validation has a bypass path ✅ FIXED

**File:** [W365OAuthMgt.Codeunit.al](src/Codeunits/W365OAuthMgt.Codeunit.al#L82-L84)
**Problem:** The CSRF state validation in `ExchangeCodeForToken` only validates when BOTH `StateParam` and `StoredState` are non-empty:

```al
if (StateParam <> '') and (StoredState <> '') and (StateParam <> StoredState) then
    Error(StateMismatchErr);
```

This means if either side is empty, validation is silently skipped. An attacker could strip the `state` parameter from the redirect URL to bypass CSRF protection entirely.

**Recommendation:** Change to fail-closed. If a state was stored (i.e. `StoredState <> ''`), then `StateParam` MUST be present and must match. If no state was stored, that is a flow error and should also fail.

```
if StoredState = '' then
    Error('Consent flow state not found. Please restart the consent flow.');
if StateParam <> StoredState then
    Error(StateMismatchErr);
```

### 3. HIGH - Plan/code mismatch on object 50108 ✅ FIXED

**Plan says:** Object 50108 is `W365 OAuth Callback` (Page API) - "Receives the auth code redirect from Azure AD, completes token exchange."
**Code says:** Object 50108 is `W365 OAuth Consent` (regular Card page) - a manual 2-step UI where the user pastes the redirect URL.

These are fundamentally different designs. The plan describes an automatic callback API page; the code implements a manual paste-the-URL flow. The code's approach is actually the pragmatic choice for Phase 1 (BC SaaS cannot host arbitrary callback endpoints), but the plan table needs updating to match reality.

**Recommendation:** Update the AL Objects table in PLAN.md to show `W365 OAuth Consent` as a Card page with its actual purpose: "User-facing consent flow page - user opens auth URL, pastes redirect URL back, exchanges code for token."

### 4. MEDIUM - Missing `W365 User Setup Ext` (50102) mentioned in plan Step 2 ✅ FIXED

**Plan says:** Step 2 includes "Build `W365 User Setup Ext` table extension."
**AL Objects table says:** No object 50102 is listed.
**Code says:** No table extension exists. Guest detection uses `#EXT#` check on the `User` table directly.

The plan contradicts itself - the "Key Technical Decisions" section explicitly says "No manual flag, no User Setup extension needed" for member vs guest routing, but Step 2 still references building it.

**Recommendation:** Remove the `W365 User Setup Ext` reference from Step 2. The `#EXT#` detection approach is correct and simpler.

### 5. MEDIUM - `app.json` runtime mismatch ✅ FIXED

**Resolution:** `app.json` is correct at BC27 (`runtime: 16.0`, `application: 27.0.0.0`) - this matches the installed alpackages. The project instructions file `.github/instructions/al-coding-standards.instructions.md` was incorrect (said BC28/runtime 17.0) and has been updated to BC27.

**app.json says:** `"runtime": "16.0"`
**Project instructions say:** BC28 (runtime 17.0)
**The `application` field says:** `"application": "27.0.0.0"` (BC27, not BC28)

Runtime 16.0 corresponds to BC27. If targeting BC28, the runtime should be `"runtime": "17.0"` and application should be `"application": "28.0.0.0"`. If targeting BC27, the project instructions file should be corrected.

**Recommendation:** Decide the actual target version and align `app.json`, project instructions, and PLAN.md. BC28 is recommended since the Email Connector interface has moved to v4 in BC28 (the v3 interface is now obsolete), which is relevant for Phase 2.

### 6. MEDIUM - No permission set defined ✅ FIXED

**Resolution:** Created `src/PermissionSets/W365GuestEmail.PermissionSet.al` (object 50109). Added to PLAN.md AL objects table.

**Problem:** The plan does not include a permission set object, and none exists in the code. Per-tenant extensions require at least one permission set for the objects they create. Without one, administrators cannot grant users access to the `W365 Email Setup` and `W365 User Email Token` tables.

**Recommendation:** Add a `W365 Guest Email` permission set (e.g. object 50109) granting:
- Table `W365 Email Setup`: Read (admin config - insert handled by OnOpenPage)
- Table `W365 User Email Token`: Read, Insert, Modify, Delete
- Page `W365 Email Setup Card`: Execute
- Page `W365 User Token List`: Execute
- Page `W365 OAuth Consent`: Execute
- Codeunit `W365 Graph Mail Mgt`: Execute
- Codeunit `W365 OAuth Mgt`: Execute
- Codeunit `W365 Email Subscriber`: Execute

### 7. MEDIUM - `SendEmail` uses plain `Error()` instead of `ErrorInfo`

**File:** [W365GraphMailMgt.Codeunit.al](src/Codeunits/W365GraphMailMgt.Codeunit.al)
**Problem:** The project coding standards say "use `ErrorInfo` with a Show-it action pointing to the token setup page where appropriate." The `NoTokenErr` and 401 error both tell users to go to the OAuth Consent page, but use plain `Error()` strings instead of `ErrorInfo` with `AddNavigationAction`.

**Recommendation:** For Phase 1 PoC this is acceptable as-is, but add a TODO to upgrade to `ErrorInfo` with `PageNo := Page::"W365 OAuth Consent"` for the no-token and 401 errors. The plan's Phase 2 notes already mention this.

### 8. LOW - JSON building uses string concatenation instead of JsonObject ✅ FIXED

**Resolution:** `BuildSendMailJson` now uses `JsonObject`/`JsonArray` natively. `EscapeJsonString` procedure removed entirely.

**File:** [W365GraphMailMgt.Codeunit.al](src/Codeunits/W365GraphMailMgt.Codeunit.al)
**Problem:** `BuildSendMailJson` manually concatenates JSON strings with a custom `EscapeJsonString` function. AL has native `JsonObject` / `JsonArray` types that handle escaping automatically and are less error-prone.

**Recommendation:** Refactor to use `JsonObject.Add()` and `JsonObject.WriteTo()`. This eliminates the custom escaper entirely and reduces injection risk from edge-case escaping bugs. This can be done in Phase 1 since the change is low-risk and improves security.

## Issues Noted but Acceptable for Phase 1

### A. PKCE uses `plain` method instead of `S256`

The plan explicitly documents this as a Phase 1 trade-off with a Phase 2 upgrade path. This is acceptable for a PoC. The code has appropriate TODO comments.

### B. No CC/BCC/attachment support

Documented as out of scope. The Graph `sendMail` payload only includes `toRecipients`. Acceptable for Phase 1.

### C. `Tenant ID` field on setup table is unused

The `W365 Email Setup` table has a `Tenant ID` field but the OAuth endpoints use `/common/` (not `/tenantId/`). This is actually correct for multi-tenant app registrations (the `/common` endpoint handles tenant routing). The field is present for documentation/admin reference. No change needed.

### D. Token stored as `Text` not `SecretText`

The plan's "Key Technical Decisions" section explains this: tokens originate as HTTP response `Text`, so `SecretText` wrapping adds minimal benefit. `IsolatedStorage` encrypts at rest regardless. Acceptable reasoning.

### E. No telemetry instrumentation

BC Code Intel flagged the lack of telemetry as a gap. For an open-source per-tenant extension, custom telemetry is not essential in Phase 1. Phase 2 could add `Session.LogMessage()` for token refresh failures and Graph errors (without logging token values).

## Plan Structural Issues

### F. Development Sequence is outdated ✅ FIXED

The "Development Sequence" section (Steps 1-7) describes building the app from scratch, but the app is already built. Steps 1-4 are complete. The plan should either be marked as partially complete or the sequence should be updated to reflect current status.

### G. Phase 2 Email Connector should target v4 interface ✅ FIXED

BC28 obsoletes `Email Connector v3` and introduces `Email Connector v4`. The plan's Phase 2 section references "BC's `Email Account` extension interface" generically. When Phase 2 work begins, it must implement the v4 connector interface (`Email Connector v4` codeunit interface), not v2 or v3.

### H. Open Questions section is stale ✅ FIXED

The first open question about `OAuthLanding.htm` has been answered by the implementation - the app uses the `nativeclient` redirect URI with manual URL paste-back. This question should be marked as resolved.

## BC Code Intel Analysis Summary

| Metric | Value |
|---|---|
| Files analyzed | 9 |
| Total issues | 4 (all low severity) |
| Critical issues | 0 |
| High issues | 0 |
| Patterns detected | Facade pattern (correct), Temporary table safety (correct), Subscriber codeunit organization (correct) |
| Security validation | Passed - no critical security findings |

The analysis confirmed the code follows good AL patterns: proper facade pattern for external API calls, correct error response handling, and appropriate codeunit organization. The "optimization opportunities" flagged (permission checks, separation of concerns) are generic recommendations and not actionable blockers.

## Final Recommendations

**Before adopting the plan, make these changes:**

1. ✅ **Fix the state validation bypass** (Issue #2) - security fix, must be done before any testing
2. ✅ **Update PLAN.md object table** to match actual code (Issue #3 - object 50108 description)
3. ✅ **Remove `W365 User Setup Ext` reference** from Step 2 (Issue #4)
4. ✅ **Align `app.json` runtime/application version** with project instructions (Issue #5)
5. ✅ **Add a permission set object** to the plan and code (Issue #6)
6. ✅ **Refactor JSON building** to use native `JsonObject` types (Issue #8)

**Can be deferred to Phase 2:**

- Issue #1 - UrlEncode RFC 3986 completeness
- Issue #7 - ErrorInfo with NavigationAction on error paths
- Issue A - PKCE S256 upgrade
- Issue E - Telemetry instrumentation
- Issue G - Email Connector v4 interface

**Housekeeping:**

- ✅ Mark Development Sequence steps 1-4 as complete (Issue F)
- ✅ Resolve the OAuthLanding.htm open question (Issue H)
- ✅ Update Phase 2 notes to specify Email Connector v4 (Issue G)
