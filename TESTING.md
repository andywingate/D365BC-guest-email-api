# Testing Guide

Manual sandbox test paths for the W365 Guest Email API extension. Each path covers a specific scenario and includes the expected outcome.

## Pre-test Checklist

Before running any paths, confirm:

- Extension deployed to sandbox (F5 from VS Code)
- Entra app registration exists with `Mail.Send` delegated permission and admin consent granted
- A client secret value is available to paste
- You have a guest user account (home-tenancy) available to sign in with during the consent flow

---

## Path 1 - Admin Setup

**Goal:** Confirm the setup card saves app registration details and stores the client secret correctly.

1. Search BC for **W365 Email Setup** and open it
2. Enter the following:
   - **App (Client) ID** - `deda566a-3ed3-4b8e-9238-e1eb3665c3f7`
   - **Host Tenant ID** - `585f2caa-d65b-4e77-92bd-f83b9697165c`
   - **Redirect URI** - `https://login.microsoftonline.com/common/oauth2/nativeclient`
3. In **Enter New Client Secret**, paste the secret value from `.env` and press Tab

**Expected:**
- "Client secret saved." message appears
- **Client Secret Status** field shows **Configured**
- App ID, Tenant ID, and Redirect URI persist on the page

---

## Path 2 - Guest Auto-Detection

**Goal:** Confirm the app correctly identifies guest vs member accounts using the `#EXT#` check on Authentication Email - no manual flagging required.

1. Search BC for **Users** and open the card for your guest user
2. Check the **Authentication Email** field - it should contain `#EXT#` (e.g. `user_theircompany.com#EXT#@yourtenant.onmicrosoft.com`)
3. Open the card for a regular member account - their Authentication Email should have no `#EXT#`

**Expected:**
- Guest user Authentication Email contains `#EXT#`
- Member user Authentication Email does not contain `#EXT#`
- No User Setup configuration is needed - the distinction is automatic

---

## Path 3 - OAuth Consent (Happy Path)

**Goal:** Complete the full consent flow and confirm a token is stored.

1. Search BC for **W365 User Token Status**
2. Click **Authorise (Consent Flow)**
3. On **W365 - Authorise Email Access**:
   - Click **Step 1 - Open Consent Page** - browser opens to Microsoft login
   - Sign in with the guest user's home-tenancy account and click **Accept**
   - Copy the full URL from the browser address bar (starts with `https://login.microsoftonline.com/common/oauth2/nativeclient?code=...`)
   - Paste into **Paste Redirect URL Here**
   - Click **Step 2 - Exchange Code**

**Expected:**
- "Authorisation successful." message appears
- **Token Status** on the consent page shows **Active**
- **W365 User Token Status** list shows the user with status **Active** and a populated expiry timestamp

---

## Path 4 - Send Test Email (Happy Path)

**Goal:** Confirm an email is sent from the guest user's home-tenancy address via Microsoft Graph.

1. On **W365 - Authorise Email Access** (or re-open from User Token Status > Authorise)
2. Enter your own email address in **Test Recipient Address**
3. Click **Send Test Email**

**Expected:**
- "Test email sent successfully via Microsoft Graph." message appears
- Email arrives in the recipient inbox
- The **From** address is the guest user's home-tenancy address (e.g. `user@theircompany.com`), not a BC host tenant address

---

## Path 5 - Error Scenarios

**Goal:** Confirm all user-facing errors are surfaced correctly with helpful messages.

| Scenario | How to trigger | Expected error message |
|---|---|---|
| Setup not configured | Clear the App ID field, attempt consent (Step 1) | "W365 Email Setup has not been configured. Open the W365 Email Setup Card..." |
| Client secret missing | Use **Clear Client Secret** action on the Setup Card, attempt Step 2 | "Client secret has not been configured. Open the W365 Email Setup Card..." |
| Redirect URL missing `code=` param | Paste any URL without `code=` into the redirect field and Tab out | "The URL does not appear to contain an authorisation code..." |
| No recipient entered | Leave Test Recipient blank, click Send Test Email | "Please enter a recipient email address in the Test Recipient Address field." |
| No redirect URL when clicking Step 2 | Click Step 2 without pasting a URL | "Please paste the redirect URL from your browser..." |
| Token cleared | Click **Disconnect** on the consent page | Confirms prompt appears; after confirming, Token Status shows blank/none and Send Test Email is greyed out |

---

## Path 6 - Disconnect and Re-Authorise

**Goal:** Confirm a user can revoke their token and re-connect cleanly.

1. On **W365 - Authorise Email Access**, confirm Token Status is **Active**
2. Click **Disconnect** and confirm
3. **Expected:** Token Status clears; Send Test Email button is disabled
4. Complete Path 3 again (Steps 1 and 2 of the consent flow)
5. **Expected:** Token Status returns to **Active**; Send Test Email works again

---

## Path 7 - IsGuestUser Auto-Detection Logic

**Goal:** Confirm that `IsGuestUser()` in `W365 Email Subscriber` correctly identifies guest vs member accounts.

This is verified indirectly via Path 2 (visual inspection of Authentication Email) and the test email flow:

1. Complete the consent flow as a **guest** user (Path 3) and confirm Send Test Email works
2. Attempt to open the consent page as a **member** account - the token list may show them but Phase 2 will restrict routing automatically

The `#EXT#` detection requires no admin action. It is a read of `User."Authentication Email"` via `UserSecurityId()`. No User Setup record or custom flag is involved.

---

## Notes

- Automated AL test codeunits are not included in Phase 1 - the OAuth consent flow requires a real browser session that cannot be simulated in automated tests
- Phase 2 (email connector integration) will introduce testable unit logic that warrants dedicated test codeunits
- Token refresh is tested implicitly - after a successful consent, wait for the access token to expire (typically 1 hour) and confirm the next Send Test Email still succeeds without re-consenting
