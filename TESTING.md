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
   - **Redirect URI** - `https://businesscentral.dynamics.com/OAuthLanding.htm`
3. In **Enter New Client Secret**, paste the secret value and press Tab

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

## Path 3 - OAuth Consent via Direct Page (Happy Path)

**Goal:** Complete the full consent flow via the direct consent page and confirm a token is stored.

1. Search BC for **W365 User Token Status**
2. Click **Authorise (Consent Flow)** to open the **Connect Your Email** page
3. Click **Connect my Email**
4. A sign-in popup opens automatically - sign in with the guest user's home-tenancy account (e.g. `user@theircompany.com`) and click **Accept**
5. The popup closes automatically

**Expected:**
- Page updates to show **Connected** status without any manual URL copy/paste
- **W365 User Token Status** list shows the user with status **Active** and a populated expiry timestamp

---

## Path 3b - OAuth Consent via Email Accounts Wizard

**Goal:** Complete the full consent flow via BC's native Set Up Email Account wizard.

1. Search BC for **Email Accounts** and click **New**
2. Select **Guest Email (Microsoft Graph)** from the account type list and click **Next**
3. The **Connect Your Email** page opens - click **Connect my Email**
4. Sign in with the guest user's home-tenancy account in the popup and click **Accept**
5. Click **Next** then **Finish** in the wizard

**Expected:**
- The wizard completes and the account appears in the **Email Accounts** list
- Account name shows the user's display name; email address shows their UPN
- Account type shows **Guest Email (Microsoft Graph)**

---

## Path 4 - Send Test Email via Connect Your Email Page

**Goal:** Confirm an email is sent from the guest user's home-tenancy address via the consent page.

1. On the **Connect Your Email** page (after a successful consent), scroll to the **Test Email** section
2. Enter your own email address in **Test Recipient Address**
3. Click **Send Test Email**

**Expected:**
- "Test email sent successfully via Microsoft Graph." message appears
- Email arrives in the recipient inbox
- The **From** address is the guest user's home-tenancy address (e.g. `user@theircompany.com`), not a BC host tenant address

---

## Path 4b - Send Test Email via Email Accounts Page

**Goal:** Confirm BC's native Send Test Email works through the connector.

1. Search BC for **Email Accounts**
2. Select the guest user's account row
3. Click **Send Test Email** from the action bar

**Expected:**
- BC sends a test email using the connector
- Email arrives in the recipient inbox from the guest user's home-tenancy address
- No errors in BC

---

## Path 4c - Email Connector Account Information

**Goal:** Confirm Show Account Information opens the correct page.

1. Search BC for **Email Accounts**
2. Select the guest user's account row
3. Click **Show Account Information** (or open account details)

**Expected:**
- The **Connect Your Email** page opens for the current user
- Connected status and user details are shown correctly

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
