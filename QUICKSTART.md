# Quick Start and Setup Guide

This guide covers everything needed to get the W365 Guest Email API extension deployed and working in a Business Central sandbox or production environment.

## Prerequisites

- Business Central 27 (SaaS or on-prem runtime 16.0+)
- An Azure / Microsoft Entra app registration in your **host tenant** (the tenant that owns the BC environment)
- Admin consent granted to the app registration for the `Mail.Send` delegated permission
- The AL Language VS Code extension with access to download symbols for your BC environment

---

## Part 1 - Entra App Registration

If you do not already have an app registration, create one in the [Azure portal](https://portal.azure.com) under **Microsoft Entra ID > App registrations**.

### Required settings

| Setting | Value |
|---|---|
| Supported account types | **Multiple Entra ID tenants** |
| Tenant restriction | **Allow only certain tenants** - then add each guest home tenant ID (see below) |
| Platform | **Web** |
| Redirect URI | `https://businesscentral.dynamics.com/OAuthLanding.htm` |

> **Why "Allow only certain tenants"?** This is the safer option. You explicitly list the Tenant IDs of home organisations whose guest users are permitted to consent. Users from any other tenant are blocked by Entra ID even if they reach the sign-in page. If you later onboard guests from a new organisation, add their Tenant ID to the allowed list.
>
> To find a home tenant's ID: ask the guest user to go to [portal.azure.com](https://portal.azure.com) and check **Microsoft Entra ID > Overview** - the **Tenant ID** is shown there. Alternatively it appears in the guest user's Authentication Email after `#EXT#@` - for example `user_contoso.com#EXT#@yourtenant.onmicrosoft.com` belongs to a guest whose home tenant you can look up via their domain.
>
> To add allowed tenants: go to the app registration > **Authentication** > **Supported accounts** tab > **Manage allowed tenants** > add each Tenant ID.

### API permissions

Add the following delegated permission under **API permissions > Microsoft Graph**:

| Permission | Type |
|---|---|
| `Mail.Send` | Delegated |

Click **Grant admin consent** after adding it.

### Client secret

Under **Certificates & secrets**, create a new client secret. Copy the **Value** immediately - it is only shown once. Store it somewhere safe (e.g. a password manager) until you are ready to paste it into the BC setup page.

### Values you will need

From the app registration **Overview** page, note:

- **Application (Client) ID**
- **Directory (Tenant) ID** of your host tenant

---

## Part 2 - Deploy the Extension

1. Open the workspace `C:\AL\D365BC-guest-email-api` in VS Code
2. Make sure your `launch.json` points to your target sandbox (tenant `Sandbox-Andy`)
3. Press **F5** (or run **AL: Publish Without Debugging**) to compile and deploy
4. Confirm the extension appears in BC under **Extension Management**

---

## Part 3 - Admin Configuration in BC

### 3a. Open the setup page

Search BC for **W365 Email Setup** and open the card.

### 3b. Enter the Entra app details

| Field | Value |
|---|---|
| App (Client) ID | The Application ID from your app registration |
| Host Tenant ID | Your host tenant's Directory ID |
| Redirect URI | `https://businesscentral.dynamics.com/OAuthLanding.htm` |

### 3c. Store the client secret

In the **Enter New Client Secret** field, paste the client secret value and press Tab or Enter. The field is masked and the value is stored encrypted in IsolatedStorage - it cannot be read back.

The **Client Secret Status** field should change to **Configured**.

### 3d. Assign permissions to guest users

Each guest user needs the **W365 Guest Email** permission set before they can access the consent page.

1. Search BC for **Users** and open the card for each guest user
2. Go to the **User Permission Sets** section
3. Add a new line: **Permission Set** = `W365 GUEST EMAIL`
4. Repeat for each guest user

> Without this step, guest users will not see the **W365 - Authorise Email Access** page when they search BC.

---

## Part 4 - User Consent Flow (per user, one-time)

> **This step must be completed by each guest user themselves.** Each guest user should sign in to BC with their own account and follow the steps below. It cannot be done by an admin on their behalf - the consent grants a token tied to the individual's home-tenancy identity.

Each guest user completes a one-time OAuth consent step after signing in to BC. The app automatically identifies them as a guest - no admin configuration per user is required.

There are two entry points for the consent flow - both open the same **Connect Your Email** page:

### Option A - via Email Accounts (recommended)

1. Search BC for **Email Accounts** and open the page
2. Click **New** to open the **Set Up Email Account** wizard
3. Select **Guest Email (Microsoft Graph)** from the account type list and click **Next**
4. The **Connect Your Email** page opens - click **Connect my Email**
5. A sign-in popup opens automatically - sign in with the guest user's **home-tenancy work account** (e.g. `user@theircompany.com`) and click **Accept**
6. The popup closes and the page updates to show **Connected**
7. Click **Next** then **Finish** in the wizard

### Option B - direct consent page

1. Search BC for **W365 User Token Status** (or use the **User Tokens** action from the W365 Email Setup Card)
2. Click **Authorise (Consent Flow)** to open the **Connect Your Email** page
3. Click **Connect my Email**
4. A sign-in popup opens automatically - sign in with the guest user's home-tenancy account and click **Accept**
5. The popup closes and the page shows **Connected**

Tokens are refreshed automatically before they expire. Users should not need to repeat this process unless they explicitly disconnect or their refresh token is revoked.

---

## Part 5 - Verify with a Test Email

### Option A - via BC native Email (recommended)

1. Search BC for **Email Accounts** and find the guest user's account
2. Select it and click **Send Test Email** from the action bar
3. BC sends a test message using the connector - the recipient receives an email **from the guest user's home-tenancy address**

### Option B - via the Connect Your Email page

1. Open the **Connect Your Email** page (search **W365 User Token Status** > **Authorise**)
2. Scroll to **Test Email** and enter a recipient address
3. Click **Send Test Email**
4. The recipient should receive an email **from the guest user's home-tenancy address** (e.g. `user@theircompany.com`) - not from the BC host tenant

---

## Part 6 - Email Accounts Integration

The extension registers itself as a native BC email connector. Once deployed, it appears alongside Microsoft 365, SMTP, and Current User on the **Email Accounts** page.

### What this gives you out of the box

- Guest users appear as named email accounts in **Email Accounts**
- BC's **Send Test Email** works directly from the accounts page
- Email Scenarios can be assigned to the connector (e.g. route sales invoices through guest addresses)
- The **Set Up Email Account** wizard walks users through consent automatically

### Email Scenarios (optional)

1. Search BC for **Email Scenarios**
2. Assign the **Guest Email (Microsoft Graph)** account to any scenario where guest users should be the sender
3. BC will automatically use the guest user's account when sending emails for those scenarios

---

## Troubleshooting

| Symptom | Likely cause | Resolution |
|---|---|---|
| "AADSTS50011: redirect URI does not match" | The OAuthLanding.htm redirect URI is not registered on the app | In Azure Portal, go to app registration > Authentication > **+ Add a platform** > **Web** > enter `https://businesscentral.dynamics.com/OAuthLanding.htm` > Save |
| "AADSTS50194: not configured as a multi-tenant application" | App registration Supported account types is set to single tenant | In Azure Portal, go to the app registration > Authentication > Supported accounts tab > change to **Multiple Entra ID tenants**, select **Allow only certain tenants**, add the guest's home Tenant ID, and Save |
| "W365 Email Setup has not been configured" | App ID / Tenant ID fields are empty | Complete Part 3 above |
| "Client secret has not been configured" | Secret was not stored | Re-enter the secret in the Enter New Client Secret field |
| "No authorisation code found in the redirect URL" | Partial URL pasted | Copy the entire browser address bar URL after consent redirect |
| "Authorisation state mismatch" | Page was refreshed between Step 1 and Step 2 | Click Step 1 again to generate a new state and verifier, then repeat |
| "Token request failed (error: invalid_client)" | Wrong client secret or App ID | Check the Entra app registration and re-enter the correct secret |
| "Microsoft Graph rejected the authorisation token (401)" | Token expired and refresh failed | Use **Disconnect** on the consent page and re-authorise |
| Send Test Email button is greyed out | No active token for this user | Complete the consent flow (Step 1 + Step 2) first |
| Email arrives from wrong address | User is not a B2B guest (no `#EXT#` in Authentication Email) | Confirm the user was invited as an Entra B2B guest, not created as a member |
