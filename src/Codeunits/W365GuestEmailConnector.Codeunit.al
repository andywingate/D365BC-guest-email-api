namespace Wingate365.GuestEmailAPI;

using System.Email;
using System.Security.AccessControl;

codeunit 50110 "W365 Guest Email Connector" implements "Email Connector", "Email Connector v4", "Default Email Rate Limit"
{
    Access = Internal;

    // =========================================================================
    // "Email Connector" interface
    // =========================================================================

    /// <summary>
    /// Sends an email via Microsoft Graph using the current user's delegated token.
    /// Resolves the App Registration for the current user's home domain, acquires a token
    /// (silently via SSO first, then interactively if in an interactive session), and
    /// calls Graph with all To/Cc/Bcc recipients and attachments.
    /// In background sessions with no cached token, raises a user-friendly error so BC
    /// queues the message in the Outbox.
    /// </summary>
    procedure Send(EmailMessage: Codeunit "Email Message"; AccountId: Guid)
    var
        OAuthMgt: Codeunit "W365 OAuth Mgt";
        GraphMailMgt: Codeunit "W365 Graph Mail Mgt";
        AppReg: Record "W365 App Registration";
        AccessToken: SecretText;
        Recipients: List of [Text];
        CcRecipients: List of [Text];
        BccRecipients: List of [Text];
        NoRecipientsErr: Label 'The email message has no recipients.';
        NoTokenBackgroundErr: Label 'Email authentication is required. Sign in to Business Central interactively to renew your email authorisation, then retry.';
        NoAppRegErr: Label 'No App Registration has been configured. Create one on the App Registrations page.';
    begin
        EmailMessage.GetRecipients(Enum::"Email Recipient Type"::"To", Recipients);
        EmailMessage.GetRecipients(Enum::"Email Recipient Type"::"Cc", CcRecipients);
        EmailMessage.GetRecipients(Enum::"Email Recipient Type"::"Bcc", BccRecipients);
        if (Recipients.Count() + CcRecipients.Count() + BccRecipients.Count()) = 0 then
            Error(NoRecipientsErr);

        if not OAuthMgt.GetAppRegistrationForCurrentUser(AppReg) then
            Error(NoAppRegErr);

        if not OAuthMgt.GetOrAcquireToken(AppReg, AccessToken) then begin
            if not OAuthMgt.IsInteractiveSession() then
                Error(NoTokenBackgroundErr);
            Error('Could not acquire an email authorisation token. Check the App Registration configuration and try again.');
        end;

        // If first auth, fetch and cache home email for display in Email Accounts
        StoreHomeEmailIfNeeded(AccessToken);

        GraphMailMgt.SendEmail(EmailMessage, AccessToken);
    end;

    /// <summary>
    /// Returns one account with a fixed well-known GUID - the same pattern as BC's built-in
    /// Current User connector. All users share this one logical account; the connector resolves
    /// the actual sender credentials at send time via the current user's in-memory token.
    /// The email address shown reflects the user's cached home email if available.
    /// </summary>
    procedure GetAccounts(var EmailAccount: Record "Email Account")
    var
        UserToken: Record "W365 User Email Token";
        UserName: Code[50];
        AccountNameLbl: Label 'Current User Email API', Locked = true;
        NotConnectedLbl: Label '(sign in to connect your email)', Locked = true;
    begin
        EmailAccount.Init();
        EmailAccount."Account Id" := GetFixedAccountId();
        EmailAccount.Name := AccountNameLbl;
        EmailAccount.Connector := Enum::"Email Connector"::"W365 Guest Email";

        UserName := CopyStr(UserId(), 1, MaxStrLen(UserName));
        if UserToken.Get(UserName) and (UserToken."Home Email" <> '') then
            EmailAccount."Email Address" := CopyStr(UserToken."Home Email", 1, MaxStrLen(EmailAccount."Email Address"))
        else
            EmailAccount."Email Address" := CopyStr(NotConnectedLbl, 1, MaxStrLen(EmailAccount."Email Address"));

        if EmailAccount.Insert() then;
    end;

    /// <summary>
    /// Shows account information for the given account.
    /// Opens the App Registrations page so the admin can review or update the configuration.
    /// </summary>
    procedure ShowAccountInformation(AccountId: Guid)
    begin
        Page.Run(Page::"W365 App Registrations");
    end;

    /// <summary>
    /// Called when "Set Up Email Account" wizard reaches this connector.
    /// Verifies at least one App Registration exists and returns the fixed account.
    /// </summary>
    procedure RegisterAccount(var EmailAccount: Record "Email Account"): Boolean
    var
        AppReg: Record "W365 App Registration";
        NoAppRegErr: Label 'No App Registration has been configured. Open App Registrations from the action menu and create one before adding this account.';
    begin
        if not AppReg.FindFirst() then
            Error(NoAppRegErr);

        EmailAccount."Account Id" := GetFixedAccountId();
        EmailAccount.Name := 'Current User Email API';
        EmailAccount.Connector := Enum::"Email Connector"::"W365 Guest Email";
        exit(true);
    end;

    /// <summary>
    /// Deletes the account. Clears the in-memory token for the current session.
    /// </summary>
    procedure DeleteAccount(AccountId: Guid): Boolean
    var
        GraphSession: Codeunit "W365 Graph Session";
        ConfirmMsg: Label 'This will disconnect the email account for this session. The account can be reconnected on next send. Continue?';
    begin
        if not Confirm(ConfirmMsg) then
            exit(false);
        GraphSession.ClearToken();
        exit(true);
    end;

    /// <summary>Returns a base64-encoded logo. Empty string uses the default connector icon.</summary>
    procedure GetLogoAsBase64(): Text
    begin
        exit('');
    end;

    /// <summary>Short description shown in the "Set Up Email Account" wizard.</summary>
    procedure GetDescription(): Text[250]
    begin
        exit('Send emails from Business Central using your own work address via Microsoft Graph. One account - each user authenticates once per session via SSO.');
    end;

    // =========================================================================
    // "Email Connector v4" interface - read/reply (send-only connector)
    // =========================================================================

    /// <summary>Reply to an email. Not implemented - send-only connector.</summary>
    procedure Reply(var EmailMessage: Codeunit "Email Message"; AccountId: Guid)
    begin
        Error('Reply is not supported by this connector. Use Send instead.');
    end;

    /// <summary>Retrieve emails from inbox. Not implemented - send-only connector.</summary>
    procedure RetrieveEmails(AccountId: Guid; var EmailInbox: Record "Email Inbox"; var Filters: Record "Email Retrieval Filters" temporary)
    begin
        // Send-only connector - no inbox retrieval
    end;

    /// <summary>Mark email as read. Not implemented - send-only connector.</summary>
    procedure MarkAsRead(AccountId: Guid; ExternalId: Text)
    begin
        // Send-only connector - no read state management
    end;

    /// <summary>Get email folders. Not implemented - send-only connector.</summary>
    procedure GetEmailFolders(AccountId: Guid; var EmailFolders: Record "Email Folders" temporary)
    begin
        // Send-only connector - no folder management
    end;

    // =========================================================================
    // "Default Email Rate Limit" interface
    // =========================================================================

    /// <summary>
    /// Graph delegated Mail.Send has no strict per-connector rate limit we need to enforce.
    /// Returning 0 means no limit imposed by this connector.
    /// </summary>
    procedure GetDefaultEmailRateLimit(): Integer
    begin
        exit(0);
    end;

    // =========================================================================
    // Helpers
    // =========================================================================

    local procedure GetFixedAccountId(): Guid
    var
        AccountId: Guid;
    begin
        Evaluate(AccountId, 'a1b2c3d4-e5f6-7890-abcd-ef1234567890');
        exit(AccountId);
    end;

    [NonDebuggable]
    local procedure StoreHomeEmailIfNeeded(AccessToken: SecretText)
    var
        UserToken: Record "W365 User Email Token";
        OAuthMgt: Codeunit "W365 OAuth Mgt";
        GraphMailMgt: Codeunit "W365 Graph Mail Mgt";
        UserName: Code[50];
        HomeEmail: Text;
    begin
        UserName := CopyStr(UserId(), 1, MaxStrLen(UserName));
        if UserToken.Get(UserName) and (UserToken."Home Email" <> '') then
            exit; // Already cached

        HomeEmail := GraphMailMgt.GetCurrentUserEmail(AccessToken);
        if HomeEmail <> '' then
            OAuthMgt.StoreHomeEmail(HomeEmail);
    end;
}

