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
    /// Called by BC's Email module when sending any email assigned to this connector.
    /// </summary>
    procedure Send(EmailMessage: Codeunit "Email Message"; AccountId: Guid)
    var
        GraphMailMgt: Codeunit "W365 Graph Mail Mgt";
        Recipients: List of [Text];
        ToAddress: Text;
        Subject: Text;
        Body: Text;
        NoRecipientsErr: Label 'The email message has no recipients.';
    begin
        EmailMessage.GetRecipients(Enum::"Email Recipient Type"::"To", Recipients);
        if Recipients.Count() = 0 then
            Error(NoRecipientsErr);

        // Graph sendMail sends to all recipients; we pass the first To address.
        // Full multi-recipient support is a Phase 2 enhancement.
        Recipients.Get(1, ToAddress);
        Subject := EmailMessage.GetSubject();
        Body := EmailMessage.GetBody();

        GraphMailMgt.SendEmail(ToAddress, Subject, Body);
    end;

    /// <summary>
    /// Returns one account with a fixed well-known GUID - the same pattern as BC's built-in
    /// "Current User" connector. There is only ever one entry in Email Accounts; set it as
    /// the system default once and every user's sends resolve to their own Graph token at
    /// runtime via UserSecurityId(). The email address shown reflects the current user's
    /// home address from their stored token, so each user sees their own address.
    /// </summary>
    procedure GetAccounts(var EmailAccount: Record "Email Account")
    var
        UserToken: Record "W365 User Email Token";
        UserName: Code[50];
        AccountNameLbl: Label 'Current User Email API', Locked = true;
        NotConnectedLbl: Label '(not connected - run Email Account Setup)', Locked = true;
    begin
        EmailAccount.Init();
        EmailAccount."Account Id" := GetFixedAccountId();
        EmailAccount.Name := AccountNameLbl;
        EmailAccount.Connector := Enum::"Email Connector"::"W365 Guest Email";

        // Show the current user's home email if they have an active token;
        // otherwise show a placeholder so the account is still visible and
        // can be set as default before users have individually consented.
        UserName := CopyStr(UserId(), 1, MaxStrLen(UserName));
        if UserToken.Get(UserName) and (UserToken."Consent Status" = "W365 Consent Status"::Active) then begin
            if UserToken."Home Email" <> '' then
                EmailAccount."Email Address" := CopyStr(UserToken."Home Email", 1, MaxStrLen(EmailAccount."Email Address"))
            else
                EmailAccount."Email Address" := CopyStr(UserToken."User Name", 1, MaxStrLen(EmailAccount."Email Address"));
        end else
            EmailAccount."Email Address" := CopyStr(NotConnectedLbl, 1, MaxStrLen(EmailAccount."Email Address"));

        if EmailAccount.Insert() then;
    end;

    /// <summary>
    /// Fixed well-known GUID used as the single Account Id for this connector.
    /// All users share this one logical account; the connector resolves the actual
    /// sender credentials at send time via UserSecurityId().
    /// </summary>
    local procedure GetFixedAccountId(): Guid
    var
        AccountId: Guid;
    begin
        Evaluate(AccountId, 'a1b2c3d4-e5f6-7890-abcd-ef1234567890');
        exit(AccountId);
    end;

    /// <summary>
    /// Shows the account information page for the given account.
    /// Opens the consent page for the relevant user.
    /// </summary>
    procedure ShowAccountInformation(AccountId: Guid)
    begin
        // No account information page - the Email Accounts list row is sufficient
    end;

    /// <summary>
    /// Called when "Set Up Email Account" wizard reaches this connector.
    /// Opens the consent page - the user connects their account there.
    /// Returns true once the user has an active token.
    /// </summary>
    procedure RegisterAccount(var EmailAccount: Record "Email Account"): Boolean
    var
        UserToken: Record "W365 User Email Token";
        UserName: Code[50];
        Setup: Record "W365 Email Setup";
        NoSetupErr: Label 'W365 Email Setup has not been configured. Ask your administrator to complete the setup first.';
    begin
        if not Setup.Get('') then
            Error(NoSetupErr);

        Page.RunModal(Page::"W365 OAuth Consent");

        // Check if consent was granted during the modal
        UserName := CopyStr(UserId(), 1, MaxStrLen(UserName));
        if not UserToken.Get(UserName) then
            exit(false);
        if UserToken."Consent Status" <> "W365 Consent Status"::Active then
            exit(false);

        EmailAccount."Account Id" := GetFixedAccountId();
        if UserToken."Home Email" <> '' then
            EmailAccount."Email Address" := CopyStr(UserToken."Home Email", 1, MaxStrLen(EmailAccount."Email Address"))
        else
            EmailAccount."Email Address" := CopyStr(UserToken."User Name", 1, MaxStrLen(EmailAccount."Email Address"));
        EmailAccount.Name := 'Current User Email API';
        EmailAccount.Connector := Enum::"Email Connector"::"W365 Guest Email";
        exit(true);
    end;

    /// <summary>
    /// Deletes the stored token for the given account, disconnecting the user.
    /// </summary>
    procedure DeleteAccount(AccountId: Guid): Boolean
    var
        OAuthMgt: Codeunit "W365 OAuth Mgt";
        ConfirmMsg: Label 'This will disconnect the email account. The user will need to reconnect. Continue?';
    begin
        if not Confirm(ConfirmMsg) then
            exit(false);
        OAuthMgt.ClearTokens();
        exit(true);
    end;

    /// <summary>
    /// Returns a base64-encoded logo shown in the Email Account setup wizard.
    /// Returning empty string uses the default connector icon.
    /// </summary>
    procedure GetLogoAsBase64(): Text
    begin
        exit('');
    end;

    /// <summary>
    /// Short description shown in the "Set Up Email Account" wizard account type list.
    /// </summary>
    procedure GetDescription(): Text[250]
    begin
        exit('Send emails from Business Central using your own work address via Microsoft Graph. One account - each user sends as themselves.');
    end;

    // =========================================================================
    // "Email Connector v4" interface - read/reply (not implemented in Phase 2)
    // =========================================================================

    /// <summary>Reply to an email. Not implemented - send-only connector.</summary>
    procedure Reply(var EmailMessage: Codeunit "Email Message"; AccountId: Guid)
    begin
        Error('Reply is not supported by the Guest Email connector. Use Send instead.');
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
}
