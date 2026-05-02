namespace Wingate365.GuestEmailAPI;

using System.Email;

codeunit 50118 "W365 Shared Mailbox Connector" implements "Email Connector", "Email Connector v4", "Default Email Rate Limit"
{
    Access = Internal;

    // =========================================================================
    // "Email Connector" interface
    // =========================================================================

    /// <summary>
    /// Sends an email from the shared mailbox associated with the given AccountId.
    /// Uses the App Registration linked to the shared mailbox to acquire a token,
    /// then calls Graph POST /v1.0/users/{mailboxEmail}/sendMail.
    /// The user must have delegated send-as or send-on-behalf-of access to the mailbox.
    /// The App Registration must have the Mail.Send.Shared delegated permission.
    /// </summary>
    procedure Send(EmailMessage: Codeunit "Email Message"; AccountId: Guid)
    var
        OAuthMgt: Codeunit "W365 OAuth Mgt";
        GraphMailMgt: Codeunit "W365 Graph Mail Mgt";
        MailboxAccount: Record "W365 Shared Mailbox Account";
        AppReg: Record "W365 App Registration";
        AccessToken: SecretText;
        Recipients: List of [Text];
        NoRecipientsErr: Label 'The email message has no recipients.';
        NoAccountErr: Label 'Shared mailbox account %1 not found.';
        NoAppRegErr: Label 'No App Registration is linked to shared mailbox %1.';
        NoTokenBackgroundErr: Label 'Email authentication is required. Sign in to Business Central interactively to renew your email authorisation, then retry.';
    begin
        EmailMessage.GetRecipients(Enum::"Email Recipient Type"::"To", Recipients);
        if Recipients.Count() = 0 then
            Error(NoRecipientsErr);

        if not MailboxAccount.GetBySystemId(AccountId) then
            Error(NoAccountErr, AccountId);

        if not AppReg.Get(MailboxAccount."App Registration Code") then
            Error(NoAppRegErr, MailboxAccount."Code");

        if not OAuthMgt.GetOrAcquireToken(AppReg, AccessToken) then begin
            if not OAuthMgt.IsInteractiveSession() then
                Error(NoTokenBackgroundErr);
            Error('Could not acquire an email authorisation token for shared mailbox %1. Check the App Registration configuration.', MailboxAccount."Display Name");
        end;

        GraphMailMgt.SendEmailAsSharedMailbox(EmailMessage, AccessToken, MailboxAccount."Mailbox Email");
    end;

    /// <summary>
    /// Returns one Email Account entry per row in W365 Shared Mailbox Account.
    /// Each account uses the row's SystemId as the AccountId for routing.
    /// </summary>
    procedure GetAccounts(var EmailAccount: Record "Email Account")
    var
        MailboxAccount: Record "W365 Shared Mailbox Account";
    begin
        if not MailboxAccount.FindSet() then
            exit;

        repeat
            EmailAccount.Init();
            EmailAccount."Account Id" := MailboxAccount.SystemId;
            EmailAccount.Name := MailboxAccount."Display Name";
            EmailAccount."Email Address" := CopyStr(MailboxAccount."Mailbox Email", 1, MaxStrLen(EmailAccount."Email Address"));
            EmailAccount.Connector := Enum::"Email Connector"::"W365 Shared Mailbox";
            if EmailAccount.Insert() then;
        until MailboxAccount.Next() = 0;
    end;

    /// <summary>Opens the Shared Mailbox Card for the given account.</summary>
    procedure ShowAccountInformation(AccountId: Guid)
    var
        MailboxAccount: Record "W365 Shared Mailbox Account";
    begin
        if MailboxAccount.GetBySystemId(AccountId) then
            Page.Run(Page::"W365 Shared Mailbox Card", MailboxAccount);
    end;

    /// <summary>
    /// Registers a new shared mailbox account. Opens the Shared Mailbox Accounts list
    /// so the admin can create or select a mailbox.
    /// </summary>
    procedure RegisterAccount(var EmailAccount: Record "Email Account"): Boolean
    var
        MailboxAccount: Record "W365 Shared Mailbox Account";
        NoMailboxErr: Label 'No Shared Mailbox Accounts have been configured. Create one on the Shared Mailbox Accounts page first.';
    begin
        if not MailboxAccount.FindFirst() then
            Error(NoMailboxErr);

        // Use the first mailbox as the registered account; the admin can configure more later
        EmailAccount."Account Id" := MailboxAccount.SystemId;
        EmailAccount.Name := MailboxAccount."Display Name";
        EmailAccount."Email Address" := CopyStr(MailboxAccount."Mailbox Email", 1, MaxStrLen(EmailAccount."Email Address"));
        EmailAccount.Connector := Enum::"Email Connector"::"W365 Shared Mailbox";
        exit(true);
    end;

    /// <summary>Deletes the shared mailbox account row identified by AccountId.</summary>
    procedure DeleteAccount(AccountId: Guid): Boolean
    var
        MailboxAccount: Record "W365 Shared Mailbox Account";
        ConfirmMsg: Label 'Delete this shared mailbox account from BC? This does not affect the mailbox itself.';
    begin
        if not Confirm(ConfirmMsg) then
            exit(false);
        if MailboxAccount.GetBySystemId(AccountId) then
            MailboxAccount.Delete();
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
        exit('Send emails from a shared Microsoft 365 mailbox via Microsoft Graph. Multiple mailboxes can be configured; each appears as a separate Email Account.');
    end;

    // =========================================================================
    // "Email Connector v4" interface - send-only connector
    // =========================================================================

    procedure Reply(var EmailMessage: Codeunit "Email Message"; AccountId: Guid)
    begin
        Error('Reply is not supported by this connector. Use Send instead.');
    end;

    procedure RetrieveEmails(AccountId: Guid; var EmailInbox: Record "Email Inbox"; var Filters: Record "Email Retrieval Filters" temporary)
    begin
        // Send-only connector - no inbox retrieval
    end;

    procedure MarkAsRead(AccountId: Guid; ExternalId: Text)
    begin
        // Send-only connector - no read state management
    end;

    procedure GetEmailFolders(AccountId: Guid; var EmailFolders: Record "Email Folders" temporary)
    begin
        // Send-only connector - no folder management
    end;

    // =========================================================================
    // "Default Email Rate Limit" interface
    // =========================================================================

    procedure GetDefaultEmailRateLimit(): Integer
    begin
        exit(0);
    end;
}
