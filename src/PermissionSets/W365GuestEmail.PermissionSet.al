namespace Wingate365.GuestEmailAPI;

permissionset 50109 "W365 Guest Email"
{
    Caption = 'Guest Email';
    Assignable = true;

    Permissions =
        tabledata "W365 App Registration" = RIMD,
        tabledata "W365 User Email Token" = RIMD,
        tabledata "W365 Shared Mailbox Account" = RIMD,
        page "W365 App Registrations" = X,
        page "W365 App Registration Card" = X,
        page "W365 User Token List" = X,
        page "W365 Shared Mailbox Accounts" = X,
        page "W365 Shared Mailbox Card" = X,
        codeunit "W365 Graph Mail Mgt" = X,
        codeunit "W365 Graph Session" = X,
        codeunit "W365 OAuth Mgt" = X,
        codeunit "W365 Guest Email Connector" = X,
        codeunit "W365 Shared Mailbox Connector" = X;
}
