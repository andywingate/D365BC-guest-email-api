namespace Wingate365.GuestEmailAPI;

permissionset 50110 "W365 Guest Email Admin"
{
    Caption = 'Guest Email Admin';
    Assignable = true;

    // Grants full read/insert/modify/delete access to configuration tables
    // and access to all setup pages. Assign to administrators only.
    Permissions =
        tabledata "W365 App Registration" = RIMD,
        tabledata "W365 Shared Mailbox Account" = RIMD,
        page "W365 App Registrations" = X,
        page "W365 App Registration Card" = X,
        page "W365 Shared Mailbox Accounts" = X,
        page "W365 Shared Mailbox Card" = X,
        codeunit "W365 OAuth Mgt" = X;
}
