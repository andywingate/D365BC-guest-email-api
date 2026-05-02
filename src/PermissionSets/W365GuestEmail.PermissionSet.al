namespace Wingate365.GuestEmailAPI;

permissionset 50109 "W365 Guest Email"
{
    Caption = 'Guest Email';
    Assignable = true;

    // Grants send-only access. Assign to all users who need to send email via the
    // Current User or Shared Mailbox connectors. For administration of App Registrations
    // and Shared Mailbox Accounts, also assign "Guest Email Admin" (50110).
    Permissions =
        tabledata "W365 App Registration" = R,
        tabledata "W365 User Email Token" = RIMD,
        tabledata "W365 Shared Mailbox Account" = R,
        page "W365 User Token List" = X,
        codeunit "W365 Graph Mail Mgt" = X,
        codeunit "W365 Graph Session" = X,
        codeunit "W365 OAuth Mgt" = X,
        codeunit "W365 Guest Email Connector" = X,
        codeunit "W365 Shared Mailbox Connector" = X;
}
