namespace Wingate365.GuestEmailAPI;

using System.Email;

enumextension 50101 "W365 Shared Mailbox Connector" extends "Email Connector"
{
    value(50101; "W365 Shared Mailbox")
    {
        Caption = 'Shared Mailbox (Microsoft Graph)';
        Implementation = "Email Connector" = "W365 Shared Mailbox Connector",
                         "Email Connector v4" = "W365 Shared Mailbox Connector",
                         "Default Email Rate Limit" = "W365 Shared Mailbox Connector";
    }
}
