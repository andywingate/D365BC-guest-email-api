namespace Wingate365.GuestEmailAPI;

using System.Email;

pageextension 50100 "W365 Email Accounts Ext" extends "Email Accounts"
{
    // Disable "Set as default" for Guest Email connector rows.
    // Guest accounts are bound to a specific user's OAuth token - setting one
    // as the global default would cause all BC email sends to attempt to use
    // that individual user's token, which will fail for any other sender.
    actions
    {
        modify(MakeDefault)
        {
            Enabled = not IsGuestAccount;
        }
    }

    trigger OnAfterGetRecord()
    begin
        IsGuestAccount := Rec.Connector = Enum::"Email Connector"::"W365 Guest Email";
    end;

    var
        IsGuestAccount: Boolean;
}
