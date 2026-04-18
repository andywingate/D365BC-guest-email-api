namespace Wingate365.GuestEmailAPI;

using System.Email;
using System.Security.AccessControl;

codeunit 50107 "W365 Email Subscriber"
{
    Access = Internal;

    // =========================================================================
    // Helper available to other objects in this app
    // =========================================================================

    /// <summary>
    /// Returns true if the current user is an Entra B2B guest in this tenant.
    /// Detection is automatic: Entra always places #EXT# in the Authentication
    /// Email (UPN) of guest accounts. No manual flagging is required.
    /// Member accounts with a native UPN return false and use native BC email.
    /// </summary>
    procedure IsGuestUser(): Boolean
    var
        User: Record User;
    begin
        if not User.Get(UserSecurityId()) then
            exit(false);

        exit(StrPos(User."Authentication Email", '#EXT#') > 0);
    end;
}
