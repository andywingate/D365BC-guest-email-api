namespace Wingate365.GuestEmailAPI;

using System.Security.Authentication;
using System.Security.AccessControl;

codeunit 50106 "W365 OAuth Mgt"
{
    Access = Internal;

    // -------------------------------------------------------------------------
    // IsolatedStorage keys
    // -------------------------------------------------------------------------
    // 'W365_CS_{AppIdHex}' - client secret (SecretText, DataScope::Company, SetEncrypted)
    //   where AppIdHex = Format(AppId) with braces and dashes stripped

    // -------------------------------------------------------------------------
    // App Registration lookup
    // -------------------------------------------------------------------------

    /// <summary>
    /// Resolves the App Registration to use for the current user.
    /// Extracts the user's home domain from their Authentication Email, then
    /// looks for a matching Domain Filter. Falls back to the default registration.
    /// Returns false if no registration is found.
    /// </summary>
    procedure GetAppRegistrationForCurrentUser(var AppReg: Record "W365 App Registration"): Boolean
    var
        HomeDomain: Text;
    begin
        HomeDomain := GetCurrentUserHomeDomain();
        exit(GetAppRegistration(HomeDomain, AppReg));
    end;

    /// <summary>
    /// Resolves the App Registration for a given home domain.
    /// Tries exact Domain Filter match first, then falls back to Is Default = true.
    /// </summary>
    procedure GetAppRegistration(HomeDomain: Text; var AppReg: Record "W365 App Registration"): Boolean
    begin
        if HomeDomain <> '' then begin
            AppReg.SetRange("Domain Filter", HomeDomain);
            if AppReg.FindFirst() then
                exit(true);
            AppReg.SetRange("Domain Filter");
        end;

        AppReg.SetRange("Is Default", true);
        if AppReg.FindFirst() then
            exit(true);

        AppReg.SetRange("Is Default");
        // Last resort: use the first registration in the table.
        // This applies when there is exactly one registration and neither Domain Filter nor Is Default is set.
        // For predictable behaviour with multiple registrations, always set Is Default on one of them.
        exit(AppReg.FindFirst());
    end;

    // -------------------------------------------------------------------------
    // Home domain extraction
    // -------------------------------------------------------------------------

    /// <summary>
    /// Returns the home domain for the current BC user, derived from their Authentication Email.
    /// Guest users have Authentication Email in the form user_domain.com#EXT#@hosttenant.onmicrosoft.com.
    /// Member users have a standard user@domain.com UPN.
    /// </summary>
    procedure GetCurrentUserHomeDomain(): Text
    var
        BCUser: Record User;
        AuthEmail: Text;
    begin
        if not BCUser.Get(UserSecurityId()) then
            exit('');
        AuthEmail := BCUser."Authentication Email";
        exit(ExtractHomeDomain(AuthEmail));
    end;

    /// <summary>
    /// Extracts the home domain from an Authentication Email address.
    /// Handles both member format (user@domain.com) and guest format
    /// (user_domain.com#EXT#@hosttenant.onmicrosoft.com).
    /// </summary>
    procedure ExtractHomeDomain(AuthEmail: Text): Text
    var
        ExtPos: Integer;
        GuestPrefix: Text;
        AtPos: Integer;
        i: Integer;
        LastUnderPos: Integer;
    begin
        if AuthEmail = '' then
            exit('');

        // Guest format: {localpart}_{homedomain}#EXT#@hosttenant.onmicrosoft.com
        // e.g. andy_contoso.com#EXT#@wingate365.onmicrosoft.com -> contoso.com
        ExtPos := StrPos(AuthEmail, '#EXT#');
        if ExtPos > 0 then begin
            GuestPrefix := CopyStr(AuthEmail, 1, ExtPos - 1);
            // Find the last underscore - everything after it is the home domain
            LastUnderPos := 0;
            for i := 1 to StrLen(GuestPrefix) do
                if GuestPrefix[i] = '_' then
                    LastUnderPos := i;
            if LastUnderPos > 0 then
                exit(CopyStr(GuestPrefix, LastUnderPos + 1));
        end;

        // Member format: user@domain.com
        AtPos := StrPos(AuthEmail, '@');
        if AtPos > 0 then
            exit(CopyStr(AuthEmail, AtPos + 1));

        exit('');
    end;

    // -------------------------------------------------------------------------
    // Client secret management
    // -------------------------------------------------------------------------

    /// <summary>
    /// Stores the client secret for an App Registration in encrypted Company-scoped IsolatedStorage.
    /// The value is never read back to the UI.
    /// </summary>
    procedure SetClientSecret(AppId: Guid; ClientSecret: SecretText)
    var
        AppReg: Record "W365 App Registration";
        SecretKey: Text;
    begin
        AppReg."App ID" := AppId;
        AppReg."Code" := '';
        SecretKey := AppReg.GetClientSecretKey();
        IsolatedStorage.Set(SecretKey, ClientSecret, DataScope::Company);
    end;

    /// <summary>Returns true if a client secret has been configured for the given App ID.</summary>
    procedure HasClientSecret(AppId: Guid): Boolean
    var
        AppReg: Record "W365 App Registration";
    begin
        AppReg."App ID" := AppId;
        AppReg."Code" := '';
        exit(IsolatedStorage.Contains(AppReg.GetClientSecretKey(), DataScope::Company));
    end;

    /// <summary>
    /// Retrieves the client secret for an App Registration from IsolatedStorage.
    /// Returns false if no secret is found.
    /// </summary>
    [NonDebuggable]
    procedure GetClientSecret(AppId: Guid; var ClientSecret: SecretText): Boolean
    var
        AppReg: Record "W365 App Registration";
        SecretKey: Text;
    begin
        AppReg."App ID" := AppId;
        AppReg."Code" := '';
        SecretKey := AppReg.GetClientSecretKey();
        if not IsolatedStorage.Contains(SecretKey, DataScope::Company) then
            exit(false);
        IsolatedStorage.Get(SecretKey, DataScope::Company, ClientSecret);
        exit(true);
    end;

    /// <summary>Removes the client secret for an App Registration from IsolatedStorage.</summary>
    procedure ClearClientSecret(AppId: Guid)
    var
        AppReg: Record "W365 App Registration";
        SecretKey: Text;
    begin
        AppReg."App ID" := AppId;
        AppReg."Code" := '';
        SecretKey := AppReg.GetClientSecretKey();
        if IsolatedStorage.Contains(SecretKey, DataScope::Company) then
            IsolatedStorage.Delete(SecretKey, DataScope::Company);
    end;

    // -------------------------------------------------------------------------
    // Token acquisition
    // -------------------------------------------------------------------------

    /// <summary>
    /// Attempts to acquire an access token for the current user using the resolved App Registration.
    /// First tries silent SSO (PromptInteraction::None). If that fails and the session is
    /// interactive, tries a full consent prompt. Returns false if acquisition fails.
    /// Never stores the token - it lives in W365 Graph Session (SingleInstance) memory only.
    /// </summary>
    [NonDebuggable]
    procedure GetOrAcquireToken(var AppReg: Record "W365 App Registration"; var AccessToken: SecretText): Boolean
    var
        GraphSession: Codeunit "W365 Graph Session";
        ClientSecret: SecretText;
        NoSecretErr: Label 'No client secret has been configured for App Registration %1. Open App Registrations and set the client secret.';
        NoAppRegErr: Label 'No App Registration found for the current user. Create and configure one on the App Registrations page.';
    begin
        // Check session cache first
        if GraphSession.HasValidToken() then begin
            GraphSession.GetToken(AccessToken);
            exit(true);
        end;

        if IsNullGuid(AppReg."App ID") then
            Error(NoAppRegErr);

        if not GetClientSecret(AppReg."App ID", ClientSecret) then
            Error(NoSecretErr, AppReg."Code");

        // Try silent SSO first (no popup)
        if TryAcquireToken(AppReg, ClientSecret, "Prompt Interaction"::None, AccessToken) then begin
            GraphSession.SetToken(AccessToken);
            exit(true);
        end;

        // Silent failed - try interactive if in a client session
        if not IsInteractiveSession() then
            exit(false);

        if TryAcquireToken(AppReg, ClientSecret, "Prompt Interaction"::Login, AccessToken) then begin
            GraphSession.SetToken(AccessToken);
            exit(true);
        end;

        exit(false);
    end;

    [NonDebuggable]
    local procedure TryAcquireToken(AppReg: Record "W365 App Registration"; ClientSecret: SecretText; PromptInteraction: Enum "Prompt Interaction"; var AccessToken: SecretText): Boolean
    var
        OAuth20: Codeunit OAuth2;
        Scopes: List of [Text];
        AuthorityUrl: Text;
        AuthCodeErr: Text;
        RedirectUrl: Label 'https://businesscentral.dynamics.com/OAuthLanding.htm', Locked = true;
    begin
        Scopes.Add('https://graph.microsoft.com/Mail.Send');
        AuthorityUrl := AppReg.GetAuthorityUrl();

        exit(OAuth20.AcquireTokenByAuthorizationCode(
            Format(AppReg."App ID"),
            ClientSecret,
            AuthorityUrl,
            RedirectUrl,
            Scopes,
            PromptInteraction,
            AccessToken,
            AuthCodeErr));
    end;

    // -------------------------------------------------------------------------
    // Session helpers
    // -------------------------------------------------------------------------

    /// <summary>Returns true if the current session is an interactive client session.</summary>
    procedure IsInteractiveSession(): Boolean
    begin
        case CurrentClientType() of
            ClientType::Web, ClientType::Windows, ClientType::Tablet, ClientType::Phone:
                exit(true);
            else
                exit(false);
        end;
    end;

    // -------------------------------------------------------------------------
    // Home email cache
    // -------------------------------------------------------------------------

    /// <summary>
    /// Stores the user's home email address (fetched from Graph /me) in the cache table.
    /// Called after a successful first authentication to populate the email address for display.
    /// </summary>
    procedure StoreHomeEmail(HomeEmail: Text)
    var
        UserToken: Record "W365 User Email Token";
        UserName: Code[50];
    begin
        if HomeEmail = '' then
            exit;

        UserName := CopyStr(UserId(), 1, MaxStrLen(UserName));
        if not UserToken.Get(UserName) then begin
            UserToken.Init();
            UserToken."User Name" := UserName;
            UserToken.Insert();
        end;
        UserToken."Home Email" := CopyStr(HomeEmail, 1, MaxStrLen(UserToken."Home Email"));
        UserToken.Modify();
    end;
}

