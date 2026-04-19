namespace Wingate365.GuestEmailAPI;

using System.Security.Authentication;

/// <summary>
/// Phase 3 validation helper. Called from the W365 Auth Test page to simulate
/// the "called from a codeunit" context that the connector Send() will use.
/// This tests whether BC's OAuth2 codeunit triggers an interactive popup
/// when invoked from codeunit code rather than directly from a page action trigger.
/// </summary>
codeunit 50112 "W365 Auth Test Mgt"
{
    Access = Internal;

    /// <summary>
    /// Attempts to acquire an access token using BC's OAuth2 codeunit.
    /// Returns a human-readable result string and whether the token was acquired.
    /// Designed to be called from a page action to simulate the Send() call chain.
    /// </summary>
    procedure TryAcquireToken(var ResultText: Text; var AcquiredToken: Boolean)
    var
        OAuth20: Codeunit OAuth2;
        Setup: Record "W365 Email Setup";
        ClientSecret: SecretText;
        AccessToken: SecretText;
        ClientSecretText: Text;
        AuthCodeErr: Text;
        Scopes: List of [Text];
        AuthorityUrl: Label 'https://login.microsoftonline.com/common/oauth2/v2.0', Locked = true;
        RedirectUrl: Label 'https://businesscentral.dynamics.com/OAuthLanding.htm', Locked = true;
        NoSetupErr: Label 'W365 Email Setup not found.';
        NoSecretErr: Label 'Client secret not configured.';
    begin
        if not Setup.Get('') then begin
            ResultText := 'FAILED: ' + NoSetupErr;
            exit;
        end;

        if not IsolatedStorage.Get('W365_CS', DataScope::Company, ClientSecretText) then begin
            ResultText := 'FAILED: ' + NoSecretErr;
            exit;
        end;

        ClientSecret := ClientSecretText;
        Scopes.Add('https://graph.microsoft.com/Mail.Send');

        if OAuth20.AcquireTokenByAuthorizationCode(
            Setup."App ID",
            ClientSecret,
            AuthorityUrl,
            RedirectUrl,
            Scopes,
            "Prompt Interaction"::Consent,
            AccessToken,
            AuthCodeErr)
        then begin
            AcquiredToken := true;
            ResultText := 'SUCCESS: Token acquired from codeunit context. Popup worked from Send() simulation.';
        end else begin
            AcquiredToken := false;
            ResultText := 'FAILED: ' + AuthCodeErr;
        end;
    end;
}
