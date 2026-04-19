namespace Wingate365.GuestEmailAPI;

using System.Security.Authentication;

/// <summary>
/// Phase 3 validation test page. Tests whether BC's OAuth2 codeunit can trigger
/// an interactive consent prompt from within connector-style AL code.
///
/// HOW TO TEST:
/// 1. Deploy to a BC sandbox with W365 Email Setup configured (App ID, Client Secret).
/// 2. Open this page via the search bar: "W365 Auth Test".
/// 3. Run "Test: Interactive Auth" - a sign-in popup should appear.
///    Success = popup appears, you sign in, and the status updates to confirmed.
///    Fail    = no popup, error, or blank result.
/// 4. Run "Test: Auth from Send Context" - simulates what the connector Send() will do.
///    Success = same popup behaviour triggered from codeunit code (not page code).
///    Fail    = indicates the popup is blocked in non-page contexts.
/// 5. Check status messages carefully - they will tell us which path works.
/// </summary>
page 50111 "W365 Auth Test"
{
    Caption = 'W365 Phase 3 Auth Test';
    PageType = Card;
    UsageCategory = Administration;
    ApplicationArea = All;

    layout
    {
        area(content)
        {
            group(ResultGroup)
            {
                Caption = 'Test Results';

                field(StatusField; StatusText)
                {
                    ApplicationArea = All;
                    Caption = 'Status';
                    Editable = false;
                    MultiLine = true;
                    ToolTip = 'Result of the last test action.';
                }
                field(TokenAcquiredField; TokenAcquired)
                {
                    ApplicationArea = All;
                    Caption = 'Token Acquired';
                    Editable = false;
                    ToolTip = 'True if an access token was successfully acquired.';
                }
            }
        }
    }

    actions
    {
        area(processing)
        {
            action(TestInteractiveAuth)
            {
                ApplicationArea = All;
                Caption = 'Test: Interactive Auth';
                Image = Start;
                ToolTip = 'Calls OAuth2 codeunit directly from a page action. A consent popup should appear.';

                trigger OnAction()
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
                    NoSetupErr: Label 'W365 Email Setup not found. Configure App ID and client secret first.';
                    NoSecretErr: Label 'Client secret not configured. Use the Set Client Secret action on W365 Email Setup Card.';
                begin
                    if not Setup.Get('') then
                        Error(NoSetupErr);

                    if not IsolatedStorage.Get('W365_CS', DataScope::Company, ClientSecretText) then
                        Error(NoSecretErr);

                    ClientSecret := ClientSecretText;
                    Scopes.Add('https://graph.microsoft.com/Mail.Send');

                    StatusText := 'Calling OAuth2.AcquireTokenByAuthorizationCode with Prompt Interaction Consent...';
                    CurrPage.Update(false);

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
                        TokenAcquired := true;
                        StatusText := 'SUCCESS: Token acquired from page action context. Popup worked.';
                    end else begin
                        TokenAcquired := false;
                        StatusText := 'FAILED: ' + AuthCodeErr;
                    end;

                    CurrPage.Update(false);
                end;
            }

            action(TestFromSendContext)
            {
                ApplicationArea = All;
                Caption = 'Test: Auth from Send Context';
                Image = SendMail;
                ToolTip = 'Calls auth via a codeunit (simulating what the connector Send() will do). Tests whether popup works outside direct page code.';

                trigger OnAction()
                var
                    AuthTestMgt: Codeunit "W365 Auth Test Mgt";
                    ResultText: Text;
                    ResultAcquired: Boolean;
                begin
                    StatusText := 'Calling auth via W365 Auth Test Mgt codeunit (simulates Send() context)...';
                    CurrPage.Update(false);

                    AuthTestMgt.TryAcquireToken(ResultText, ResultAcquired);

                    StatusText := ResultText;
                    TokenAcquired := ResultAcquired;
                    CurrPage.Update(false);
                end;
            }

            action(ClearStatus)
            {
                ApplicationArea = All;
                Caption = 'Clear';
                Image = Delete;
                ToolTip = 'Clears the test result fields.';

                trigger OnAction()
                begin
                    StatusText := '';
                    TokenAcquired := false;
                    CurrPage.Update(false);
                end;
            }
        }
    }

    var
        StatusText: Text;
        TokenAcquired: Boolean;
}
