namespace Wingate365.GuestEmailAPI;

page 50108 "W365 OAuth Consent"
{
    Caption = 'Connect Current User Email API';
    PageType = Card;
    UsageCategory = Tasks;
    ApplicationArea = All;

    layout
    {
        area(content)
        {
            group(NotConnectedGroup)
            {
                ShowCaption = false;
                Visible = not TokenIsActive;

                group(NotConnectedIntro)
                {
                    ShowCaption = false;
                    InstructionalText = 'Click "Connect my Email" below. A sign-in window will open - sign in with your work account and approve access. That is it - no extra steps needed.';
                }
                field(StatusNotConnectedField; StatusMessage)
                {
                    ApplicationArea = All;
                    Caption = ' ';
                    Editable = false;
                    StyleExpr = TokenStatusStyle;
                    ToolTip = 'Your current connection status.';
                }
            }
            group(ConnectedGroup)
            {
                ShowCaption = false;
                Visible = TokenIsActive;

                group(ConnectedIntro)
                {
                    ShowCaption = false;
                    InstructionalText = 'Your email account is connected. Business Central will send emails on your behalf using your work address.';
                }
                field(StatusConnectedField; StatusMessage)
                {
                    ApplicationArea = All;
                    Caption = ' ';
                    Editable = false;
                    StyleExpr = TokenStatusStyle;
                    ToolTip = 'Your current connection status.';
                }
            }
            group(AddinGroup)
            {
                ShowCaption = false;

                usercontrol(OAuthPopup; "W365 OAuth Popup")
                {
                    ApplicationArea = All;

                    trigger CodeReceived(RedirectUrl: Text)
                    var
                        OAuthMgt: Codeunit "W365 OAuth Mgt";
                        SuccessMsg: Label 'You are connected! Emails sent from Business Central will now use your work address.';
                    begin
                        if OAuthMgt.ExchangeCodeForToken(RedirectUrl) then begin
                            LoadStatus();
                            CurrPage.Update(false);
                            Message(SuccessMsg);
                        end;
                    end;

                    trigger PopupClosed()
                    begin
                        LoadStatus();
                        CurrPage.Update(false);
                    end;

                    trigger PopupError(ErrorMessage: Text)
                    begin
                        Error(ErrorMessage);
                    end;
                }
            }
        }
    }

    actions
    {
        area(processing)
        {
            action(ConnectAction)
            {
                ApplicationArea = All;
                Caption = 'Connect my Email';
                Image = Setup;
                Enabled = not TokenIsActive;
                ToolTip = 'Opens a sign-in window. Approve access and the connection completes automatically.';

                trigger OnAction()
                var
                    OAuthMgt: Codeunit "W365 OAuth Mgt";
                    Setup: Record "W365 Email Setup";
                    NoSetupErr: Label 'Email connection has not been configured yet. Please contact your administrator.';
                begin
                    if not Setup.Get('') then
                        Error(NoSetupErr);
                    CurrPage.OAuthPopup.OpenAuthPopup(OAuthMgt.GetAuthorizationUrl());
                end;
            }
            action(DisconnectAction)
            {
                ApplicationArea = All;
                Caption = 'Disconnect';
                Image = Delete;
                Enabled = TokenIsActive;
                ToolTip = 'Removes your stored connection. You can reconnect at any time.';

                trigger OnAction()
                var
                    OAuthMgt: Codeunit "W365 OAuth Mgt";
                    ConfirmMsg: Label 'This will disconnect your email account. You can reconnect at any time by clicking Connect my Email. Continue?';
                begin
                    if Confirm(ConfirmMsg) then begin
                        OAuthMgt.ClearTokens();
                        LoadStatus();
                        CurrPage.Update(false);
                    end;
                end;
            }
            action(TestConnectionAction)
            {
                ApplicationArea = All;
                Caption = 'Email Accounts';
                Image = LaunchWeb;
                Enabled = TokenIsActive;
                ToolTip = 'Opens the Email Accounts page.';

                trigger OnAction()
                begin
                    Page.Run(8887);
                end;
            }
        }
        area(Promoted)
        {
            actionref(ConnectRef; ConnectAction) { }
            actionref(TestConnectionRef; TestConnectionAction) { }
            actionref(DisconnectRef; DisconnectAction) { }
        }
    }

    trigger OnOpenPage()
    begin
        LoadStatus();
    end;

    var
        StatusMessage: Text;
        TokenExpiryDisplay: Text;
        TokenStatusStyle: Text;
        TokenIsActive: Boolean;

    local procedure LoadStatus()
    var
        UserToken: Record "W365 User Email Token";
        UserName: Code[50];
        ConnectedTxt: Label 'Connected';
        ExpiredTxt: Label 'Connected';
        ErrorTxt: Label 'Not connected - please disconnect and reconnect';
        NoneTxt: Label 'Not connected';
    begin
        UserName := CopyStr(UserId(), 1, MaxStrLen(UserName));

        if UserToken.Get(UserName) then begin
            case UserToken."Consent Status" of
                "W365 Consent Status"::Active:
                    begin
                        if UserToken.IsTokenExpired() then begin
                            StatusMessage := ExpiredTxt;
                            TokenStatusStyle := 'Ambiguous';
                            TokenIsActive := true;
                        end else begin
                            StatusMessage := ConnectedTxt;
                            TokenStatusStyle := 'Favorable';
                            TokenIsActive := true;
                        end;
                        TokenExpiryDisplay := Format(UserToken."Token Expiry");
                    end;
                "W365 Consent Status"::Error:
                    begin
                        StatusMessage := ErrorTxt;
                        TokenStatusStyle := 'Unfavorable';
                        TokenIsActive := false;
                        TokenExpiryDisplay := '';
                    end;
                else begin
                    StatusMessage := NoneTxt;
                    TokenStatusStyle := 'Subordinate';
                    TokenIsActive := false;
                    TokenExpiryDisplay := '';
                end;
            end;
        end else begin
            StatusMessage := NoneTxt;
            TokenStatusStyle := 'Subordinate';
            TokenIsActive := false;
            TokenExpiryDisplay := '';
        end;
    end;
}
