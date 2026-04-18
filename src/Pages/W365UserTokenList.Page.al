namespace Wingate365.GuestEmailAPI;

page 50104 "W365 User Token List"
{
    Caption = 'W365 User Token Status';
    PageType = List;
    SourceTable = "W365 User Email Token";
    UsageCategory = Administration;
    ApplicationArea = All;
    InsertAllowed = false;
    ModifyAllowed = false;
    DeleteAllowed = true;

    layout
    {
        area(content)
        {
            repeater(TokenList)
            {
                field("User Name"; Rec."User Name")
                {
                    ApplicationArea = All;
                    ToolTip = 'The BC user name.';
                }
                field("Consent Status"; Rec."Consent Status")
                {
                    ApplicationArea = All;
                    ToolTip = 'Current token status for this user.';
                    StyleExpr = StatusStyle;
                }
                field("Token Expiry"; Rec."Token Expiry")
                {
                    ApplicationArea = All;
                    ToolTip = 'When the current access token expires. The app refreshes automatically before expiry.';
                }
                field("Last Error"; Rec."Last Error")
                {
                    ApplicationArea = All;
                    ToolTip = 'Last error recorded during token acquisition or refresh.';
                }
            }
        }
        area(factboxes)
        {
            systempart(Links; Links) { ApplicationArea = RecordLinks; }
            systempart(Notes; Notes) { ApplicationArea = Notes; }
        }
    }

    actions
    {
        area(processing)
        {
            action(StartConsent)
            {
                ApplicationArea = All;
                Caption = 'Authorise (Consent Flow)';
                Image = Setup;
                ToolTip = 'Open the OAuth consent page for the current user to grant Mail.Send access.';

                trigger OnAction()
                begin
                    Page.Run(Page::"W365 OAuth Consent");
                end;
            }
            action(ClearToken)
            {
                ApplicationArea = All;
                Caption = 'Clear Token';
                Image = Delete;
                ToolTip = 'Clears the stored token for the current user. They will need to re-authorise before sending email.';

                trigger OnAction()
                var
                    OAuthMgt: Codeunit "W365 OAuth Mgt";
                    ConfirmMsg: Label 'Clear the stored token for the current user? They will need to re-authorise before sending email.';
                begin
                    if Confirm(ConfirmMsg) then begin
                        OAuthMgt.ClearTokens();
                        CurrPage.Update(false);
                    end;
                end;
            }
        }
        area(navigation)
        {
            action(EmailSetup)
            {
                ApplicationArea = All;
                Caption = 'Email Setup';
                Image = Setup;
                RunObject = Page "W365 Email Setup Card";
                ToolTip = 'Open the W365 Email Setup card to configure the Entra app registration.';
            }
        }
        area(Promoted)
        {
            actionref(StartConsentRef; StartConsent) { }
        }
    }

    trigger OnAfterGetRecord()
    begin
        SetStatusStyle();
    end;

    var
        StatusStyle: Text;

    local procedure SetStatusStyle()
    begin
        case Rec."Consent Status" of
            "W365 Consent Status"::Active:
                StatusStyle := 'Favorable';
            "W365 Consent Status"::Error:
                StatusStyle := 'Unfavorable';
            else
                StatusStyle := 'Subordinate';
        end;
    end;
}
