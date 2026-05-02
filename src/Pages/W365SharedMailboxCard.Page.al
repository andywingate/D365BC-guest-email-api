namespace Wingate365.GuestEmailAPI;

page 50117 "W365 Shared Mailbox Card"
{
    Caption = 'Shared Mailbox Account';
    PageType = Card;
    SourceTable = "W365 Shared Mailbox Account";
    UsageCategory = Administration;
    ApplicationArea = All;

    layout
    {
        area(content)
        {
            group(MailboxDetails)
            {
                Caption = 'Mailbox Details';

                field("Code"; Rec."Code")
                {
                    ApplicationArea = All;
                    ToolTip = 'Short code identifying this shared mailbox account (e.g. SALES).';
                }
                field("Display Name"; Rec."Display Name")
                {
                    ApplicationArea = All;
                    ToolTip = 'Friendly name shown in BC''s Email Accounts list (e.g. Sales Team).';
                }
                field("Mailbox Email"; Rec."Mailbox Email")
                {
                    ApplicationArea = All;
                    ToolTip = 'The email address or UPN of the shared mailbox. Used in the Graph API: /v1.0/users/{MailboxEmail}/sendMail.';
                }
                field("App Registration Code"; Rec."App Registration Code")
                {
                    ApplicationArea = All;
                    ToolTip = 'The App Registration that provides OAuth credentials. The registered app must have Mail.Send.Shared permission.';
                }
                field("Description"; Rec."Description")
                {
                    ApplicationArea = All;
                    ToolTip = 'Optional description for this shared mailbox account.';
                    MultiLine = true;
                }
            }
        }
    }

    actions
    {
        area(navigation)
        {
            action(AppRegistration)
            {
                ApplicationArea = All;
                Caption = 'App Registration';
                Image = Setup;
                ToolTip = 'Open the App Registration linked to this shared mailbox.';

                trigger OnAction()
                var
                    AppReg: Record "W365 App Registration";
                begin
                    if AppReg.Get(Rec."App Registration Code") then
                        Page.Run(Page::"W365 App Registration Card", AppReg);
                end;
            }
            action(AllMailboxes)
            {
                ApplicationArea = All;
                Caption = 'All Shared Mailboxes';
                Image = List;
                RunObject = Page "W365 Shared Mailbox Accounts";
                ToolTip = 'View all configured shared mailbox accounts.';
            }
        }
        area(Promoted)
        {
            actionref(AppRegistrationRef; AppRegistration) { }
            actionref(AllMailboxesRef; AllMailboxes) { }
        }
    }
}
