namespace Wingate365.GuestEmailAPI;

page 50116 "W365 Shared Mailbox Accounts"
{
    Caption = 'Shared Mailbox Accounts';
    PageType = List;
    SourceTable = "W365 Shared Mailbox Account";
    UsageCategory = Administration;
    ApplicationArea = All;
    CardPageId = "W365 Shared Mailbox Card";
    InsertAllowed = true;
    DeleteAllowed = true;
    ModifyAllowed = true;

    layout
    {
        area(content)
        {
            repeater(MailboxList)
            {
                field("Code"; Rec."Code")
                {
                    ApplicationArea = All;
                    ToolTip = 'Short code identifying this shared mailbox account.';
                }
                field("Display Name"; Rec."Display Name")
                {
                    ApplicationArea = All;
                    ToolTip = 'Friendly name shown in BC''s Email Accounts list (e.g. Sales Team).';
                }
                field("Mailbox Email"; Rec."Mailbox Email")
                {
                    ApplicationArea = All;
                    ToolTip = 'The email address or UPN of the shared mailbox used in the Graph API call.';
                }
                field("App Registration Code"; Rec."App Registration Code")
                {
                    ApplicationArea = All;
                    ToolTip = 'The App Registration that provides the OAuth credentials for sending from this mailbox.';
                }
                field("Description"; Rec."Description")
                {
                    ApplicationArea = All;
                    ToolTip = 'Optional description for this shared mailbox account.';
                }
            }
        }
    }

    actions
    {
        area(processing)
        {
            action(OpenCard)
            {
                ApplicationArea = All;
                Caption = 'Edit';
                Image = Edit;
                RunObject = Page "W365 Shared Mailbox Card";
                RunPageOnRec = true;
                ToolTip = 'Open the shared mailbox card to edit its details.';
            }
        }
        area(Promoted)
        {
            actionref(OpenCardRef; OpenCard) { }
        }
    }
}
