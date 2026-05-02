namespace Wingate365.GuestEmailAPI;

page 50104 "W365 User Token List"
{
    Caption = 'User Email Cache';
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
            repeater(UserList)
            {
                field("User Name"; Rec."User Name")
                {
                    ApplicationArea = All;
                    ToolTip = 'The BC user name.';
                }
                field("Home Email"; Rec."Home Email")
                {
                    ApplicationArea = All;
                    ToolTip = 'The user''s home email address cached from Microsoft Graph after their first authentication.';
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
            action(ClearHomeEmail)
            {
                ApplicationArea = All;
                Caption = 'Clear Home Email';
                Image = Delete;
                ToolTip = 'Clears the cached home email for the selected user. It will be refreshed automatically on the next successful authentication.';

                trigger OnAction()
                var
                    ConfirmMsg: Label 'Clear the cached home email for this user? It will be refreshed on next authentication.';
                begin
                    if Confirm(ConfirmMsg) then begin
                        Rec."Home Email" := '';
                        Rec.Modify();
                        CurrPage.Update(false);
                    end;
                end;
            }
        }
        area(navigation)
        {
            action(AppRegistrations)
            {
                ApplicationArea = All;
                Caption = 'App Registrations';
                Image = Setup;
                RunObject = Page "W365 App Registrations";
                ToolTip = 'Open the App Registrations page to configure or review Entra app registrations.';
            }
        }
        area(Promoted)
        {
            actionref(AppRegistrationsRef; AppRegistrations) { }
        }
    }
}
