namespace Wingate365.GuestEmailAPI;

page 50112 "W365 App Registrations"
{
    Caption = 'App Registrations';
    PageType = List;
    SourceTable = "W365 App Registration";
    UsageCategory = Administration;
    ApplicationArea = All;
    CardPageId = "W365 App Registration Card";
    InsertAllowed = true;
    DeleteAllowed = true;
    ModifyAllowed = true;

    layout
    {
        area(content)
        {
            repeater(RegistrationList)
            {
                field("Code"; Rec."Code")
                {
                    ApplicationArea = All;
                    ToolTip = 'Short code identifying this app registration (e.g. CONTOSO).';
                }
                field("Description"; Rec."Description")
                {
                    ApplicationArea = All;
                    ToolTip = 'Friendly description for this registration.';
                }
                field("App ID"; Rec."App ID")
                {
                    ApplicationArea = All;
                    ToolTip = 'The Application (Client) ID from the Entra app registration.';
                }
                field("Tenant ID"; Rec."Tenant ID")
                {
                    ApplicationArea = All;
                    ToolTip = 'The Tenant ID. Leave blank or use ''common'' for multi-tenant. Enter the tenant GUID for single-tenant.';
                }
                field("Domain Filter"; Rec."Domain Filter")
                {
                    ApplicationArea = All;
                    ToolTip = 'Optional. Restricts this registration to users whose home domain matches (e.g. contoso.com). Leave blank to use as default.';
                }
                field("Is Default"; Rec."Is Default")
                {
                    ApplicationArea = All;
                    ToolTip = 'Mark this registration as the default. Used when no domain filter matches a user''s home domain.';
                }
                field(ClientSecretStatus; GetClientSecretStatus())
                {
                    ApplicationArea = All;
                    Caption = 'Client Secret';
                    Editable = false;
                    ToolTip = 'Indicates whether a client secret has been stored for this registration.';
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
                RunObject = Page "W365 App Registration Card";
                RunPageOnRec = true;
                ToolTip = 'Open the registration card to set or change the client secret.';
            }
        }
        area(Promoted)
        {
            actionref(OpenCardRef; OpenCard) { }
        }
    }

    local procedure GetClientSecretStatus(): Text
    var
        OAuthMgt: Codeunit "W365 OAuth Mgt";
        ConfiguredTxt: Label 'Configured';
        NotConfiguredTxt: Label 'Not configured';
    begin
        if OAuthMgt.HasClientSecret(Rec."App ID") then
            exit(ConfiguredTxt)
        else
            exit(NotConfiguredTxt);
    end;
}
