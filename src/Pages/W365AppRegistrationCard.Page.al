namespace Wingate365.GuestEmailAPI;

page 50113 "W365 App Registration Card"
{
    Caption = 'App Registration';
    PageType = Card;
    SourceTable = "W365 App Registration";
    UsageCategory = Administration;
    ApplicationArea = All;

    layout
    {
        area(content)
        {
            group(RegistrationDetails)
            {
                Caption = 'Registration Details';

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
                    ToolTip = 'The Application (Client) ID from the Entra app registration in Azure.';
                }
                field("Tenant ID"; Rec."Tenant ID")
                {
                    ApplicationArea = All;
                    ToolTip = 'The Tenant ID. Leave blank or use ''common'' for multi-tenant. Enter the tenant GUID for single-tenant.';
                }
                field("Domain Filter"; Rec."Domain Filter")
                {
                    ApplicationArea = All;
                    ToolTip = 'Optional. Restricts this registration to users from a specific home domain (e.g. contoso.com). Leave blank to use as the default fallback.';
                }
                field("Is Default"; Rec."Is Default")
                {
                    ApplicationArea = All;
                    ToolTip = 'Mark this registration as the default fallback when no domain filter matches.';
                }
            }
            group(ClientSecretGroup)
            {
                Caption = 'Client Secret';

                field(ClientSecretStatus; GetClientSecretStatus())
                {
                    ApplicationArea = All;
                    Caption = 'Client Secret Status';
                    Editable = false;
                    ToolTip = 'Indicates whether a client secret has been stored. Use the Set Client Secret action to configure it.';
                }
                field(ClientSecretInput; ClientSecretText)
                {
                    ApplicationArea = All;
                    Caption = 'Enter New Client Secret';
                    ExtendedDatatype = Masked;
                    ToolTip = 'Type the client secret value here and press Tab or Enter. The value is stored encrypted and cannot be read back.';

                    trigger OnValidate()
                    var
                        OAuthMgt: Codeunit "W365 OAuth Mgt";
                        ClientSecretSecret: SecretText;
                        SecretSavedMsg: Label 'Client secret saved.';
                        NoAppIdErr: Label 'Enter an App (Client) ID before setting the client secret.';
                    begin
                        if ClientSecretText <> '' then begin
                            if IsNullGuid(Rec."App ID") then
                                Error(NoAppIdErr);
                            ClientSecretSecret := ClientSecretText;
                            OAuthMgt.SetClientSecret(Rec."App ID", ClientSecretSecret);
                            ClientSecretText := '';
                            Message(SecretSavedMsg);
                        end;
                    end;
                }
            }
        }
    }

    actions
    {
        area(processing)
        {
            action(ClearClientSecret)
            {
                ApplicationArea = All;
                Caption = 'Clear Client Secret';
                Image = Delete;
                ToolTip = 'Removes the stored client secret. You will need to set a new one before users can authenticate.';

                trigger OnAction()
                var
                    OAuthMgt: Codeunit "W365 OAuth Mgt";
                    ConfirmMsg: Label 'Are you sure you want to clear the stored client secret for this registration?';
                begin
                    if Confirm(ConfirmMsg) then
                        OAuthMgt.ClearClientSecret(Rec."App ID");
                end;
            }
        }
        area(navigation)
        {
            action(AllRegistrations)
            {
                ApplicationArea = All;
                Caption = 'All Registrations';
                Image = List;
                RunObject = Page "W365 App Registrations";
                ToolTip = 'View all configured app registrations.';
            }
        }
        area(Promoted)
        {
            actionref(ClearSecretRef; ClearClientSecret) { }
            actionref(AllRegistrationsRef; AllRegistrations) { }
        }
    }

    var
        ClientSecretText: Text[250];

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
