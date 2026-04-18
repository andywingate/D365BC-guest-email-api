namespace Wingate365.GuestEmailAPI;

page 50103 "W365 Email Setup Card"
{
    Caption = 'W365 Email Setup';
    PageType = Card;
    SourceTable = "W365 Email Setup";
    UsageCategory = Administration;
    ApplicationArea = All;
    InsertAllowed = false;
    DeleteAllowed = false;

    layout
    {
        area(content)
        {
            group(AppRegistration)
            {
                Caption = 'Entra App Registration';

                field("App ID"; Rec."App ID")
                {
                    ApplicationArea = All;
                    ToolTip = 'The Application (Client) ID of the Azure App Registration in your host tenant.';
                }
                field("Tenant ID"; Rec."Tenant ID")
                {
                    ApplicationArea = All;
                    ToolTip = 'The Directory (Tenant) ID of your host Azure AD tenant.';
                }
                field("Redirect URI"; Rec."Redirect URI")
                {
                    ApplicationArea = All;
                    ToolTip = 'The redirect URI registered on the Entra app. Recommended: https://login.microsoftonline.com/common/oauth2/nativeclient';
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
                        SecretSavedMsg: Label 'Client secret saved.';
                    begin
                        if ClientSecretText <> '' then begin
                            OAuthMgt.SetClientSecret(ClientSecretText);
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
                ToolTip = 'Removes the stored client secret. You will need to set a new one before the consent flow can run.';

                trigger OnAction()
                var
                    ConfirmMsg: Label 'Are you sure you want to clear the stored client secret?';
                begin
                    if Confirm(ConfirmMsg) then
                        IsolatedStorage.Delete('W365_CS', DataScope::Company);
                end;
            }
            action(UserTokens)
            {
                ApplicationArea = All;
                Caption = 'User Tokens';
                Image = Users;
                RunObject = Page "W365 User Token List";
                ToolTip = 'View the OAuth token status for all users and trigger the consent flow.';
            }
        }
        area(Promoted)
        {
            actionref(UserTokensRef; UserTokens) { }
        }
    }

    trigger OnOpenPage()
    begin
        Rec.GetOrInit();
        if not Rec.Insert() then;
    end;

    var
        ClientSecretText: Text[250];

    local procedure GetClientSecretStatus(): Text
    var
        OAuthMgt: Codeunit "W365 OAuth Mgt";
        ConfiguredTxt: Label 'Configured';
        NotConfiguredTxt: Label 'Not configured';
    begin
        if OAuthMgt.HasClientSecret() then
            exit(ConfiguredTxt)
        else
            exit(NotConfiguredTxt);
    end;
}
