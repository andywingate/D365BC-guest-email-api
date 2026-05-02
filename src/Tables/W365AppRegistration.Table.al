namespace Wingate365.GuestEmailAPI;

table 50111 "W365 App Registration"
{
    Caption = 'App Registration';
    DataClassification = SystemMetadata;
    DrillDownPageId = "W365 App Registration Card";
    LookupPageId = "W365 App Registrations";

    fields
    {
        field(1; "Code"; Code[20])
        {
            Caption = 'Code';
            DataClassification = SystemMetadata;
            NotBlank = true;
        }
        field(2; "Description"; Text[100])
        {
            Caption = 'Description';
            DataClassification = SystemMetadata;
        }
        field(3; "App ID"; Guid)
        {
            Caption = 'App (Client) ID';
            DataClassification = SystemMetadata;
        }
        field(4; "Tenant ID"; Text[100])
        {
            Caption = 'Tenant ID';
            DataClassification = SystemMetadata;
            // Leave blank or set to 'common' for multi-tenant. Set to a specific tenant GUID for single-tenant.
        }
        field(5; "Domain Filter"; Text[250])
        {
            Caption = 'Domain Filter';
            DataClassification = SystemMetadata;
            // Optional. e.g. 'contoso.com'. Blank = matches all users. Used to route users from a specific home domain.
        }
        field(6; "Is Default"; Boolean)
        {
            Caption = 'Is Default';
            DataClassification = SystemMetadata;
            // Only one registration can be the default. The default is used when no domain filter matches.

            trigger OnValidate()
            var
                OtherReg: Record "W365 App Registration";
            begin
                if Rec."Is Default" then begin
                    OtherReg.LockTable();
                    OtherReg.SetRange("Is Default", true);
                    OtherReg.SetFilter("Code", '<>%1', Rec."Code");
                    OtherReg.ModifyAll("Is Default", false, true);
                end;
            end;
        }
    }

    keys
    {
        key(PK; "Code")
        {
            Clustered = true;
        }
        key(DomainFilter; "Domain Filter") { }
        key(IsDefault; "Is Default") { }
    }

    /// <summary>
    /// Returns the authority URL for this registration.
    /// Uses the Tenant ID if set; otherwise falls back to 'common' for multi-tenant.
    /// </summary>
    procedure GetAuthorityUrl(): Text
    var
        TenantId: Text;
        AuthBase: Label 'https://login.microsoftonline.com/', Locked = true;
        OAuthPath: Label '/oauth2/v2.0', Locked = true;
    begin
        TenantId := Rec."Tenant ID";
        if TenantId = '' then
            TenantId := 'common';
        exit(AuthBase + TenantId + OAuthPath);
    end;

    /// <summary>
    /// Returns the IsolatedStorage key for this registration's client secret.
    /// </summary>
    procedure GetClientSecretKey(): Text
    var
        AppIdText: Text;
    begin
        AppIdText := Format(Rec."App ID").Replace('{', '').Replace('}', '').Replace('-', '');
        exit('W365_CS_' + AppIdText);
    end;
}
