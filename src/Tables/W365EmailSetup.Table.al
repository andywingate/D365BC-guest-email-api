namespace Wingate365.GuestEmailAPI;

table 50100 "W365 Email Setup"
{
    Caption = 'W365 Email Setup';
    DataClassification = SystemMetadata;
    DrillDownPageId = "W365 Email Setup Card";
    LookupPageId = "W365 Email Setup Card";

    fields
    {
        field(1; "Primary Key"; Code[10])
        {
            Caption = 'Primary Key';
            DataClassification = SystemMetadata;
        }
        field(2; "App ID"; Text[100])
        {
            Caption = 'App (Client) ID';
            DataClassification = SystemMetadata;
        }
        field(3; "Tenant ID"; Text[100])
        {
            Caption = 'Host Tenant ID';
            DataClassification = SystemMetadata;
        }
        field(4; "Redirect URI"; Text[250])
        {
            Caption = 'Redirect URI';
            DataClassification = SystemMetadata;
        }
    }

    keys
    {
        key(PK; "Primary Key")
        {
            Clustered = true;
        }
    }

    procedure GetOrError()
    var
        SetupMissingErr: Label 'W365 Email Setup has not been configured. Open the W365 Email Setup Card and enter the Entra app registration details.';
    begin
        if not Rec.Get('') then
            Error(SetupMissingErr);
    end;

    procedure GetOrInit()
    begin
        if not Rec.Get('') then begin
            Rec.Init();
            Rec."Primary Key" := '';
        end;
    end;
}
