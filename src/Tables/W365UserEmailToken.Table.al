namespace Wingate365.GuestEmailAPI;

using System.Security.AccessControl;

table 50101 "W365 User Email Token"
{
    Caption = 'W365 User Email Token';
    DataClassification = EndUserIdentifiableInformation;
    DrillDownPageId = "W365 User Token List";
    LookupPageId = "W365 User Token List";

    fields
    {
        field(1; "User Name"; Code[50])
        {
            Caption = 'User Name';
            DataClassification = EndUserIdentifiableInformation;
            TableRelation = User."User Name";
            ValidateTableRelation = false;
        }
        field(2; "Token Expiry"; DateTime)
        {
            Caption = 'Token Expiry';
            DataClassification = SystemMetadata;
        }
        field(3; "Consent Status"; Enum "W365 Consent Status")
        {
            Caption = 'Consent Status';
            DataClassification = SystemMetadata;
        }
        field(4; "Last Error"; Text[250])
        {
            Caption = 'Last Error';
            DataClassification = SystemMetadata;
        }
    }

    keys
    {
        key(PK; "User Name")
        {
            Clustered = true;
        }
    }

    procedure IsTokenExpired(): Boolean
    begin
        if Rec."Token Expiry" = 0DT then
            exit(true);
        exit(CurrentDateTime() >= Rec."Token Expiry" - 60000); // 60-second buffer
    end;
}
