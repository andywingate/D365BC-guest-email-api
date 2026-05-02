namespace Wingate365.GuestEmailAPI;

using System.Security.AccessControl;

table 50101 "W365 User Email Token"
{
    Caption = 'User Email Cache';
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
        field(5; "Home Email"; Text[250])
        {
            Caption = 'Home Email';
            DataClassification = EndUserIdentifiableInformation;
        }
    }

    keys
    {
        key(PK; "User Name")
        {
            Clustered = true;
        }
    }
}
