namespace Wingate365.GuestEmailAPI;

table 50115 "W365 Shared Mailbox Account"
{
    Caption = 'Shared Mailbox Account';
    DataClassification = SystemMetadata;
    DrillDownPageId = "W365 Shared Mailbox Card";
    LookupPageId = "W365 Shared Mailbox Accounts";

    fields
    {
        field(1; "Code"; Code[20])
        {
            Caption = 'Code';
            DataClassification = SystemMetadata;
            NotBlank = true;
        }
        field(2; "Display Name"; Text[100])
        {
            Caption = 'Display Name';
            DataClassification = SystemMetadata;
        }
        field(3; "Mailbox Email"; Text[250])
        {
            Caption = 'Mailbox Email / UPN';
            DataClassification = EndUserIdentifiableInformation;
            NotBlank = true;
            // The email address or UPN of the shared mailbox.
            // Used in the Graph endpoint: /v1.0/users/{MailboxEmail}/sendMail
        }
        field(4; "App Registration Code"; Code[20])
        {
            Caption = 'App Registration';
            DataClassification = SystemMetadata;
            NotBlank = true;
            TableRelation = "W365 App Registration"."Code";
        }
        field(5; "Description"; Text[250])
        {
            Caption = 'Description';
            DataClassification = SystemMetadata;
        }
    }

    keys
    {
        key(PK; "Code")
        {
            Clustered = true;
        }
    }
}
