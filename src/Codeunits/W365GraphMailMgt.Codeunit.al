namespace Wingate365.GuestEmailAPI;

using System.Email;
using System.Utilities;

codeunit 50105 "W365 Graph Mail Mgt"
{
    Access = Internal;

    // -------------------------------------------------------------------------
    // Public send API
    // -------------------------------------------------------------------------

    /// <summary>
    /// Sends an email via Microsoft Graph using the provided access token.
    /// Handles all To/Cc/Bcc recipients and attachments from the Email Message.
    /// Chooses inline base64 for small attachments and upload session for large ones.
    /// </summary>
    [NonDebuggable]
    procedure SendEmail(EmailMessage: Codeunit "Email Message"; AccessToken: SecretText): Boolean
    begin
        if GetTotalAttachmentSize(EmailMessage) < 4 * 1024 * 1024 then
            exit(SendMailInline(EmailMessage, AccessToken))
        else
            exit(SendMailViaUploadSession(EmailMessage, AccessToken));
    end;

    /// <summary>
    /// Sends an email via Graph for a shared mailbox.
    /// Uses /v1.0/users/{mailboxEmail}/sendMail instead of /v1.0/me/sendMail.
    /// </summary>
    [NonDebuggable]
    procedure SendEmailAsSharedMailbox(EmailMessage: Codeunit "Email Message"; AccessToken: SecretText; MailboxEmail: Text): Boolean
    begin
        if GetTotalAttachmentSize(EmailMessage) < 4 * 1024 * 1024 then
            exit(SendMailInlineAsMailbox(EmailMessage, AccessToken, MailboxEmail))
        else
            exit(SendMailViaUploadSessionAsMailbox(EmailMessage, AccessToken, MailboxEmail));
    end;

    /// <summary>
    /// Calls Graph GET /me to retrieve the authenticated user's email address.
    /// Used after first auth to populate the home email cache.
    /// Returns empty string on any failure - non-fatal.
    /// </summary>
    [NonDebuggable]
    procedure GetCurrentUserEmail(AccessToken: SecretText): Text
    var
        HttpClient: HttpClient;
        HttpReqMsg: HttpRequestMessage;
        HttpRespMsg: HttpResponseMessage;
        ReqHeaders: HttpHeaders;
        ResponseText: Text;
        JsonObj: JsonObject;
        JsonToken: JsonToken;
        MeEndpoint: Label 'https://graph.microsoft.com/v1.0/me?$select=mail,userPrincipalName', Locked = true;
    begin
        if AccessToken.IsEmpty() then
            exit('');

        HttpReqMsg.Method := 'GET';
        HttpReqMsg.SetRequestUri(MeEndpoint);
        HttpReqMsg.GetHeaders(ReqHeaders);
        ReqHeaders.Add('Authorization', SecretStrSubstNo('Bearer %1', AccessToken));
        ReqHeaders.Add('Accept', 'application/json');

        if not HttpClient.Send(HttpReqMsg, HttpRespMsg) then
            exit('');

        if not HttpRespMsg.IsSuccessStatusCode() then
            exit('');

        HttpRespMsg.Content.ReadAs(ResponseText);
        if not JsonObj.ReadFrom(ResponseText) then
            exit('');

        // Prefer 'mail' (SMTP address); fall back to 'userPrincipalName'
        if JsonObj.Get('mail', JsonToken) then
            if JsonToken.AsValue().AsText() <> '' then
                exit(JsonToken.AsValue().AsText());

        if JsonObj.Get('userPrincipalName', JsonToken) then
            exit(JsonToken.AsValue().AsText());

        exit('');
    end;

    // -------------------------------------------------------------------------
    // Inline send (total attachments < 4 MB)
    // -------------------------------------------------------------------------

    [NonDebuggable]
    local procedure SendMailInline(EmailMessage: Codeunit "Email Message"; AccessToken: SecretText): Boolean
    var
        GraphEndpoint: Label 'https://graph.microsoft.com/v1.0/me/sendMail', Locked = true;
    begin
        exit(ExecuteInlineSend(EmailMessage, AccessToken, GraphEndpoint));
    end;

    [NonDebuggable]
    local procedure SendMailInlineAsMailbox(EmailMessage: Codeunit "Email Message"; AccessToken: SecretText; MailboxEmail: Text): Boolean
    var
        GraphEndpointTpl: Label 'https://graph.microsoft.com/v1.0/users/%1/sendMail', Locked = true;
    begin
        exit(ExecuteInlineSend(EmailMessage, AccessToken, StrSubstNo(GraphEndpointTpl, MailboxEmail)));
    end;

    [NonDebuggable]
    local procedure ExecuteInlineSend(EmailMessage: Codeunit "Email Message"; AccessToken: SecretText; GraphEndpoint: Text): Boolean
    var
        HttpClient: HttpClient;
        HttpReqMsg: HttpRequestMessage;
        HttpRespMsg: HttpResponseMessage;
        HttpContent: HttpContent;
        ReqHeaders: HttpHeaders;
        ContentHeaders: HttpHeaders;
        JsonBody: Text;
        ConnectErr: Label 'Could not reach Microsoft Graph. Check BC server outbound connectivity.';
        ThrottledErr: Label 'Microsoft Graph is throttling requests. Please wait a moment and try again.';
        StatusCode: Integer;
        ResponseText: Text;
    begin
        JsonBody := BuildSendMailJson(EmailMessage, true);

        HttpContent.WriteFrom(JsonBody);
        HttpContent.GetHeaders(ContentHeaders);
        ContentHeaders.Remove('Content-Type');
        ContentHeaders.Add('Content-Type', 'application/json');

        HttpReqMsg.Method := 'POST';
        HttpReqMsg.SetRequestUri(GraphEndpoint);
        HttpReqMsg.Content := HttpContent;
        HttpReqMsg.GetHeaders(ReqHeaders);
        ReqHeaders.Add('Authorization', SecretStrSubstNo('Bearer %1', AccessToken));

        if not HttpClient.Send(HttpReqMsg, HttpRespMsg) then
            Error(ConnectErr);

        StatusCode := HttpRespMsg.HttpStatusCode();

        if StatusCode = 202 then
            exit(true);

        if StatusCode = 401 then begin
            ClearSessionToken();
            Error('Microsoft Graph rejected the authorisation token (401). Re-authentication will occur on the next send.');
        end;

        if StatusCode = 429 then
            Error(ThrottledErr);

        HttpRespMsg.Content.ReadAs(ResponseText);
        ParseAndRaiseGraphError(ResponseText, StatusCode);
        exit(false);
    end;

    // -------------------------------------------------------------------------
    // Upload session send (total attachments >= 4 MB)
    // -------------------------------------------------------------------------

    [NonDebuggable]
    local procedure SendMailViaUploadSession(EmailMessage: Codeunit "Email Message"; AccessToken: SecretText): Boolean
    var
        MessageEndpoint: Label 'https://graph.microsoft.com/v1.0/me/messages', Locked = true;
        SendEndpointTpl: Label 'https://graph.microsoft.com/v1.0/me/messages/%1/send', Locked = true;
    begin
        exit(ExecuteUploadSessionSend(EmailMessage, AccessToken, MessageEndpoint, SendEndpointTpl));
    end;

    [NonDebuggable]
    local procedure SendMailViaUploadSessionAsMailbox(EmailMessage: Codeunit "Email Message"; AccessToken: SecretText; MailboxEmail: Text): Boolean
    var
        MessageEndpointTpl: Label 'https://graph.microsoft.com/v1.0/users/%1/messages', Locked = true;
        SendEndpointTpl: Label 'https://graph.microsoft.com/v1.0/users/%1/messages/%2/send', Locked = true;
        MessageEndpoint: Text;
        SendEndpoint: Text;
    begin
        MessageEndpoint := StrSubstNo(MessageEndpointTpl, MailboxEmail);
        SendEndpoint := StrSubstNo(SendEndpointTpl, MailboxEmail, '%1');
        exit(ExecuteUploadSessionSend(EmailMessage, AccessToken, MessageEndpoint, SendEndpoint));
    end;

    [NonDebuggable]
    local procedure ExecuteUploadSessionSend(EmailMessage: Codeunit "Email Message"; AccessToken: SecretText; MessageEndpoint: Text; SendEndpointTemplate: Text): Boolean
    var
        MessageId: Text;
        SendEndpoint: Text;
        AuthHeader: Text;
    begin
        AuthHeader := SecretStrSubstNo('Bearer %1', AccessToken);

        // Step 1: create a draft message (no attachments in body)
        MessageId := CreateDraftMessage(EmailMessage, AccessToken, MessageEndpoint, AuthHeader);
        if MessageId = '' then
            exit(false);

        // Step 2: upload each attachment via upload session
        if not UploadAllAttachments(EmailMessage, AccessToken, MessageEndpoint, MessageId, AuthHeader) then
            exit(false);

        // Step 3: send the draft
        SendEndpoint := StrSubstNo(SendEndpointTemplate, MessageId);
        exit(SendDraftMessage(AccessToken, SendEndpoint, AuthHeader));
    end;

    [NonDebuggable]
    local procedure CreateDraftMessage(EmailMessage: Codeunit "Email Message"; AccessToken: SecretText; MessageEndpoint: Text; AuthHeader: Text): Text
    var
        HttpClient: HttpClient;
        HttpReqMsg: HttpRequestMessage;
        HttpRespMsg: HttpResponseMessage;
        HttpContent: HttpContent;
        ReqHeaders: HttpHeaders;
        ContentHeaders: HttpHeaders;
        JsonBody: Text;
        ResponseText: Text;
        JsonObj: JsonObject;
        JsonToken: JsonToken;
        MsgToken: JsonToken;
        ConnectErr: Label 'Could not reach Microsoft Graph when creating draft message. Check BC server outbound connectivity.';
        StatusCode: Integer;
    begin
        // Build message JSON without attachments (draft format - no wrapper)
        JsonBody := BuildDraftMessageJson(EmailMessage);

        HttpContent.WriteFrom(JsonBody);
        HttpContent.GetHeaders(ContentHeaders);
        ContentHeaders.Remove('Content-Type');
        ContentHeaders.Add('Content-Type', 'application/json');

        HttpReqMsg.Method := 'POST';
        HttpReqMsg.SetRequestUri(MessageEndpoint);
        HttpReqMsg.Content := HttpContent;
        HttpReqMsg.GetHeaders(ReqHeaders);
        ReqHeaders.Add('Authorization', AuthHeader);

        if not HttpClient.Send(HttpReqMsg, HttpRespMsg) then
            Error(ConnectErr);

        StatusCode := HttpRespMsg.HttpStatusCode();
        HttpRespMsg.Content.ReadAs(ResponseText);

        if StatusCode <> 201 then begin
            ParseAndRaiseGraphError(ResponseText, StatusCode);
            exit('');
        end;

        if not JsonObj.ReadFrom(ResponseText) then
            exit('');

        if JsonObj.Get('id', JsonToken) then
            exit(JsonToken.AsValue().AsText());

        // Draft messages may wrap the ID in a 'message' object
        if JsonObj.Get('message', MsgToken) then begin
            JsonObj := MsgToken.AsObject();
            if JsonObj.Get('id', JsonToken) then
                exit(JsonToken.AsValue().AsText());
        end;

        exit('');
    end;

    [NonDebuggable]
    local procedure UploadAllAttachments(EmailMessage: Codeunit "Email Message"; AccessToken: SecretText; MessageEndpoint: Text; MessageId: Text; AuthHeader: Text): Boolean
    var
        Base64Convert: Codeunit "Base64 Convert";
        TempBlob: Codeunit "Temp Blob";
        AttachOutStr: OutStream;
        AttachInStr: InStream;
        AttachName: Text;
        AttachContentType: Text;
        AttachBase64: Text;
        AttachSize: Integer;
        UploadUrl: Text;
    begin
        if not EmailMessage.Attachments_First() then
            exit(true); // No attachments - nothing to do

        repeat
            AttachName := EmailMessage.Attachments_GetName();
            AttachContentType := EmailMessage.Attachments_GetContentType();
            AttachBase64 := EmailMessage.Attachments_GetContentBase64();
            AttachSize := EmailMessage.Attachments_GetLength();

            // Decode base64 to binary in a TempBlob
            TempBlob.CreateOutStream(AttachOutStr);
            Base64Convert.FromBase64(AttachBase64, AttachOutStr);
            TempBlob.CreateInStream(AttachInStr);

            // Create upload session for this attachment
            UploadUrl := CreateAttachmentUploadSession(AccessToken, MessageEndpoint, MessageId, AuthHeader, AttachName, AttachContentType, AttachSize);
            if UploadUrl = '' then
                exit(false);

            // Upload binary content in a single PUT (up to 60 MB per Graph spec)
            if not PutAttachmentChunk(UploadUrl, AttachInStr, AttachSize) then
                exit(false);

        until not EmailMessage.Attachments_Next();

        exit(true);
    end;

    [NonDebuggable]
    local procedure CreateAttachmentUploadSession(AccessToken: SecretText; MessageEndpoint: Text; MessageId: Text; AuthHeader: Text; AttachName: Text; AttachContentType: Text; AttachSize: Integer): Text
    var
        HttpClient: HttpClient;
        HttpReqMsg: HttpRequestMessage;
        HttpRespMsg: HttpResponseMessage;
        HttpContent: HttpContent;
        ReqHeaders: HttpHeaders;
        ContentHeaders: HttpHeaders;
        UploadSessionEndpointTpl: Label '%1/%2/attachments/createUploadSession', Locked = true;
        UploadSessionEndpoint: Text;
        RequestJson: Text;
        ResponseText: Text;
        JsonObj: JsonObject;
        JsonToken: JsonToken;
        StatusCode: Integer;
        ConnectErr: Label 'Could not reach Microsoft Graph when creating attachment upload session.';
    begin
        UploadSessionEndpoint := StrSubstNo(UploadSessionEndpointTpl, MessageEndpoint, MessageId);

        RequestJson := '{"AttachmentItem":{"attachmentType":"file","name":"' +
            EscapeJsonString(AttachName) + '","size":' + Format(AttachSize) +
            ',"contentType":"' + EscapeJsonString(AttachContentType) + '"}}';

        HttpContent.WriteFrom(RequestJson);
        HttpContent.GetHeaders(ContentHeaders);
        ContentHeaders.Remove('Content-Type');
        ContentHeaders.Add('Content-Type', 'application/json');

        HttpReqMsg.Method := 'POST';
        HttpReqMsg.SetRequestUri(UploadSessionEndpoint);
        HttpReqMsg.Content := HttpContent;
        HttpReqMsg.GetHeaders(ReqHeaders);
        ReqHeaders.Add('Authorization', AuthHeader);

        if not HttpClient.Send(HttpReqMsg, HttpRespMsg) then
            Error(ConnectErr);

        StatusCode := HttpRespMsg.HttpStatusCode();
        HttpRespMsg.Content.ReadAs(ResponseText);

        if StatusCode <> 200 then begin
            ParseAndRaiseGraphError(ResponseText, StatusCode);
            exit('');
        end;

        if not JsonObj.ReadFrom(ResponseText) then
            exit('');

        if JsonObj.Get('uploadUrl', JsonToken) then
            exit(JsonToken.AsValue().AsText());

        exit('');
    end;

    local procedure PutAttachmentChunk(UploadUrl: Text; AttachInStr: InStream; AttachSize: Integer): Boolean
    var
        HttpClient: HttpClient;
        HttpReqMsg: HttpRequestMessage;
        HttpRespMsg: HttpResponseMessage;
        HttpContent: HttpContent;
        ContentHeaders: HttpHeaders;
        ReqHeaders: HttpHeaders;
        ContentRangeHeader: Text;
        StatusCode: Integer;
        ResponseText: Text;
        ConnectErr: Label 'Could not reach Microsoft Graph when uploading attachment chunk.';
    begin
        // Upload the entire attachment in one PUT (Graph allows up to 60 MB per PUT)
        HttpContent.WriteFrom(AttachInStr);
        HttpContent.GetHeaders(ContentHeaders);
        ContentHeaders.Remove('Content-Type');
        ContentHeaders.Add('Content-Type', 'application/octet-stream');

        ContentRangeHeader := StrSubstNo('bytes 0-%1/%2', AttachSize - 1, AttachSize);

        HttpReqMsg.Method := 'PUT';
        HttpReqMsg.SetRequestUri(UploadUrl);
        HttpReqMsg.Content := HttpContent;
        HttpReqMsg.GetHeaders(ReqHeaders);
        ReqHeaders.Add('Content-Range', ContentRangeHeader);

        if not HttpClient.Send(HttpReqMsg, HttpRespMsg) then
            Error(ConnectErr);

        StatusCode := HttpRespMsg.HttpStatusCode();

        // 200 = complete, 201 = created, 202 = accepted (partial - should not occur for single-PUT)
        if StatusCode in [200, 201, 202] then
            exit(true);

        HttpRespMsg.Content.ReadAs(ResponseText);
        ParseAndRaiseGraphError(ResponseText, StatusCode);
        exit(false);
    end;

    [NonDebuggable]
    local procedure SendDraftMessage(AccessToken: SecretText; SendEndpoint: Text; AuthHeader: Text): Boolean
    var
        HttpClient: HttpClient;
        HttpReqMsg: HttpRequestMessage;
        HttpRespMsg: HttpResponseMessage;
        HttpContent: HttpContent;
        ContentHeaders: HttpHeaders;
        ReqHeaders: HttpHeaders;
        StatusCode: Integer;
        ResponseText: Text;
        ConnectErr: Label 'Could not reach Microsoft Graph when sending draft message.';
        ThrottledErr: Label 'Microsoft Graph is throttling requests. Please wait a moment and try again.';
    begin
        // POST to /send requires an empty body with Content-Length: 0
        HttpContent.WriteFrom('');
        HttpContent.GetHeaders(ContentHeaders);
        ContentHeaders.Remove('Content-Type');
        ContentHeaders.Add('Content-Type', 'application/json');

        HttpReqMsg.Method := 'POST';
        HttpReqMsg.SetRequestUri(SendEndpoint);
        HttpReqMsg.Content := HttpContent;
        HttpReqMsg.GetHeaders(ReqHeaders);
        ReqHeaders.Add('Authorization', AuthHeader);

        if not HttpClient.Send(HttpReqMsg, HttpRespMsg) then
            Error(ConnectErr);

        StatusCode := HttpRespMsg.HttpStatusCode();

        if StatusCode = 202 then
            exit(true);

        if StatusCode = 429 then
            Error(ThrottledErr);

        HttpRespMsg.Content.ReadAs(ResponseText);
        ParseAndRaiseGraphError(ResponseText, StatusCode);
        exit(false);
    end;

    // -------------------------------------------------------------------------
    // JSON builders
    // -------------------------------------------------------------------------

    /// <summary>
    /// Builds the sendMail / message JSON for the given Email Message.
    /// When includeAttachments = true, embeds attachments as inline base64 fileAttachments.
    /// When false, builds a plain message body without attachments (used for draft + upload).
    /// Includes all To, Cc, and Bcc recipients.
    /// When wrapForSendMail = true, wraps in { "message": {...}, "saveToSentItems": true }.
    /// When false, returns the message object directly (for POST /messages draft creation).
    /// </summary>
    local procedure BuildSendMailJson(EmailMessage: Codeunit "Email Message"; IncludeAttachments: Boolean): Text
    var
        MsgObj: JsonObject;
        BodyObj: JsonObject;
        WrapperObj: JsonObject;
        ToRecipientsArr: JsonArray;
        CcRecipientsArr: JsonArray;
        BccRecipientsArr: JsonArray;
        AttachmentsArr: JsonArray;
        ToList: List of [Text];
        CcList: List of [Text];
        BccList: List of [Text];
        Subject: Text;
        Body: Text;
        HasAttachments: Boolean;
        Result: Text;
    begin
        Subject := EmailMessage.GetSubject();
        Body := EmailMessage.GetBody();

        BodyObj.Add('contentType', 'HTML');
        BodyObj.Add('content', Body);

        EmailMessage.GetRecipients(Enum::"Email Recipient Type"::"To", ToList);
        EmailMessage.GetRecipients(Enum::"Email Recipient Type"::"Cc", CcList);
        EmailMessage.GetRecipients(Enum::"Email Recipient Type"::"Bcc", BccList);

        ToRecipientsArr := BuildRecipientArray(ToList);
        CcRecipientsArr := BuildRecipientArray(CcList);
        BccRecipientsArr := BuildRecipientArray(BccList);

        MsgObj.Add('subject', Subject);
        MsgObj.Add('body', BodyObj);
        MsgObj.Add('toRecipients', ToRecipientsArr);
        if CcList.Count() > 0 then
            MsgObj.Add('ccRecipients', CcRecipientsArr);
        if BccList.Count() > 0 then
            MsgObj.Add('bccRecipients', BccRecipientsArr);

        if IncludeAttachments then begin
            AttachmentsArr := BuildAttachmentsArray(EmailMessage, HasAttachments);
            if HasAttachments then
                MsgObj.Add('attachments', AttachmentsArr);
        end;

        // sendMail endpoint: { "message": {...}, "saveToSentItems": true }
        WrapperObj.Add('message', MsgObj);
        WrapperObj.Add('saveToSentItems', true);
        WrapperObj.WriteTo(Result);
        exit(Result);
    end;

    /// <summary>
    /// Builds the message JSON for draft creation (POST /messages).
    /// Returns the message object directly - no sendMail wrapper.
    /// </summary>
    local procedure BuildDraftMessageJson(EmailMessage: Codeunit "Email Message"): Text
    var
        MsgObj: JsonObject;
        BodyObj: JsonObject;
        ToRecipientsArr: JsonArray;
        CcRecipientsArr: JsonArray;
        BccRecipientsArr: JsonArray;
        ToList: List of [Text];
        CcList: List of [Text];
        BccList: List of [Text];
        Subject: Text;
        Body: Text;
        Result: Text;
    begin
        Subject := EmailMessage.GetSubject();
        Body := EmailMessage.GetBody();

        BodyObj.Add('contentType', 'HTML');
        BodyObj.Add('content', Body);

        EmailMessage.GetRecipients(Enum::"Email Recipient Type"::"To", ToList);
        EmailMessage.GetRecipients(Enum::"Email Recipient Type"::"Cc", CcList);
        EmailMessage.GetRecipients(Enum::"Email Recipient Type"::"Bcc", BccList);

        ToRecipientsArr := BuildRecipientArray(ToList);
        CcRecipientsArr := BuildRecipientArray(CcList);
        BccRecipientsArr := BuildRecipientArray(BccList);

        MsgObj.Add('subject', Subject);
        MsgObj.Add('body', BodyObj);
        MsgObj.Add('toRecipients', ToRecipientsArr);
        if CcList.Count() > 0 then
            MsgObj.Add('ccRecipients', CcRecipientsArr);
        if BccList.Count() > 0 then
            MsgObj.Add('bccRecipients', BccRecipientsArr);

        MsgObj.WriteTo(Result);
        exit(Result);
    end;

    local procedure BuildRecipientArray(RecipientList: List of [Text]): JsonArray
    var
        RecipientsArr: JsonArray;
        RecipientObj: JsonObject;
        EmailAddressObj: JsonObject;
        Address: Text;
    begin
        foreach Address in RecipientList do begin
            Clear(EmailAddressObj);
            Clear(RecipientObj);
            EmailAddressObj.Add('address', Address);
            RecipientObj.Add('emailAddress', EmailAddressObj);
            RecipientsArr.Add(RecipientObj);
        end;
        exit(RecipientsArr);
    end;

    local procedure BuildAttachmentsArray(EmailMessage: Codeunit "Email Message"; var HasAttachments: Boolean): JsonArray
    var
        AttachmentsArr: JsonArray;
        AttachObj: JsonObject;
        AttachName: Text;
        AttachContentType: Text;
        AttachBase64: Text;
    begin
        HasAttachments := false;
        if not EmailMessage.Attachments_First() then
            exit(AttachmentsArr);

        repeat
            if not EmailMessage.Attachments_IsInline() then begin
                AttachName := EmailMessage.Attachments_GetName();
                AttachContentType := EmailMessage.Attachments_GetContentType();
                AttachBase64 := EmailMessage.Attachments_GetContentBase64();

                Clear(AttachObj);
                AttachObj.Add('@odata.type', '#microsoft.graph.fileAttachment');
                AttachObj.Add('name', AttachName);
                AttachObj.Add('contentType', AttachContentType);
                AttachObj.Add('contentBytes', AttachBase64);
                AttachmentsArr.Add(AttachObj);
                HasAttachments := true;
            end;
        until not EmailMessage.Attachments_Next();

        exit(AttachmentsArr);
    end;

    // -------------------------------------------------------------------------
    // Attachment size helper
    // -------------------------------------------------------------------------

    local procedure GetTotalAttachmentSize(EmailMessage: Codeunit "Email Message"): Integer
    var
        TotalSize: Integer;
    begin
        TotalSize := 0;
        if not EmailMessage.Attachments_First() then
            exit(0);

        repeat
            if not EmailMessage.Attachments_IsInline() then
                TotalSize += EmailMessage.Attachments_GetLength();
        until not EmailMessage.Attachments_Next();

        exit(TotalSize);
    end;

    // -------------------------------------------------------------------------
    // Graph error handling
    // -------------------------------------------------------------------------

    local procedure ParseAndRaiseGraphError(ResponseText: Text; StatusCode: Integer)
    var
        JsonObj: JsonObject;
        ErrorObj: JsonObject;
        JsonToken: JsonToken;
        ErrorCode: Text;
        GenericErr: Label 'Microsoft Graph returned status %1. Check the app registration permissions and try again.';
    begin
        ErrorCode := '';

        if JsonObj.ReadFrom(ResponseText) then
            if JsonObj.Get('error', JsonToken) then begin
                ErrorObj := JsonToken.AsObject();
                if ErrorObj.Get('code', JsonToken) then
                    ErrorCode := JsonToken.AsValue().AsText();
                // error.message intentionally NOT surfaced - may contain user or tenant data
            end;

        if ErrorCode <> '' then
            Error('Microsoft Graph error: %1 (HTTP %2). Check the app registration and try again.', ErrorCode, StatusCode)
        else
            Error(GenericErr, StatusCode);
    end;

    // -------------------------------------------------------------------------
    // Session token helpers
    // -------------------------------------------------------------------------

    local procedure ClearSessionToken()
    var
        GraphSession: Codeunit "W365 Graph Session";
    begin
        GraphSession.ClearToken();
    end;

    // -------------------------------------------------------------------------
    // String helpers
    // -------------------------------------------------------------------------

    local procedure EscapeJsonString(InputText: Text): Text
    begin
        exit(InputText
            .Replace('\', '\\')
            .Replace('"', '\"')
            .Replace(#10, '\n')
            .Replace(#13, '\r')
            .Replace(#9, '\t'));
    end;
}

