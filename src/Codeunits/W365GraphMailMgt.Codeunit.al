namespace Wingate365.GuestEmailAPI;

codeunit 50105 "W365 Graph Mail Mgt"
{
    Access = Internal;

    var
        OAuthMgt: Codeunit "W365 OAuth Mgt";

    // -------------------------------------------------------------------------
    // Public API
    // -------------------------------------------------------------------------

    /// <summary>
    /// Calls Graph /me and returns the authenticated user's mail address.
    /// Used after token exchange to populate the Home Email field on the token record.
    /// Returns empty string on any failure - non-fatal, falls back to BC username display.
    /// </summary>
    procedure GetCurrentUserEmail(): Text
    var
        HttpClient: HttpClient;
        HttpReqMsg: HttpRequestMessage;
        HttpRespMsg: HttpResponseMessage;
        ReqHeaders: HttpHeaders;
        ResponseText: Text;
        JsonObj: JsonObject;
        JsonToken: JsonToken;
        AccessToken: Text;
        MeEndpoint: Label 'https://graph.microsoft.com/v1.0/me?$select=mail,userPrincipalName', Locked = true;
    begin
        if not OAuthMgt.GetAccessToken(AccessToken) then
            exit('');

        HttpReqMsg.Method := 'GET';
        HttpReqMsg.SetRequestUri(MeEndpoint);
        HttpReqMsg.GetHeaders(ReqHeaders);
        ReqHeaders.Add('Authorization', 'Bearer ' + AccessToken);
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

    /// <summary>
    /// Sends an email via Microsoft Graph using the current user's delegated token.
    /// Returns true on success. Raises an error (with consent action) if no token exists.
    /// </summary>
    procedure SendEmail(ToAddress: Text; Subject: Text; BodyHtml: Text): Boolean
    var
        AccessToken: Text;
        NoTokenErr: Label 'You do not have an active email authorisation. Open the Connect Current User Email API page to connect your email account first.';
    begin
        if not OAuthMgt.GetAccessToken(AccessToken) then
            Error(NoTokenErr);

        exit(DoSendMail(AccessToken, ToAddress, Subject, BodyHtml));
    end;

    /// <summary>
    /// Overload accepting plain-text body as well as HTML. HTML is preferred by Graph.
    /// </summary>
    procedure SendEmail(ToAddress: Text; Subject: Text; BodyHtml: Text; BodyText: Text): Boolean
    var
        AccessToken: Text;
        NoTokenErr: Label 'You do not have an active email authorisation. Open the Connect Current User Email API page to connect your email account first.';
    begin
        if not OAuthMgt.GetAccessToken(AccessToken) then
            Error(NoTokenErr);

        if BodyHtml <> '' then
            exit(DoSendMail(AccessToken, ToAddress, Subject, BodyHtml))
        else
            exit(DoSendMailPlainText(AccessToken, ToAddress, Subject, BodyText));
    end;

    // -------------------------------------------------------------------------
    // Internal send helpers
    // -------------------------------------------------------------------------

    local procedure DoSendMail(AccessToken: Text; ToAddress: Text; Subject: Text; BodyHtml: Text): Boolean
    begin
        exit(ExecuteGraphSendMail(AccessToken, ToAddress, Subject, BodyHtml, 'HTML'));
    end;

    local procedure DoSendMailPlainText(AccessToken: Text; ToAddress: Text; Subject: Text; BodyText: Text): Boolean
    begin
        exit(ExecuteGraphSendMail(AccessToken, ToAddress, Subject, BodyText, 'Text'));
    end;

    local procedure ExecuteGraphSendMail(AccessToken: Text; ToAddress: Text; Subject: Text; Body: Text; ContentType: Text): Boolean
    var
        HttpClient: HttpClient;
        HttpReqMsg: HttpRequestMessage;
        HttpRespMsg: HttpResponseMessage;
        HttpContent: HttpContent;
        ReqHeaders: HttpHeaders;
        ContentHeaders: HttpHeaders;
        JsonBody: Text;
        ResponseText: Text;
        StatusCode: Integer;
        GraphEndpoint: Label 'https://graph.microsoft.com/v1.0/me/sendMail', Locked = true;
        ConnectErr: Label 'Could not reach Microsoft Graph. Check BC server outbound connectivity.';
        ThrottledErr: Label 'Microsoft Graph is throttling requests. Please wait a moment and try again.';
    begin
        JsonBody := BuildSendMailJson(ToAddress, Subject, Body, ContentType);

        HttpContent.WriteFrom(JsonBody);
        HttpContent.GetHeaders(ContentHeaders);
        ContentHeaders.Remove('Content-Type');
        ContentHeaders.Add('Content-Type', 'application/json');

        HttpReqMsg.Method := 'POST';
        HttpReqMsg.SetRequestUri(GraphEndpoint);
        HttpReqMsg.Content := HttpContent;

        // Set Authorization header
        // Note: access token is stored as plain Text in IsolatedStorage (encrypted at rest).
        // It is never written to logs, error messages, or UI fields.
        HttpReqMsg.GetHeaders(ReqHeaders);
        ReqHeaders.Add('Authorization', 'Bearer ' + AccessToken);

        if not HttpClient.Send(HttpReqMsg, HttpRespMsg) then
            Error(ConnectErr);

        StatusCode := HttpRespMsg.HttpStatusCode();

        // Graph sendMail returns 202 Accepted on success (no response body)
        if StatusCode = 202 then
            exit(true);

        // 401 - token rejected; clear cached token so next attempt re-prompts
        if StatusCode = 401 then begin
            OAuthMgt.ClearTokens();
            Error('Microsoft Graph rejected the authorisation token (401). Please re-authorise on the Connect Current User Email API page.');
        end;

        // 429 - throttled; surface gracefully without retrying in a tight loop
        if StatusCode = 429 then
            Error(ThrottledErr);

        // Any other non-success status
        HttpRespMsg.Content.ReadAs(ResponseText);
        ParseAndRaiseGraphError(ResponseText, StatusCode);
        exit(false); // ParseAndRaiseGraphError always raises
    end;

    // -------------------------------------------------------------------------
    // JSON helpers
    // -------------------------------------------------------------------------

    local procedure BuildSendMailJson(ToAddress: Text; Subject: Text; Body: Text; ContentType: Text): Text
    var
        MsgObj: JsonObject;
        BodyObj: JsonObject;
        RecipientsArr: JsonArray;
        RecipientObj: JsonObject;
        EmailAddressObj: JsonObject;
        RootObj: JsonObject;
        Result: Text;
    begin
        EmailAddressObj.Add('address', ToAddress);
        RecipientObj.Add('emailAddress', EmailAddressObj);
        RecipientsArr.Add(RecipientObj);

        BodyObj.Add('contentType', ContentType);
        BodyObj.Add('content', Body);

        MsgObj.Add('subject', Subject);
        MsgObj.Add('body', BodyObj);
        MsgObj.Add('toRecipients', RecipientsArr);

        RootObj.Add('message', MsgObj);
        RootObj.Add('saveToSentItems', true);

        RootObj.WriteTo(Result);
        exit(Result);
    end;

    // -------------------------------------------------------------------------
    // Error handling
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
}
