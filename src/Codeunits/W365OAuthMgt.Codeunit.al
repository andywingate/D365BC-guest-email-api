namespace Wingate365.GuestEmailAPI;

codeunit 50106 "W365 OAuth Mgt"
{
    Access = Internal;

    // -------------------------------------------------------------------------
    // IsolatedStorage keys (all DataScope::User except client secret)
    // -------------------------------------------------------------------------
    // 'W365_AT'    - access token (Text, DataScope::User)
    // 'W365_RT'    - refresh token (Text, DataScope::User)
    // 'W365_EXP'   - token expiry DateTime as Text (DataScope::User)
    // 'W365_CV'    - PKCE code verifier, temporary (DataScope::User)
    // 'W365_STATE' - OAuth state for CSRF check, temporary (DataScope::User)
    // 'W365_CS'    - client secret (Text, DataScope::Company) - set by admin

    // -------------------------------------------------------------------------
    // Public API
    // -------------------------------------------------------------------------

    /// <summary>
    /// Builds the OAuth 2.0 authorisation URL including PKCE challenge and state.
    /// Stores the code verifier and state in IsolatedStorage for later validation.
    /// Open this URL in a browser via Hyperlink() to start the consent flow.
    /// </summary>
    procedure GetAuthorizationUrl(): Text
    var
        Setup: Record "W365 Email Setup";
        CodeVerifier: Text;
        StateValue: Text;
        AuthUrl: Text;
        BaseUrl: Label 'https://login.microsoftonline.com/common/oauth2/v2.0/authorize', Locked = true;
    begin
        Setup.GetOrError();

        CodeVerifier := GeneratePKCEVerifier();
        StateValue := GenerateState();

        // Store temporarily - needed when the user returns with the redirect URL
        IsolatedStorage.Set('W365_CV', CodeVerifier, DataScope::User);
        IsolatedStorage.Set('W365_STATE', StateValue, DataScope::User);

        AuthUrl :=
            BaseUrl +
            '?client_id=' + Setup."App ID" +
            '&response_type=code' +
            '&redirect_uri=' + UrlEncode(Setup."Redirect URI") +
            '&scope=' + UrlEncode('Mail.Send offline_access') +
            '&code_challenge=' + CodeVerifier +
            '&code_challenge_method=plain' +
            // TODO Phase 2: upgrade to S256 - compute SHA-256 of verifier then base64url-encode
            '&state=' + StateValue +
            '&prompt=select_account';

        exit(AuthUrl);
    end;

    /// <summary>
    /// Completes the OAuth flow by exchanging the auth code for tokens.
    /// The user pastes the full redirect URL they received after consent.
    /// Tokens are stored in IsolatedStorage; metadata is written to W365 User Email Token.
    /// </summary>
    procedure ExchangeCodeForToken(RedirectUrl: Text): Boolean
    var
        Setup: Record "W365 Email Setup";
        AuthCode: Text;
        StateParam: Text;
        StoredState: Text;
        CodeVerifier: Text;
        ClientSecret: Text;
        RequestBody: Text;
        AccessToken: Text;
        RefreshToken: Text;
        ExpiresIn: Integer;
        NoCodeErr: Label 'No authorisation code found in the redirect URL. Make sure you pasted the full URL from the browser address bar.';
        StateMismatchErr: Label 'Authorisation state mismatch. Please restart the consent flow.';
        NoVerifierErr: Label 'PKCE code verifier not found. Please restart the consent flow.';
        NoSecretErr: Label 'Client secret has not been configured. Open the W365 Email Setup Card and use the Set Client Secret action.';
    begin
        AuthCode := ExtractQueryParam(RedirectUrl, 'code');
        if AuthCode = '' then
            Error(NoCodeErr);

        // State validation (CSRF protection) - fail-closed: stored state MUST be present and MUST match
        StateParam := ExtractQueryParam(RedirectUrl, 'state');
        IsolatedStorage.Get('W365_STATE', DataScope::User, StoredState);
        if StoredState = '' then
            Error('Consent flow state not found. Please restart the consent flow.');
        if StateParam <> StoredState then
            Error(StateMismatchErr);

        IsolatedStorage.Get('W365_CV', DataScope::User, CodeVerifier);
        if CodeVerifier = '' then
            Error(NoVerifierErr);

        Setup.GetOrError();

        if not IsolatedStorage.Get('W365_CS', DataScope::Company, ClientSecret) then
            Error(NoSecretErr);

        RequestBody :=
            'client_id=' + UrlEncode(Setup."App ID") +
            '&grant_type=authorization_code' +
            '&code=' + UrlEncode(AuthCode) +
            '&redirect_uri=' + UrlEncode(Setup."Redirect URI") +
            '&code_verifier=' + CodeVerifier +
            '&client_secret=' + UrlEncode(ClientSecret) +
            '&scope=' + UrlEncode('Mail.Send offline_access');

        if not DoTokenRequest(RequestBody, AccessToken, RefreshToken, ExpiresIn) then
            exit(false);

        StoreTokens(AccessToken, RefreshToken, ExpiresIn);

        // Clean up temporary consent state
        IsolatedStorage.Delete('W365_CV', DataScope::User);
        IsolatedStorage.Delete('W365_STATE', DataScope::User);

        exit(true);
    end;

    /// <summary>
    /// Returns a valid access token for the current user.
    /// Automatically refreshes the token if it is about to expire.
    /// Returns false if no token exists and consent is required.
    /// </summary>
    procedure GetAccessToken(var AccessToken: Text): Boolean
    var
        ExpiryText: Text;
        TokenExpiry: DateTime;
    begin
        if not IsolatedStorage.Contains('W365_AT', DataScope::User) then
            exit(false);

        // Check whether the stored token is still valid
        if IsolatedStorage.Get('W365_EXP', DataScope::User, ExpiryText) then
            if Evaluate(TokenExpiry, ExpiryText) then
                if CurrentDateTime() < TokenExpiry - 60000 then begin // 60-second buffer
                    IsolatedStorage.Get('W365_AT', DataScope::User, AccessToken);
                    exit(true);
                end;

        // Token is expired (or expiry unreadable) - try silent refresh
        exit(RefreshAccessToken(AccessToken));
    end;

    /// <summary>Returns true if the current user has a valid (or refreshable) token.</summary>
    procedure HasValidToken(): Boolean
    var
        AccessToken: Text;
    begin
        exit(GetAccessToken(AccessToken));
    end;

    /// <summary>
    /// Stores the client secret in Company-scoped IsolatedStorage.
    /// Called from the Email Setup Card. The value is never read back to the UI.
    /// </summary>
    procedure SetClientSecret(ClientSecretValue: Text)
    begin
        IsolatedStorage.Set('W365_CS', ClientSecretValue, DataScope::Company);
    end;

    /// <summary>
    /// Returns true if a client secret has been configured (does not return the value).
    /// </summary>
    procedure HasClientSecret(): Boolean
    begin
        exit(IsolatedStorage.Contains('W365_CS', DataScope::Company));
    end;

    /// <summary>Removes all tokens for the current user. Forces re-consent on next send.</summary>
    procedure ClearTokens()
    var
        UserToken: Record "W365 User Email Token";
        UserName: Code[50];
    begin
        IsolatedStorage.Delete('W365_AT', DataScope::User);
        IsolatedStorage.Delete('W365_RT', DataScope::User);
        IsolatedStorage.Delete('W365_EXP', DataScope::User);

        UserName := CopyStr(UserId(), 1, MaxStrLen(UserName));
        if UserToken.Get(UserName) then begin
            UserToken."Consent Status" := "W365 Consent Status"::None;
            UserToken."Token Expiry" := 0DT;
            UserToken."Last Error" := '';
            UserToken.Modify();
        end;
    end;

    // -------------------------------------------------------------------------
    // Token refresh
    // -------------------------------------------------------------------------

    local procedure RefreshAccessToken(var AccessToken: Text): Boolean
    var
        Setup: Record "W365 Email Setup";
        RefreshToken: Text;
        ClientSecret: Text;
        RequestBody: Text;
        NewRefreshToken: Text;
        ExpiresIn: Integer;
    begin
        if not IsolatedStorage.Get('W365_RT', DataScope::User, RefreshToken) then
            exit(false);

        if RefreshToken = '' then
            exit(false);

        Setup.GetOrError();

        if not IsolatedStorage.Get('W365_CS', DataScope::Company, ClientSecret) then
            exit(false);

        RequestBody :=
            'client_id=' + UrlEncode(Setup."App ID") +
            '&grant_type=refresh_token' +
            '&refresh_token=' + UrlEncode(RefreshToken) +
            '&client_secret=' + UrlEncode(ClientSecret) +
            '&scope=' + UrlEncode('Mail.Send offline_access');

        if not DoTokenRequest(RequestBody, AccessToken, NewRefreshToken, ExpiresIn) then
            exit(false);

        StoreTokens(AccessToken, NewRefreshToken, ExpiresIn);
        exit(true);
    end;

    // -------------------------------------------------------------------------
    // HTTP token endpoint call
    // -------------------------------------------------------------------------

    local procedure DoTokenRequest(RequestBody: Text; var AccessToken: Text; var RefreshToken: Text; var ExpiresIn: Integer): Boolean
    var
        HttpClient: HttpClient;
        HttpReqMsg: HttpRequestMessage;
        HttpRespMsg: HttpResponseMessage;
        HttpContent: HttpContent;
        ContentHeaders: HttpHeaders;
        ResponseText: Text;
        JsonObj: JsonObject;
        JsonToken: JsonToken;
        ErrorCode: Text;
        TokenEndpoint: Label 'https://login.microsoftonline.com/common/oauth2/v2.0/token', Locked = true;
        ConnectErr: Label 'Could not reach the Microsoft identity platform. Check the BC server outbound connectivity.';
        InvalidResponseErr: Label 'Unexpected response from the token endpoint.';
    begin
        HttpContent.WriteFrom(RequestBody);
        HttpContent.GetHeaders(ContentHeaders);
        ContentHeaders.Remove('Content-Type');
        ContentHeaders.Add('Content-Type', 'application/x-www-form-urlencoded');

        HttpReqMsg.Method := 'POST';
        HttpReqMsg.SetRequestUri(TokenEndpoint);
        HttpReqMsg.Content := HttpContent;

        if not HttpClient.Send(HttpReqMsg, HttpRespMsg) then begin
            UpdateTokenError('Connection failed: ' + GetLastErrorText());
            Error(ConnectErr);
        end;

        HttpRespMsg.Content.ReadAs(ResponseText);

        if not JsonObj.ReadFrom(ResponseText) then begin
            UpdateTokenError(InvalidResponseErr);
            Error(InvalidResponseErr);
        end;

        if not HttpRespMsg.IsSuccessStatusCode() then begin
            ErrorCode := '';
            if JsonObj.Get('error', JsonToken) then
                ErrorCode := JsonToken.AsValue().AsText();
            // error_description is intentionally NOT included in user-facing messages
            // as it may contain tenant or user claim information
            UpdateTokenError('Token error: ' + ErrorCode);
            Error('Token request failed (error: %1). Check the app registration and try again.', ErrorCode);
        end;

        if JsonObj.Get('access_token', JsonToken) then
            AccessToken := JsonToken.AsValue().AsText();

        if JsonObj.Get('refresh_token', JsonToken) then
            RefreshToken := JsonToken.AsValue().AsText();

        if JsonObj.Get('expires_in', JsonToken) then
            ExpiresIn := JsonToken.AsValue().AsInteger()
        else
            ExpiresIn := 3600;

        exit(true);
    end;

    // -------------------------------------------------------------------------
    // Storage helpers
    // -------------------------------------------------------------------------

    local procedure StoreTokens(AccessToken: Text; RefreshToken: Text; ExpiresInSeconds: Integer)
    var
        UserToken: Record "W365 User Email Token";
        GraphMailMgt: Codeunit "W365 Graph Mail Mgt";
        UserName: Code[50];
        ExpiryDt: DateTime;
        HomeEmail: Text;
    begin
        ExpiryDt := CurrentDateTime() + (ExpiresInSeconds * 1000);

        // Store tokens - IsolatedStorage encrypts at rest
        // Tokens are never read back to the UI or included in any log/telemetry
        IsolatedStorage.Set('W365_AT', AccessToken, DataScope::User);
        if RefreshToken <> '' then
            IsolatedStorage.Set('W365_RT', RefreshToken, DataScope::User);
        IsolatedStorage.Set('W365_EXP', Format(ExpiryDt), DataScope::User);

        // Update metadata table so admin can see token status
        UserName := CopyStr(UserId(), 1, MaxStrLen(UserName));
        if not UserToken.Get(UserName) then begin
            UserToken.Init();
            UserToken."User Name" := UserName;
            UserToken.Insert();
        end;
        UserToken."Token Expiry" := ExpiryDt;
        UserToken."Consent Status" := "W365 Consent Status"::Active;
        UserToken."Last Error" := '';

        // Fetch real home email from Graph /me and store it for display
        HomeEmail := GraphMailMgt.GetCurrentUserEmail();
        if HomeEmail <> '' then
            UserToken."Home Email" := CopyStr(HomeEmail, 1, MaxStrLen(UserToken."Home Email"));

        UserToken.Modify();
    end;

    local procedure UpdateTokenError(ErrorMessage: Text)
    var
        UserToken: Record "W365 User Email Token";
        UserName: Code[50];
    begin
        UserName := CopyStr(UserId(), 1, MaxStrLen(UserName));
        if not UserToken.Get(UserName) then begin
            UserToken.Init();
            UserToken."User Name" := UserName;
            UserToken.Insert();
        end;
        UserToken."Consent Status" := "W365 Consent Status"::Error;
        UserToken."Last Error" := CopyStr(ErrorMessage, 1, MaxStrLen(UserToken."Last Error"));
        UserToken.Modify();
    end;

    // -------------------------------------------------------------------------
    // URL and PKCE helpers
    // -------------------------------------------------------------------------

    local procedure GeneratePKCEVerifier(): Text
    var
        Result: Text;
        GuidText: Text;
        i: Integer;
    begin
        // Build a 96-char verifier from 3 GUIDs (each stripped to 32 hex chars).
        // Using code_challenge_method=plain so verifier = challenge (no hashing needed).
        // TODO Phase 2: implement S256 (SHA-256 + base64url) for stronger security.
        for i := 1 to 3 do begin
            GuidText := Format(CreateGuid());
            GuidText := GuidText.Replace('{', '').Replace('}', '').Replace('-', '');
            Result += GuidText;
        end;
        exit(CopyStr(Result, 1, 96));
    end;

    local procedure GenerateState(): Text
    var
        StateGuid: Guid;
        StateText: Text;
    begin
        StateGuid := CreateGuid();
        StateText := Format(StateGuid);
        exit(StateText.Replace('{', '').Replace('}', '').Replace('-', ''));
    end;

    internal procedure ExtractQueryParam(Url: Text; ParamName: Text): Text
    var
        SearchKey: Text;
        StartPos: Integer;
        Value: Text;
        AmpPos: Integer;
    begin
        SearchKey := ParamName + '=';
        StartPos := StrPos(Url, SearchKey);
        if StartPos = 0 then
            exit('');

        Value := CopyStr(Url, StartPos + StrLen(SearchKey));
        AmpPos := StrPos(Value, '&');
        if AmpPos > 0 then
            Value := CopyStr(Value, 1, AmpPos - 1);

        exit(Value);
    end;

    local procedure UrlEncode(InputText: Text): Text
    var
        Result: Text;
        i: Integer;
        OneChar: Text[1];
    begin
        for i := 1 to StrLen(InputText) do begin
            OneChar := CopyStr(InputText, i, 1);
            case OneChar of
                ' ':
                    Result += '%20';
                ':':
                    Result += '%3A';
                '/':
                    Result += '%2F';
                '?':
                    Result += '%3F';
                '#':
                    Result += '%23';
                '@':
                    Result += '%40';
                '&':
                    Result += '%26';
                '=':
                    Result += '%3D';
                '+':
                    Result += '%2B';
                '%':
                    Result += '%25';
                '~':
                    Result += '%7E';
                else
                    Result += OneChar;
            end;
        end;
        exit(Result);
    end;
}
