namespace Wingate365.GuestEmailAPI;

/// <summary>
/// SingleInstance codeunit that holds the in-memory access token for the current BC session.
/// Tokens live for the lifetime of the BC session and are never written to storage.
/// Re-authentication via SSO is transparent when the browser session is still signed in to Entra.
/// Each BC user has their own instance because SingleInstance = true scopes to the client session.
/// </summary>
codeunit 50114 "W365 Graph Session"
{
    SingleInstance = true;
    Access = Internal;

    var
        SessionToken: SecretText;
        HasToken: Boolean;
        TokenExpiry: DateTime;

    /// <summary>Returns true if there is a valid non-expired token in memory.</summary>
    [NonDebuggable]
    procedure HasValidToken(): Boolean
    var
        ExpiryBufferMs: Integer;
    begin
        ExpiryBufferMs := 60000;
        if not HasToken then
            exit(false);
        if TokenExpiry = 0DT then
            exit(false);
        exit(CurrentDateTime() < TokenExpiry - ExpiryBufferMs);
    end;

    /// <summary>
    /// Stores an access token in session memory.
    /// The expiry is set 50 minutes from now (conservative - MSAL manages actual expiry via silent re-acquire).
    /// </summary>
    [NonDebuggable]
    procedure SetToken(AccessToken: SecretText)
    begin
        SessionToken := AccessToken;
        HasToken := true;
        // 50 minutes (conservative; MSAL caches the real expiry and handles silent refresh)
        TokenExpiry := CurrentDateTime() + (50 * 60 * 1000);
    end;

    /// <summary>
    /// Retrieves the in-memory access token.
    /// Call HasValidToken() first to ensure a valid token exists.
    /// </summary>
    [NonDebuggable]
    procedure GetToken(var AccessToken: SecretText)
    begin
        AccessToken := SessionToken;
    end;

    /// <summary>Clears the in-memory token, forcing re-authentication on the next send.</summary>
    [NonDebuggable]
    procedure ClearToken()
    var
        EmptyToken: SecretText;
    begin
        SessionToken := EmptyToken;
        HasToken := false;
        TokenExpiry := 0DT;
    end;
}
