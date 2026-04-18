namespace Wingate365.GuestEmailAPI;

controladdin "W365 OAuth Popup"
{
    Scripts = 'src/ControlAddins/W365OAuthPopup.js';
    MinimumHeight = 0;
    MinimumWidth = 0;
    MaximumHeight = 0;
    MaximumWidth = 0;
    HorizontalStretch = false;
    VerticalStretch = false;

    /// <summary>Opens a popup window to the OAuth authorisation URL and monitors for the redirect code.</summary>
    procedure OpenAuthPopup(AuthUrl: Text);

    /// <summary>Fired when the redirect URL containing the code is detected in the popup.</summary>
    event CodeReceived(RedirectUrl: Text);

    /// <summary>Fired when the user closes the popup without completing authorisation.</summary>
    event PopupClosed();

    /// <summary>Fired when the popup could not be opened (e.g. blocked by browser).</summary>
    event PopupError(ErrorMessage: Text);
}
