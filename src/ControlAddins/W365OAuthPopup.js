// W365OAuthPopup.js
// Opens an OAuth consent popup and automatically captures the redirect code.
// BC and OAuthLanding.htm are both on businesscentral.dynamics.com (same origin),
// so the parent window can read popup.location.href once the redirect arrives.

var w365OAuthTimer = null;
var w365OAuthPopup = null;

function OpenAuthPopup(authUrl) {
    if (w365OAuthPopup && !w365OAuthPopup.closed) {
        w365OAuthPopup.focus();
        return;
    }

    w365OAuthPopup = window.open(
        authUrl,
        'W365OAuthConsent',
        'width=640,height=760,scrollbars=yes,resizable=yes,toolbar=no,menubar=no'
    );

    if (!w365OAuthPopup || w365OAuthPopup.closed) {
        Microsoft.Dynamics.NAV.InvokeExtensibilityMethod(
            'PopupError',
            ['The sign-in window was blocked. Please allow pop-ups for this site and try again.']
        );
        return;
    }

    if (w365OAuthTimer) {
        clearInterval(w365OAuthTimer);
    }

    w365OAuthTimer = setInterval(function () {
        try {
            if (!w365OAuthPopup || w365OAuthPopup.closed) {
                clearInterval(w365OAuthTimer);
                w365OAuthTimer = null;
                Microsoft.Dynamics.NAV.InvokeExtensibilityMethod('PopupClosed', []);
                return;
            }

            var href = w365OAuthPopup.location.href;

            if (href && href.indexOf('code=') > -1) {
                clearInterval(w365OAuthTimer);
                w365OAuthTimer = null;
                w365OAuthPopup.close();
                w365OAuthPopup = null;
                Microsoft.Dynamics.NAV.InvokeExtensibilityMethod('CodeReceived', [href]);
            }
        } catch (e) {
            // Cross-origin exception while popup is navigating through Entra pages - normal, keep polling
        }
    }, 500);
}
