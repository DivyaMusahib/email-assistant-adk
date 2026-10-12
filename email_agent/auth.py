"""Microsoft sign-in (OAuth 2.0) using MSAL.

Big picture
-----------
1. `python login.py`  -> you sign in once in a browser (or with a device code).
   MSAL stores the resulting tokens in `.token_cache.json`.
2. Every tool call asks `get_access_token()`.  That silently reuses the cached
   token and refreshes it when it expires.  No browser needed again.

We use a *public client* (no client secret) because this is a local app on the
user's own machine, which is the pattern Microsoft recommends for desktop/CLI apps.
"""

import os

import msal

from . import config


class AuthError(Exception):
    """Raised when we cannot get a valid token. The message is safe to show."""


def _load_cache() -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    if config.TOKEN_CACHE_PATH.exists():
        cache.deserialize(config.TOKEN_CACHE_PATH.read_text())
    return cache


def _save_cache(cache: msal.SerializableTokenCache) -> None:
    # MSAL flips `has_state_changed` when tokens were added or refreshed.
    if cache.has_state_changed:
        config.TOKEN_CACHE_PATH.write_text(cache.serialize())
        # Only the owner may read the file (tokens are like passwords).
        os.chmod(config.TOKEN_CACHE_PATH, 0o600)


def _build_app(cache: msal.SerializableTokenCache) -> msal.PublicClientApplication:
    if not config.MS_CLIENT_ID:
        raise AuthError(
            "MS_CLIENT_ID is missing. Copy .env.example to .env and fill it in "
            "(see README, step 'Register the app in Microsoft Entra')."
        )
    return msal.PublicClientApplication(
        client_id=config.MS_CLIENT_ID,
        authority=config.MS_AUTHORITY,
        token_cache=cache,
    )


def login(use_device_code: bool = False) -> dict:
    """Interactive sign-in. Called by login.py. Returns the MSAL result dict."""
    cache = _load_cache()
    app = _build_app(cache)

    if use_device_code:
        # Shows a short code; you type it at https://microsoft.com/devicelogin.
        flow = app.initiate_device_flow(scopes=config.MS_SCOPES)
        if "user_code" not in flow:
            raise AuthError(f"Could not start device-code login: {flow}")
        print(flow["message"])
        result = app.acquire_token_by_device_flow(flow)  # blocks until you finish
    else:
        # Opens your default browser and listens on http://localhost.
        # PKCE (an anti-interception safeguard) is applied automatically.
        result = app.acquire_token_interactive(scopes=config.MS_SCOPES)

    if "access_token" not in result:
        raise AuthError(
            f"Login failed: {result.get('error')} - {result.get('error_description')}"
        )

    _save_cache(cache)
    return result


def get_access_token() -> str:
    """Return a valid access token, refreshing silently if needed."""
    cache = _load_cache()
    app = _build_app(cache)

    accounts = app.get_accounts()
    if not accounts:
        raise AuthError("You are not signed in. Run `python login.py` first.")

    result = app.acquire_token_silent(config.MS_SCOPES, account=accounts[0])
    if not result or "access_token" not in result:
        raise AuthError(
            "Your Microsoft session expired. Run `python login.py` again."
        )

    _save_cache(cache)
    return result["access_token"]


def logout() -> bool:
    """Forget the cached login. Returns True if a cache file was removed."""
    if config.TOKEN_CACHE_PATH.exists():
        config.TOKEN_CACHE_PATH.unlink()
        return True
    return False
