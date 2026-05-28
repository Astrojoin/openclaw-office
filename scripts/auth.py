#!/usr/bin/env python3
"""OAuth2 device-code flow + token refresh for Microsoft Graph API.

Uses MSAL (Microsoft Authentication Library) for Python.
Tokens are stored at the path specified in config.json (default: ~/openclaw-onedrive/tokens.json).

Usage:
  python3 auth.py login          # Start device-code flow
  python3 auth.py status         # Check if token is valid
  python3 auth.py token          # Print current access token (refreshes if needed)
  python3 auth.py logout         # Delete stored tokens
"""

from __future__ import annotations

import json
import os
import sys
import time
from pathlib import Path

import msal

# ── Config ──────────────────────────────────────────────────────────────────

SKILL_DIR = Path(__file__).resolve().parent.parent
CONFIG_PATH = SKILL_DIR / "config.json"


def _load_config() -> dict:
    with open(CONFIG_PATH) as f:
        return json.load(f)


def _token_path(cfg: dict) -> Path:
    return Path(os.path.expanduser(cfg.get("token_path", "~/openclaw-onedrive/tokens.json")))


def _load_tokens(cfg: dict) -> dict | None:
    p = _token_path(cfg)
    if p.exists():
        with open(p) as f:
            return json.load(f)
    return None


def _save_tokens(cfg: dict, tokens: dict):
    p = _token_path(cfg)
    p.parent.mkdir(parents=True, exist_ok=True)
    with open(p, "w") as f:
        json.dump(tokens, f, indent=2)
    # Restrict permissions to owner only (tokens contain secrets)
    os.chmod(p, 0o600)


def _delete_tokens(cfg: dict):
    p = _token_path(cfg)
    if p.exists():
        p.unlink()


def _build_app(cfg: dict) -> msal.PublicClientApplication:
    return msal.PublicClientApplication(
        client_id=cfg["client_id"],
        authority=cfg.get("authority", "https://login.microsoftonline.com/common"),
        token_cache=_build_cache(cfg),
    )


def _build_cache(cfg: dict) -> msal.SerializableTokenCache:
    cache = msal.SerializableTokenCache()
    tokens = _load_tokens(cfg)
    if tokens and "cache" in tokens:
        cache.deserialize(tokens["cache"])
    return cache


# ── Commands ────────────────────────────────────────────────────────────────

def cmd_login(cfg: dict):
    """Start device-code authorization flow."""
    if not cfg.get("client_id"):
        print("ERROR: client_id is empty in config.json. Set your Azure Application ID first.")
        sys.exit(1)

    app = _build_app(cfg)
    scopes = cfg.get("scopes", ["Files.ReadWrite.All", "Mail.ReadWrite", "Mail.Send",
                                 "Calendars.ReadWrite", "User.Read", "offline_access"])

    flow = app.initiate_device_flow(scopes=scopes)
    if "user_code" not in flow:
        print("ERROR: Failed to create device flow:", flow.get("error_description", flow))
        sys.exit(1)

    print(f"\n🔐 To sign in, open: {flow['verification_uri']}")
    print(f"   Enter code: {flow['user_code']}")
    print(f"   Expires in: {flow.get('expires_in', 900)} seconds\n")
    print("Waiting for authentication...")

    result = app.acquire_token_by_device_flow(flow)

    if "access_token" in result:
        # Save the full cache
        cache = app.token_cache
        token_data = {
            "access_token": result["access_token"],
            "expires_at": result.get("expires_at", time.time() + result.get("expires_in", 3600)),
            "scope": result.get("scope", ""),
            "cache": cache.serialize(),
            "obtained_at": time.time(),
        }
        _save_tokens(cfg, token_data)
        print("✅ Authentication successful. Token saved.")
    else:
        print("❌ Authentication failed:", result.get("error_description", result.get("error", "unknown")))
        sys.exit(1)


def cmd_status(cfg: dict):
    """Check if we have a valid token."""
    tokens = _load_tokens(cfg)
    if not tokens:
        print("No token found. Run: python3 auth.py login")
        return

    expires_at = tokens.get("expires_at", 0)
    now = time.time()
    remaining = expires_at - now

    if remaining > 300:  # More than 5 min left
        print(f"✅ Token valid. Expires in {int(remaining/60)} minutes.")
    elif remaining > 0:
        print(f"⚠️  Token expires in {int(remaining)} seconds. Will refresh on next use.")
    else:
        # Try refresh
        refreshed = _try_refresh(cfg)
        if refreshed:
            print("✅ Token refreshed successfully.")
        else:
            print("❌ Token expired and refresh failed. Run: python3 auth.py login")


def cmd_token(cfg: dict):
    """Get a valid access token, refreshing if necessary."""
    token = get_access_token(cfg)
    if token:
        print(token)
    else:
        print("ERROR: No valid token. Run: python3 auth.py login", file=sys.stderr)
        sys.exit(1)


def cmd_logout(cfg: dict):
    """Delete stored tokens."""
    _delete_tokens(cfg)
    print("Token deleted. Run: python3 auth.py login to re-authenticate.")


# ── Refresh logic ───────────────────────────────────────────────────────────

def _try_refresh(cfg: dict) -> bool:
    """Attempt silent token refresh using MSAL cache. Returns True on success."""
    if not cfg.get("client_id"):
        return False

    app = _build_app(cfg)
    scopes = cfg.get("scopes", ["Files.ReadWrite.All", "Mail.ReadWrite", "Mail.Send",
                                 "Calendars.ReadWrite", "User.Read", "offline_access"])

    accounts = app.get_accounts()
    if not accounts:
        return False

    result = app.acquire_token_silent(scopes, account=accounts[0])
    if result and "access_token" in result:
        cache = app.token_cache
        token_data = {
            "access_token": result["access_token"],
            "expires_at": result.get("expires_at", time.time() + result.get("expires_in", 3600)),
            "scope": result.get("scope", ""),
            "cache": cache.serialize(),
            "obtained_at": time.time(),
        }
        _save_tokens(cfg, token_data)
        return True
    return False


# ── Public API ──────────────────────────────────────────────────────────────

def get_access_token(cfg: dict | None = None) -> str | None:
    """Return a valid access token, refreshing if needed. Returns None on failure.

    Other scripts should call this function instead of managing tokens directly.
    """
    if cfg is None:
        cfg = _load_config()

    tokens = _load_tokens(cfg)
    if not tokens:
        return None

    expires_at = tokens.get("expires_at", 0)
    if time.time() < expires_at - 60:
        return tokens.get("access_token")

    # Try refresh
    if _try_refresh(cfg):
        tokens = _load_tokens(cfg)
        return tokens.get("access_token") if tokens else None

    return None


def get_config() -> dict:
    """Load and return the skill config."""
    return _load_config()


# ── CLI ─────────────────────────────────────────────────────────────────────

def main():
    if len(sys.argv) < 2:
        print(__doc__)
        sys.exit(1)

    cfg = _load_config()
    cmd = sys.argv[1].lower()

    if cmd == "login":
        cmd_login(cfg)
    elif cmd == "status":
        cmd_status(cfg)
    elif cmd == "token":
        cmd_token(cfg)
    elif cmd == "logout":
        cmd_logout(cfg)
    else:
        print(f"Unknown command: {cmd}")
        print("Valid: login, status, token, logout")
        sys.exit(1)


if __name__ == "__main__":
    main()
