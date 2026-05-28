#!/usr/bin/env python3
"""Microsoft Graph API HTTP client with Excel workbook/session support.

This module is the low-level HTTP layer for all Graph API calls.
Other scripts (onedrive.py, mail.py, outlook_calendar.py) use this client.

Excel operations use the Graph API workbook/session endpoints directly
(no openpyxl — all editing happens server-side via Microsoft Graph).

Usage from other scripts:
    from graph_client import GraphClient
    client = GraphClient()
    result = client.get("/me/drive/root/children")
    content = client.download("/me/drive/items/{item_id}/content")
    client.upload("/me/drive/root:/{path}:/content", bytes_data)
"""

from __future__ import annotations

import json
import sys
import time
from pathlib import Path
from typing import Any

import requests

# Import auth from same directory
SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))

from auth import get_access_token, get_config

GRAPH_BASE = "https://graph.microsoft.com/v1.0"


class GraphClient:
    """Low-level Microsoft Graph API client with auto-refresh."""

    def __init__(self, config: dict | None = None):
        self.config = config or get_config()
        self.base = self.config.get("graph_base", GRAPH_BASE)
        self._session = requests.Session()
        self._session.verify = True  # Respect system CA bundle

    def _headers(self) -> dict:
        token = get_access_token(self.config)
        if not token:
            raise RuntimeError("No valid access token. Run: python3 auth.py login")
        return {
            "Authorization": f"Bearer {token}",
            "Accept": "application/json",
        }

    # ── Generic HTTP methods ────────────────────────────────────────────

    def get(self, endpoint: str, params: dict | None = None) -> dict:
        """GET request. Returns parsed JSON."""
        url = f"{self.base}{endpoint}"
        resp = self._session.get(url, headers=self._headers(), params=params, timeout=30)
        resp.raise_for_status()
        return resp.json() if resp.content else {}

    def get_raw(self, endpoint: str, params: dict | None = None) -> requests.Response:
        """GET request. Returns raw Response object (for binary content)."""
        url = f"{self.base}{endpoint}"
        resp = self._session.get(url, headers=self._headers(), params=params, timeout=60)
        resp.raise_for_status()
        return resp

    def post(self, endpoint: str, data: dict | None = None,
             content_type: str = "application/json") -> dict:
        """POST request with JSON body."""
        url = f"{self.base}{endpoint}"
        headers = self._headers()
        if content_type == "application/json":
            headers["Content-Type"] = "application/json"
            body = json.dumps(data) if data else None
        else:
            body = data
        resp = self._session.post(url, headers=headers, data=body, timeout=30)
        resp.raise_for_status()
        return resp.json() if resp.content else {}

    def patch(self, endpoint: str, data: dict) -> dict:
        """PATCH request with JSON body."""
        url = f"{self.base}{endpoint}"
        headers = self._headers()
        headers["Content-Type"] = "application/json"
        resp = self._session.patch(url, headers=headers, data=json.dumps(data), timeout=30)
        resp.raise_for_status()
        return resp.json() if resp.content else {}

    def delete(self, endpoint: str) -> bool:
        """DELETE request. Returns True on success."""
        url = f"{self.base}{endpoint}"
        resp = self._session.delete(url, headers=self._headers(), timeout=30)
        resp.raise_for_status()
        return True

    def put(self, endpoint: str, data: bytes | str,
            content_type: str = "application/octet-stream") -> dict:
        """PUT request (typically for file uploads). Returns parsed JSON."""
        url = f"{self.base}{endpoint}"
        headers = self._headers()
        headers["Content-Type"] = content_type
        resp = self._session.put(url, headers=headers, data=data, timeout=120)
        resp.raise_for_status()
        return resp.json() if resp.content else {}

    def put_create_upload_session(self, endpoint: str) -> dict:
        """Create an upload session for large files (>4MB). Returns upload URL."""
        url = f"{self.base}{endpoint}"
        headers = self._headers()
        headers["Content-Type"] = "application/json"
        resp = self._session.post(url, headers=headers, data=json.dumps({}), timeout=30)
        resp.raise_for_status()
        return resp.json()

    def upload_chunk(self, upload_url: str, chunk: bytes,
                     start: int, total_size: int) -> dict:
        """Upload a chunk in a resumable upload session."""
        end = start + len(chunk) - 1
        headers = {
            "Content-Range": f"bytes {start}-{end}/{total_size}",
            "Content-Length": str(len(chunk)),
        }
        resp = self._session.put(upload_url, headers=headers, data=chunk, timeout=120)
        resp.raise_for_status()
        return resp.json() if resp.content else {}

    # ── Pagination helper ───────────────────────────────────────────────

    def get_all(self, endpoint: str, params: dict | None = None) -> list[dict]:
        """GET with automatic pagination. Returns all items."""
        items = []
        url = f"{self.base}{endpoint}"
        params = params or {}
        while url:
            if url.startswith(self.base):
                resp = self._session.get(url, headers=self._headers(), params=params, timeout=30)
            else:
                # Next page URL from @odata.nextLink (full absolute URL)
                resp = self._session.get(url, headers=self._headers(), timeout=30)
            resp.raise_for_status()
            data = resp.json()
            value = data.get("value", [])
            items.extend(value)
            url = data.get("@odata.nextLink")
            params = None  # Only apply params on first request
        return items

    # ── Download/Upload convenience ─────────────────────────────────────

    def download(self, endpoint: str) -> bytes:
        """Download file content as bytes."""
        resp = self.get_raw(endpoint)
        return resp.content

    def upload_small(self, endpoint: str, data: bytes) -> dict:
        """Upload file up to 4MB via simple PUT."""
        return self.put(endpoint, data)

    def upload_large(self, endpoint: str, data: bytes, chunk_size: int = 4 * 1024 * 1024) -> dict:
        """Upload file >4MB via resumable upload session."""
        session_info = self.put_create_upload_session(endpoint)
        upload_url = session_info["uploadUrl"]
        total_size = len(data)
        start = 0
        result = {}
        while start < total_size:
            chunk = data[start:start + chunk_size]
            result = self.upload_chunk(upload_url, chunk, start, total_size)
            start += chunk_size
        return result

    # ── Excel workbook/session operations ───────────────────────────────

    def excel_create_session(self, item_id: str, persist: bool = True) -> dict:
        """Create an Excel workbook session for ephemeral or persistent editing.

        Args:
            item_id: DriveItem ID of the .xlsx file
            persist: If True, changes are saved to the file. If False, changes are discarded.

        Returns:
            Session info dict with 'id' key.
        """
        endpoint = f"/me/drive/items/{item_id}/workbook/createSession"
        data = {"persistChanges": persist}
        return self.post(endpoint, data)

    def excel_close_session(self, item_id: str, session_id: str):
        """Close an Excel workbook session."""
        endpoint = f"/me/drive/items/{item_id}/workbook/closeSession"
        headers = self._headers()
        headers["Workbook-Session-Id"] = session_id
        headers["Content-Type"] = "application/json"
        url = f"{self.base}{endpoint}"
        self._session.post(url, headers=headers, data=json.dumps({}), timeout=30)

    def excel_get(self, item_id: str, session_id: str, path: str,
                  params: dict | None = None) -> dict:
        """GET within an Excel workbook session.

        Args:
            item_id: DriveItem ID
            session_id: Workbook session ID
            path: Path after /workbook/, e.g. "worksheets/{name}/range(address='A1:Z100')"
        """
        endpoint = f"/me/drive/items/{item_id}/workbook/{path}"
        url = f"{self.base}{endpoint}"
        headers = self._headers()
        headers["Workbook-Session-Id"] = session_id
        resp = self._session.get(url, headers=headers, params=params, timeout=30)
        resp.raise_for_status()
        return resp.json() if resp.content else {}

    def excel_post(self, item_id: str, session_id: str, path: str,
                   data: dict | None = None) -> dict:
        """POST within an Excel workbook session (create worksheets, tables, etc.)."""
        endpoint = f"/me/drive/items/{item_id}/workbook/{path}"
        url = f"{self.base}{endpoint}"
        headers = self._headers()
        headers["Workbook-Session-Id"] = session_id
        headers["Content-Type"] = "application/json"
        body = json.dumps(data) if data else json.dumps({})
        resp = self._session.post(url, headers=headers, data=body, timeout=30)
        resp.raise_for_status()
        return resp.json() if resp.content else {}

    def excel_patch(self, item_id: str, session_id: str, path: str,
                    data: dict) -> dict:
        """PATCH within an Excel workbook session (update ranges, cells, etc.)."""
        endpoint = f"/me/drive/items/{item_id}/workbook/{path}"
        url = f"{self.base}{endpoint}"
        headers = self._headers()
        headers["Workbook-Session-Id"] = session_id
        headers["Content-Type"] = "application/json"
        resp = self._session.patch(url, headers=headers, data=json.dumps(data), timeout=30)
        resp.raise_for_status()
        return resp.json() if resp.content else {}

    def excel_delete(self, item_id: str, session_id: str, path: str) -> bool:
        """DELETE within an Excel workbook session."""
        endpoint = f"/me/drive/items/{item_id}/workbook/{path}"
        url = f"{self.base}{endpoint}"
        headers = self._headers()
        headers["Workbook-Session-Id"] = session_id
        resp = self._session.delete(url, headers=headers, timeout=30)
        resp.raise_for_status()
        return True

    # ── Excel high-level operations ─────────────────────────────────────

    def excel_list_worksheets(self, item_id: str) -> list[dict]:
        """List all worksheets in an Excel workbook."""
        session = self.excel_create_session(item_id, persist=False)
        try:
            result = self.excel_get(item_id, session["id"], "worksheets")
            return result.get("value", [])
        finally:
            self.excel_close_session(item_id, session["id"])

    def excel_read_range(self, item_id: str, worksheet: str,
                         range_addr: str = "A1:Z100") -> dict:
        """Read a range from an Excel worksheet.

        Args:
            item_id: DriveItem ID
            worksheet: Worksheet name
            range_addr: Excel range address (e.g. 'A1:D10')
        """
        session = self.excel_create_session(item_id, persist=False)
        try:
            from urllib.parse import quote
            ws = quote(worksheet, safe="")
            ra = quote(range_addr, safe="")
            path = f"worksheets/{ws}/range(address='{ra}')"
            return self.excel_get(item_id, session["id"], path)
        finally:
            self.excel_close_session(item_id, session["id"])

    def excel_write_range(self, item_id: str, worksheet: str,
                          range_addr: str, values: list[list]) -> dict:
        """Write values to a range in an Excel worksheet.

        Args:
            item_id: DriveItem ID
            worksheet: Worksheet name
            range_addr: Target range (e.g. 'A1:D3')
            values: 2D array of values
        """
        session = self.excel_create_session(item_id, persist=True)
        try:
            from urllib.parse import quote
            ws = quote(worksheet, safe="")
            ra = quote(range_addr, safe="")
            path = f"worksheets/{ws}/range(address='{ra}')"
            return self.excel_patch(item_id, session["id"], path, {"values": values})
        finally:
            self.excel_close_session(item_id, session["id"])

    def excel_add_worksheet(self, item_id: str, name: str) -> dict:
        """Add a new worksheet to an Excel workbook."""
        session = self.excel_create_session(item_id, persist=True)
        try:
            return self.excel_post(item_id, session["id"], "worksheets", {"name": name})
        finally:
            self.excel_close_session(item_id, session["id"])

    def excel_add_table(self, item_id: str, worksheet: str,
                        range_addr: str, has_headers: bool = True) -> dict:
        """Add a table to an Excel worksheet."""
        session = self.excel_create_session(item_id, persist=True)
        try:
            from urllib.parse import quote
            ws = quote(worksheet, safe="")
            path = f"worksheets/{ws}/tables/add"
            data = {
                "address": f"{worksheet}!{range_addr}",
                "hasHeaders": has_headers,
            }
            return self.excel_post(item_id, session["id"], path, data)
        finally:
            self.excel_close_session(item_id, session["id"])

    def excel_add_formula(self, item_id: str, worksheet: str,
                          cell: str, formula: str) -> dict:
        """Write a formula to a single cell."""
        return self.excel_write_range(item_id, worksheet, cell, [[formula]])

    def excel_get_used_range(self, item_id: str, worksheet: str) -> dict:
        """Get the used range of a worksheet."""
        session = self.excel_create_session(item_id, persist=False)
        try:
            from urllib.parse import quote
            ws = quote(worksheet, safe="")
            path = f"worksheets/{ws}/usedRange"
            return self.excel_get(item_id, session["id"], path)
        finally:
            self.excel_close_session(item_id, session["id"])


# ── CLI convenience ─────────────────────────────────────────────────────────

def main():
    """Quick test: verify Graph API connectivity."""
    client = GraphClient()
    try:
        me = client.get("/me")
        print(f"✅ Connected as: {me.get('displayName', 'unknown')} ({me.get('mail', me.get('userPrincipalName', ''))})")
    except Exception as e:
        print(f"❌ Connection failed: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
