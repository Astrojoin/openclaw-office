#!/usr/bin/env python3
"""OneDrive cloud filesystem operations — the main entry point for cloud file work.

Uses graph_client.py for all Graph API calls.
Uses word.py and powerpoint.py for .docx/.pptx content editing (in-memory).
Uses graph_client.py Excel endpoints for .xlsx operations.

All file content stays in Python variables (bytes) — no disk writes unless
the user explicitly requests a local path.

Usage:
  from onedrive import OneDrive

  od = OneDrive()
  files = od.list("/Documents")
  content = od.download("/Documents/report.docx")
  od.upload("/Documents/new_report.docx", bytes_data)
"""

from __future__ import annotations

import json
import sys
from pathlib import Path
from urllib.parse import quote

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))

from graph_client import GraphClient
from word import WordDoc
from powerpoint import PptxDoc


class OneDrive:
    """OneDrive file operations via Microsoft Graph API."""

    def __init__(self):
        self.client = GraphClient()

    # ── Path helpers ────────────────────────────────────────────────────

    @staticmethod
    def _api_path(remote_path: str) -> str:
        """Convert a user-friendly path like /Documents/file.docx
        to a Graph API path segment: :/Documents/file.docx:"""
        path = remote_path.strip("/")
        return f":/{path}"

    # ── LIST ────────────────────────────────────────────────────────────

    def list(self, remote_path: str = "/") -> list[dict]:
        """List files and folders in a OneDrive directory.

        Args:
            remote_path: Path like "/" or "/Documents"

        Returns:
            List of item dicts with name, id, size, type, lastModified, etc.
        """
        if remote_path == "/" or remote_path == "":
            endpoint = "/me/drive/root/children"
        else:
            endpoint = f"/me/drive/root{self._api_path(remote_path)}:/children"

        items = self.client.get_all(endpoint)
        result = []
        for item in items:
            info = {
                "name": item.get("name", ""),
                "id": item.get("id", ""),
                "size": item.get("size", 0),
                "type": "folder" if "folder" in item else "file",
                "lastModified": item.get("lastModifiedDateTime", ""),
            }
            if "file" in item:
                info["mimeType"] = item["file"].get("mimeType", "")
            result.append(info)
        return result

    # ── SEARCH ──────────────────────────────────────────────────────────

    def search(self, query: str) -> list[dict]:
        """Search OneDrive for files matching a query."""
        endpoint = f"/me/drive/root/search(q='{quote(query)}')"
        items = self.client.get_all(endpoint)
        result = []
        for item in items:
            info = {
                "name": item.get("name", ""),
                "id": item.get("id", ""),
                "size": item.get("size", 0),
                "type": "folder" if "folder" in item else "file",
                "lastModified": item.get("lastModifiedDateTime", ""),
                "webUrl": item.get("webUrl", ""),
            }
            if "file" in item:
                info["mimeType"] = item["file"].get("mimeType", "")
            result.append(info)
        return result

    # ── GET INFO ────────────────────────────────────────────────────────

    def info(self, remote_path: str) -> dict:
        """Get metadata for a specific file or folder."""
        endpoint = f"/me/drive/root{self._api_path(remote_path)}:"
        item = self.client.get(endpoint)
        return {
            "name": item.get("name", ""),
            "id": item.get("id", ""),
            "size": item.get("size", 0),
            "type": "folder" if "folder" in item else "file",
            "lastModified": item.get("lastModifiedDateTime", ""),
            "created": item.get("createdDateTime", ""),
            "webUrl": item.get("webUrl", ""),
            "mimeType": item.get("file", {}).get("mimeType", ""),
        }

    # ── DOWNLOAD ────────────────────────────────────────────────────────

    def download(self, remote_path: str) -> bytes:
        """Download a file from OneDrive. Returns content as bytes (in memory)."""
        endpoint = f"/me/drive/root{self._api_path(remote_path)}:/content"
        return self.client.download(endpoint)

    def download_by_id(self, item_id: str) -> bytes:
        """Download a file by its DriveItem ID."""
        endpoint = f"/me/drive/items/{item_id}/content"
        return self.client.download(endpoint)

    # ── UPLOAD ──────────────────────────────────────────────────────────

    def upload(self, remote_path: str, data: bytes) -> dict:
        """Upload file content to OneDrive. Handles small (<4MB) and large files.

        Args:
            remote_path: Target path in OneDrive
            data: File content as bytes

        Returns:
            DriveItem metadata of the uploaded file.
        """
        if len(data) <= 4 * 1024 * 1024:
            endpoint = f"/me/drive/root{self._api_path(remote_path)}:/content"
            return self.client.upload_small(endpoint, data)
        else:
            endpoint = f"/me/drive/root{self._api_path(remote_path)}:/createUploadSession"
            return self.client.upload_large(endpoint, data)

    # ── MOVE / COPY / RENAME / DELETE ───────────────────────────────────

    def move(self, remote_path: str, dest_folder: str, new_name: str | None = None) -> dict:
        """Move a file to a different folder, optionally renaming it."""
        # Get item ID from source path
        item_info = self.info(remote_path)
        item_id = item_info["id"]

        # Get destination folder ID
        dest_info = self.info(dest_folder)
        dest_id = dest_info["id"]

        data = {"parentReference": {"id": dest_id}}
        if new_name:
            data["name"] = new_name

        endpoint = f"/me/drive/items/{item_id}"
        return self.client.patch(endpoint, data)

    def copy(self, remote_path: str, dest_folder: str, new_name: str | None = None) -> str:
        """Copy a file to a different folder. Returns status."""
        item_info = self.info(remote_path)
        item_id = item_info["id"]

        dest_info = self.info(dest_folder)
        dest_id = dest_info["id"]

        data = {"parentReference": {"id": dest_id}}
        if new_name:
            data["name"] = new_name

        endpoint = f"/me/drive/items/{item_id}/copy"
        # Copy is async — returns 202 Accepted
        url = f"{self.client.base}{endpoint}"
        headers = self.client._headers()
        headers["Content-Type"] = "application/json"
        resp = self.client._session.post(url, headers=headers, data=json.dumps(data), timeout=30)
        if resp.status_code == 202:
            return "Copy initiated (async). Check destination folder shortly."
        resp.raise_for_status()
        return "Copy completed."

    def rename(self, remote_path: str, new_name: str) -> dict:
        """Rename a file or folder."""
        item_info = self.info(remote_path)
        item_id = item_info["id"]
        endpoint = f"/me/drive/items/{item_id}"
        return self.client.patch(endpoint, {"name": new_name})

    def delete(self, remote_path: str) -> bool:
        """Delete a file or folder from OneDrive."""
        item_info = self.info(remote_path)
        item_id = item_info["id"]
        endpoint = f"/me/drive/items/{item_id}"
        return self.client.delete(endpoint)

    # ── CREATE FOLDER ───────────────────────────────────────────────────

    def create_folder(self, remote_path: str, name: str) -> dict:
        """Create a new folder in OneDrive.

        Args:
            remote_path: Parent folder path
            name: New folder name
        """
        if remote_path == "/" or remote_path == "":
            endpoint = "/me/drive/root/children"
        else:
            endpoint = f"/me/drive/root{self._api_path(remote_path)}:/children"
        return self.client.post(endpoint, {"name": name, "folder": {}})

    # ── WORD (.docx) OPERATIONS ─────────────────────────────────────────

    def docx_read(self, remote_path: str, mode: str = "plain") -> str:
        """Download a .docx from OneDrive and read its text content.

        Args:
            remote_path: Path to .docx file in OneDrive
            mode: "plain" or "structured"

        Returns:
            Document text content.
        """
        data = self.download(remote_path)
        doc = WordDoc(data=data)
        return doc.read(mode=mode)

    def docx_edit(self, remote_path: str, operations: list[dict]) -> str:
        """Download a .docx, apply operations, and upload the modified version.

        Args:
            remote_path: Path to .docx file in OneDrive
            operations: List of {"method": "...", "args": [...], "kwargs": {...}}

        Returns:
            Summary of operations applied.
        """
        data = self.download(remote_path)
        doc = WordDoc(data=data)
        result = doc.update({"operations": operations})
        new_data = doc.to_bytes()
        self.upload(remote_path, new_data)
        return result

    def docx_create(self, remote_path: str, operations: list[dict] | None = None) -> str:
        """Create a new .docx on OneDrive with optional initial operations.

        Args:
            remote_path: Target path in OneDrive
            operations: Optional list of operations to apply

        Returns:
            Summary of what was created.
        """
        doc = WordDoc.create()
        if operations:
            doc.update({"operations": operations})
        data = doc.to_bytes()
        self.upload(remote_path, data)
        return f"Created: {remote_path} ({len(data)} bytes)"

    # ── POWERPOINT (.pptx) OPERATIONS ───────────────────────────────────

    def pptx_read(self, remote_path: str, mode: str = "plain") -> str:
        """Download a .pptx from OneDrive and read its content."""
        data = self.download(remote_path)
        pptx = PptxDoc(data=data)
        return pptx.read(mode=mode)

    def pptx_edit(self, remote_path: str, operations: list[dict]) -> str:
        """Download a .pptx, apply operations, and upload the modified version."""
        data = self.download(remote_path)
        pptx = PptxDoc(data=data)
        result = pptx.update({"operations": operations})
        new_data = pptx.to_bytes()
        self.upload(remote_path, new_data)
        return result

    def pptx_create(self, remote_path: str, operations: list[dict] | None = None) -> str:
        """Create a new .pptx on OneDrive with optional initial operations."""
        pptx = PptxDoc.create()
        if operations:
            pptx.update({"operations": operations})
        data = pptx.to_bytes()
        self.upload(remote_path, data)
        return f"Created: {remote_path} ({len(data)} bytes)"

    # ── EXCEL (.xlsx) OPERATIONS ────────────────────────────────────────

    def xlsx_list_worksheets(self, remote_path: str) -> list[dict]:
        """List worksheets in an Excel file on OneDrive."""
        item_info = self.info(remote_path)
        return self.client.excel_list_worksheets(item_info["id"])

    def xlsx_read_range(self, remote_path: str, worksheet: str,
                        range_addr: str = "A1:Z100") -> dict:
        """Read a range from an Excel worksheet via Graph API.

        Returns dict with 'values' (2D array) and 'text' fields.
        """
        item_info = self.info(remote_path)
        return self.client.excel_read_range(item_info["id"], worksheet, range_addr)

    def xlsx_write_range(self, remote_path: str, worksheet: str,
                         range_addr: str, values: list[list]) -> dict:
        """Write values to a range in an Excel worksheet via Graph API."""
        item_info = self.info(remote_path)
        return self.client.excel_write_range(item_info["id"], worksheet, range_addr, values)

    def xlsx_add_worksheet(self, remote_path: str, name: str) -> dict:
        """Add a new worksheet to an Excel file on OneDrive."""
        item_info = self.info(remote_path)
        return self.client.excel_add_worksheet(item_info["id"], name)

    def xlsx_add_table(self, remote_path: str, worksheet: str,
                       range_addr: str, has_headers: bool = True) -> dict:
        """Add a table to an Excel worksheet."""
        item_info = self.info(remote_path)
        return self.client.excel_add_table(item_info["id"], worksheet, range_addr, has_headers)

    def xlsx_add_formula(self, remote_path: str, worksheet: str,
                         cell: str, formula: str) -> dict:
        """Write a formula to a cell in an Excel worksheet."""
        item_info = self.info(remote_path)
        return self.client.excel_add_formula(item_info["id"], worksheet, cell, formula)

    def xlsx_get_used_range(self, remote_path: str, worksheet: str) -> dict:
        """Get the used range of a worksheet."""
        item_info = self.info(remote_path)
        return self.client.excel_get_used_range(item_info["id"], worksheet)


# ── CLI ─────────────────────────────────────────────────────────────────────

def main():
    """Quick test: list root folder."""
    od = OneDrive()
    try:
        items = od.list("/")
        print(f"Found {len(items)} items in root:")
        for item in items[:10]:
            icon = "📁" if item["type"] == "folder" else "📄"
            print(f"  {icon} {item['name']} ({item['size']} bytes)")
        if len(items) > 10:
            print(f"  ... and {len(items) - 10} more")
    except Exception as e:
        print(f"❌ Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
