#!/usr/bin/env python3
"""Outlook mail operations via Microsoft Graph API.

Read inbox, search messages, read full message body, send emails, reply, forward, move, delete.

Usage:
  from mail import Mail

  m = Mail()
  messages = m.list_inbox(count=10)
  body = m.read(message_id)
  m.send(to=["user@example.com"], subject="Hello", body="Test email")
"""

from __future__ import annotations

import sys
from pathlib import Path

SCRIPT_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(SCRIPT_DIR))

from graph_client import GraphClient


class Mail:
    """Outlook mail operations via Microsoft Graph API."""

    def __init__(self):
        self.client = GraphClient()

    # ── LIST / SEARCH ───────────────────────────────────────────────────

    def list_inbox(self, count: int = 25, skip: int = 0,
                   order_by: str = "receivedDateTime desc") -> list[dict]:
        """List messages in the inbox.

        Args:
            count: Number of messages to return (max 200)
            skip: Number of messages to skip (pagination)
            order_by: Sort field and direction

        Returns:
            List of message summaries.
        """
        params = {
            "$top": min(count, 200),
            "$skip": skip,
            "$orderby": order_by,
            "$select": "id,subject,from,receivedDateTime,isRead,hasAttachments,importance",
        }
        items = self.client.get_all("/me/mailFolders/inbox/messages", params=params)
        return [self._summarize(m) for m in items]

    def list_folder(self, folder: str = "inbox", count: int = 25,
                    skip: int = 0) -> list[dict]:
        """List messages in a specific folder (inbox, sentitems, drafts, deleteditems, junkemail)."""
        folder_map = {
            "inbox": "inbox",
            "sent": "sentitems",
            "sentitems": "sentitems",
            "drafts": "drafts",
            "deleted": "deleteditems",
            "deleteditems": "deleteditems",
            "junk": "junkemail",
            "junkemail": "junkemail",
            "spam": "junkemail",
        }
        folder_name = folder_map.get(folder.lower(), folder.lower())
        params = {
            "$top": min(count, 200),
            "$skip": skip,
            "$orderby": "receivedDateTime desc",
            "$select": "id,subject,from,receivedDateTime,isRead,hasAttachments,importance",
        }
        items = self.client.get_all(f"/me/mailFolders/{folder_name}/messages", params=params)
        return [self._summarize(m) for m in items]

    def search(self, query: str, count: int = 25) -> list[dict]:
        """Search messages across all folders.

        Args:
            query: Search query (subject, body, sender, etc.)
            count: Max results
        """
        params = {
            "$search": f'"{query}"',
            "$top": min(count, 200),
            "$select": "id,subject,from,receivedDateTime,isRead",
        }
        items = self.client.get_all("/me/messages", params=params)
        return [self._summarize(m) for m in items]

    # ── READ ────────────────────────────────────────────────────────────

    def read(self, message_id: str, mime: bool = False) -> dict:
        """Read a full message by ID.

        Args:
            message_id: The message ID
            mime: If True, returns raw MIME content

        Returns:
            Full message dict with body content.
        """
        if mime:
            endpoint = f"/me/messages/{message_id}/$value"
            data = self.client.download(endpoint)
            return {"mime": data.decode("utf-8", errors="replace")}

        item = self.client.get(
            f"/me/messages/{message_id}",
            params={"$select": "id,subject,from,toRecipients,ccRecipients,bccRecipients,body,receivedDateTime,sentDateTime,isRead,importance,hasAttachments"}
        )
        return self._full_message(item)

    def read_body(self, message_id: str) -> str:
        """Read just the body text of a message."""
        item = self.client.get(
            f"/me/messages/{message_id}",
            params={"$select": "body"}
        )
        body = item.get("body", {})
        return body.get("content", "")

    # ── MARK / MOVE / DELETE ────────────────────────────────────────────

    def mark_read(self, message_id: str) -> dict:
        """Mark a message as read."""
        return self.client.patch(f"/me/messages/{message_id}", {"isRead": True})

    def mark_unread(self, message_id: str) -> dict:
        """Mark a message as unread."""
        return self.client.patch(f"/me/messages/{message_id}", {"isRead": False})

    def move(self, message_id: str, destination_folder: str) -> dict:
        """Move a message to a different folder.

        Args:
            destination_folder: Folder name (inbox, deleteditems, junkemail, etc.)
        """
        return self.client.post(
            f"/me/messages/{message_id}/move",
            {"destinationId": destination_folder}
        )

    def delete(self, message_id: str) -> bool:
        """Delete a message (moves to deleted items, then permanent on next delete)."""
        return self.client.delete(f"/me/messages/{message_id}")

    # ── SEND ────────────────────────────────────────────────────────────

    def send(self, to: list[str], subject: str, body: str,
             cc: list[str] | None = None, bcc: list[str] | None = None,
             body_type: str = "text", reply_to: str | None = None,
             attachments: list[dict] | None = None) -> str:
        """Send an email.

        Args:
            to: List of recipient email addresses
            subject: Email subject
            body: Email body content
            cc: List of CC addresses
            bcc: List of BCC addresses
            body_type: "text" or "html"
            reply_to: Reply-to email address
            attachments: Optional list of {"name": "...", "content_bytes": bytes} or {"name": "...", "path": "..."}

        Returns:
            Status string.
        """
        message = {
            "subject": subject,
            "body": {
                "contentType": "HTML" if body_type == "html" else "Text",
                "content": body,
            },
            "toRecipients": [{"emailAddress": {"address": addr}} for addr in to],
        }

        if cc:
            message["ccRecipients"] = [{"emailAddress": {"address": addr}} for addr in cc]
        if bcc:
            message["bccRecipients"] = [{"emailAddress": {"address": addr}} for addr in bcc]
        if reply_to:
            message["replyTo"] = [{"emailAddress": {"address": reply_to}}]

        # Handle attachments
        if attachments:
            msg_attachments = []
            for att in attachments:
                att_data = {
                    "@odata.type": "#microsoft.graph.fileAttachment",
                    "name": att["name"],
                    "contentType": att.get("content_type", "application/octet-stream"),
                }
                if "content_bytes" in att:
                    import base64
                    att_data["contentBytes"] = base64.b64encode(att["content_bytes"]).decode()
                elif "path" in att:
                    import base64
                    with open(att["path"], "rb") as f:
                        att_data["contentBytes"] = base64.b64encode(f.read()).decode()
                msg_attachments.append(att_data)
            message["attachments"] = msg_attachments

        # Send directly (no draft)
        self.client.post("/me/sendMail", {"message": message})
        return f"Sent to {', '.join(to)}: {subject}"

    def reply(self, message_id: str, body: str,
              body_type: str = "text", reply_all: bool = False) -> str:
        """Reply to a message.

        Args:
            message_id: ID of the message to reply to
            body: Reply body
            body_type: "text" or "html"
            reply_all: If True, reply to all recipients
        """
        endpoint = f"/me/messages/{message_id}/replyAll" if reply_all else f"/me/messages/{message_id}/reply"
        comment = body
        self.client.post(endpoint, {"comment": comment})
        mode = "all" if reply_all else "sender"
        return f"Replied ({mode}) to message {message_id}"

    def forward(self, message_id: str, to: list[str], comment: str = "") -> str:
        """Forward a message.

        Args:
            message_id: ID of the message to forward
            to: List of recipient addresses
            comment: Optional comment to add
        """
        data = {
            "comment": comment,
            "toRecipients": [{"emailAddress": {"address": addr}} for addr in to],
        }
        self.client.post(f"/me/messages/{message_id}/forward", data)
        return f"Forwarded message {message_id} to {', '.join(to)}"

    # ── HELPERS ─────────────────────────────────────────────────────────

    @staticmethod
    def _summarize(msg: dict) -> dict:
        """Extract summary fields from a message."""
        sender = msg.get("from", {}).get("emailAddress", {})
        return {
            "id": msg.get("id", ""),
            "subject": msg.get("subject", "(no subject)"),
            "from": sender.get("address", sender.get("name", "")),
            "from_name": sender.get("name", ""),
            "received": msg.get("receivedDateTime", ""),
            "isRead": msg.get("isRead", False),
            "importance": msg.get("importance", "normal"),
            "hasAttachments": msg.get("hasAttachments", False),
        }

    @staticmethod
    def _full_message(msg: dict) -> dict:
        """Extract full message fields."""
        sender = msg.get("from", {}).get("emailAddress", {})
        to_addrs = [r.get("emailAddress", {}).get("address", "") for r in msg.get("toRecipients", [])]
        cc_addrs = [r.get("emailAddress", {}).get("address", "") for r in msg.get("ccRecipients", [])]
        bcc_addrs = [r.get("emailAddress", {}).get("address", "") for r in msg.get("bccRecipients", [])]

        return {
            "id": msg.get("id", ""),
            "subject": msg.get("subject", "(no subject)"),
            "from": sender.get("address", ""),
            "from_name": sender.get("name", ""),
            "to": to_addrs,
            "cc": cc_addrs,
            "bcc": bcc_addrs,
            "body": msg.get("body", {}).get("content", ""),
            "body_type": msg.get("body", {}).get("contentType", ""),
            "received": msg.get("receivedDateTime", ""),
            "sent": msg.get("sentDateTime", ""),
            "isRead": msg.get("isRead", False),
            "importance": msg.get("importance", "normal"),
            "hasAttachments": msg.get("hasAttachments", False),
        }


# ── CLI ─────────────────────────────────────────────────────────────────────

def main():
    """Quick test: list recent inbox messages."""
    m = Mail()
    try:
        messages = m.list_inbox(count=5)
        print(f"Recent inbox messages ({len(messages)}):")
        for msg in messages:
            status = "✓" if msg["isRead"] else "●"
            print(f"  {status} [{msg['received'][:16]}] {msg['from_name']}: {msg['subject']}")
    except Exception as e:
        print(f"❌ Error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()
