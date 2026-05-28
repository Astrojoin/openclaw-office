# 🏢 openclaw-office

An [OpenClaw](https://openclaw.ai) skill for Microsoft 365 — read, create, and edit **Word**, **PowerPoint**, and **Excel** files on OneDrive; send and read **Outlook** email; manage your **Outlook** calendar. All via the Microsoft Graph API.

## ✨ Features

| Module | What it does |
|---|---|
| **OneDrive** | Browse, upload, download, move, copy, rename, delete, and search files |
| **Word** | Create and edit `.docx` files — headings, paragraphs, tables, images, find & replace |
| **PowerPoint** | Create and edit `.pptx` files — slides, textboxes, tables, images, speaker notes |
| **Excel** | Read/write ranges, add worksheets, tables, and formulas (server-side via Graph API) |
| **Outlook Mail** | List, search, read, send, reply, forward, move, and delete emails |
| **Outlook Calendar** | List, create, update, and delete events; accept/decline invitations |

## 📋 Prerequisites

- **Python 3.8+** installed and available as `python3`
- **OpenClaw** installed and running
- **Python packages:** `python-docx`, `python-pptx`, `msal`, `requests`, `pillow`, `lxml`

  You can install them system-wide or in a virtual environment:

  ```bash
  # System-wide
  pip3 install python-docx python-pptx msal requests pillow lxml

  # Or with a virtual environment
  python3 -m venv ~/.venvs/openclaw-office
  source ~/.venvs/openclaw-office/bin/activate
  pip install python-docx python-pptx msal requests pillow lxml
  ```

## 🚀 Installation

1. **Clone the skill** into your OpenClaw skills directory:

   ```bash
   cd ~/.openclaw/workspace/skills
   git clone https://github.com/Astrojoin/openclaw-office.git
   ```

2. **Verify installation** — the skill should appear in your OpenClaw skills menu. You can also confirm by asking your AI assistant: *"What skills do you have available?"* — `openclaw-office` should be listed.

3. **Authenticate with Microsoft** — the first time you use the skill, the AI assistant will detect that authentication is needed and guide you through the OAuth2 device-code flow. You'll visit a URL, enter a code, and authorize the app. Tokens are stored securely outside the skill source code at `~/openclaw-onedrive/tokens.json` and auto-refresh afterwards.

4. **Start using it** — just talk to your AI assistant naturally. See the [Usage](#-usage) section below for examples of what you can ask.

## 💬 Usage

This skill is triggered automatically by your AI assistant when it detects you need to work with Microsoft 365. Just ask naturally:

**OneDrive:**
- *"What files are in my OneDrive?"*
- *"Upload this file to my Documents folder"*
- *"Search my OneDrive for the quarterly report"*

**Word:**
- *"Create a new Word document with a project summary"*
- *"Read my report.docx from OneDrive"*
- *"Replace 'Draft' with 'Final' in my report"*

**PowerPoint:**
- *"Make a presentation about the product roadmap"*
- *"Add a slide with the budget numbers to my deck"*

**Excel:**
- *"What sheets are in my data.xlsx?"*
- *"Read cells A1 to D10 from the budget spreadsheet"*
- *"Write these numbers into Excel and add a SUM formula"*

**Outlook Mail:**
- *"Check my inbox for unread emails"*
- *"Send an email to the team with the report attached"*
- *"Reply to John's email saying I'll review it tomorrow"*

**Outlook Calendar:**
- *"What meetings do I have this week?"*
- *"Schedule a standup tomorrow at 9am with the team"*
- *"Accept the invite from Sarah"*

For the full API reference and all available operations, see [SKILL.md](./SKILL.md).

## 🔐 Security Notes

- **Authentication** uses OAuth2 device-code flow via [MSAL](https://github.com/AzureAD/microsoft-authentication-library-for-python) — the official Microsoft auth library
- **Tokens** are stored locally at `~/openclaw-onedrive/tokens.json` — outside the skill source code, never committed to git
- Word/PowerPoint editing happens **in memory** — files are never written to local disk

## 📄 License

MIT — see [LICENSE](./LICENSE). Free to use, modify, and redistribute; the copyright notice must be included in all copies.
