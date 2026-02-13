# Outlook Desktop MCP Server

A local MCP server that gives full control over **Outlook Desktop** via COM automation – calendar, email, contacts, and tasks.  
**No Azure AD, no admin consent, no cloud permissions required.** Windows only.

---

## Quick Setup (for Copilot agents)

> **If you are a Copilot agent setting this up for a user, follow these steps exactly.**

### Prerequisites

| Requirement | Details |
| ----------- | ------- |
| OS | Windows (COM automation requires Win32) |
| Outlook | Classic Outlook Desktop must be **installed and running** (not "New Outlook") |
| Python | 3.10 or newer, accessible from terminal |
| VS Code | With GitHub Copilot extension installed |

### Step 1 – Install dependencies

Run in the project root directory:

```powershell
cd <project-root>            # e.g. C:\GIT\outlook-mcp
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

The only dependencies are `mcp>=1.0.0` and `pywin32>=306`.

### Step 2 – Register the MCP server in VS Code

Add the following entry to the user-level MCP configuration file.

**File location:** `%APPDATA%\Code\User\mcp.json`  
**VS Code command:** `MCP: Open User Configuration` (Ctrl+Shift+P)

Add this under the `"servers"` key (adjust paths to match the actual project location):

```jsonc
{
  "servers": {
    "outlook": {
      "type": "stdio",
      "command": "<project-root>\\.venv\\Scripts\\python.exe",
      "args": [
        "<project-root>\\server.py"
      ]
    }
  }
}
```

Replace `<project-root>` with the **absolute path** to this repository, using double backslashes.  
Example: `"C:\\GIT\\outlook-mcp\\.venv\\Scripts\\python.exe"`

### Step 3 – Verify the server is running

After saving `mcp.json`, VS Code will auto-detect and start the server.

- Open **Ctrl+Shift+P** → `MCP: List Servers`
- The `outlook` server should appear with a green status indicator
- If it does not start, select it and click **Start**

### Step 4 – Use it

Open Copilot Chat (**Ctrl+Shift+I**) and try prompts like:

- *"What meetings do I have today?"*
- *"Search my emails for quarterly report"*
- *"Send an email to jane.doe@example.com with the subject Hello"*
- *"Create a meeting tomorrow at 2pm: Team Standup"*
- *"Show my unread emails"*
- *"Flag the email about budget review for follow-up"*
- *"List my open tasks"*

Copilot automatically selects the right MCP tool based on your request.

### Troubleshooting

| Problem | Solution |
| ------- | -------- |
| Server won't start | Verify the `python.exe` path in `mcp.json` points to the `.venv` Python executable |
| Import errors | Activate the venv (`.venv\Scripts\activate`) and run `pip install -r requirements.txt` |
| COM / Outlook errors | Make sure **Outlook Desktop** is open (not just minimized to system tray) |
| Tools not visible | **Ctrl+Shift+P** → `MCP: List Servers` → check that `outlook` is listed and green |

---

## Features – 35 Tools

### Calendar (9 tools)

| Tool | Description |
| ---- | ----------- |
| `get_next_meeting` | Show your next upcoming meeting |
| `list_events_today` | All remaining events for today |
| `list_events` | Events in any date range |
| `get_event_details` | Full details (attendees, body) of an event |
| `search_events` | Search by subject, organizer, or attendee |
| `create_event` | Create a new calendar event |
| `respond_to_meeting` | Accept / tentative / decline a meeting |
| `update_event` | Update subject, time, location, etc. |
| `delete_event` | Delete an event |

### Email (21 tools)

| Tool | Description |
| ---- | ----------- |
| `list_unread_emails` | Unread emails from Inbox |
| `list_recent_emails` | Most recent emails |
| `search_emails` | Fast search by subject (COM filter) |
| `search_emails_full` | Deep search across subject, body & sender with date range |
| `read_email` | Read the full body of an email |
| `send_email` | Send an email |
| `reply_to_email` | Reply or reply-all to an email |
| `forward_email` | Forward an email |
| `delete_email` | Delete an email |
| `move_email` | Move an email to a different folder |
| `mark_email` | Mark as read / unread |
| `flag_email` | Flag / unflag for follow-up |
| `list_email_folders` | List all folders with item & unread counts |
| `list_sent_emails` | Recent sent emails |
| `list_drafts` | Draft emails |
| `create_draft` | Create a draft (saved, not sent) |
| `list_email_attachments` | List attachments of an email |
| `save_attachment` | Save an attachment to local disk |

### Contacts (3 tools)

| Tool | Description |
| ---- | ----------- |
| `search_contacts` | Search by name, email, or company |
| `get_contact_details` | Full contact details (phones, addresses, etc.) |
| `create_contact` | Create a new contact |

### Tasks (5 tools)

| Tool | Description |
| ---- | ----------- |
| `list_tasks` | List tasks (optionally include completed) |
| `create_task` | Create a new task with due date, priority, etc. |
| `complete_task` | Mark a task as complete |
| `update_task` | Update subject, due date, status, progress |
| `delete_task` | Delete a task |

---

## Project Structure

```
outlook-mcp/
├── server.py                  # Entry point – FastMCP setup & tool registration
├── requirements.txt           # mcp + pywin32
├── outlook_client/            # COM automation client (mixin-based)
│   ├── __init__.py            # Exports OutlookClient
│   ├── _helpers.py            # Shared utilities, constants, serializers
│   ├── calendar.py            # CalendarMixin – 9 calendar methods
│   ├── email.py               # EmailMixin – 18 email methods
│   ├── contacts.py            # ContactsMixin – 3 contact methods
│   ├── tasks.py               # TasksMixin – 5 task methods
│   └── client.py              # OutlookClient (composes all mixins)
└── tools/                     # MCP tool definitions (thin wrappers)
    ├── __init__.py             # register_all_tools()
    ├── calendar.py             # 9 calendar tools
    ├── email.py                # 21 email tools
    ├── contacts.py             # 3 contact tools
    └── tasks.py                # 5 task tools
```
