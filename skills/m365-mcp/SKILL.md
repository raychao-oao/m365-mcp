---
name: m365-mcp
description: This skill should be used when the user asks to "read email", "check mail", "list messages", "draft a reply", "check calendar", "list events", "check Teams", "send a chat message", "list files on OneDrive", "search SharePoint", or any Microsoft 365 task. Provides access to Mail, Calendar, Teams, OneDrive, and SharePoint via Microsoft Graph API.
---

# m365-mcp

Microsoft 365 MCP server providing 24 tools across Mail, Calendar, Teams, OneDrive, and SharePoint. All access is via Microsoft Graph API using delegated permissions.

**Tool prefix**: `mcp__m365-mcp__`

## Tools

### Mail

| Tool | Description |
|------|-------------|
| `list_messages` | List messages in a folder (default: inbox). Supports `search` and `filter` params. |
| `get_message` | Get full message body by ID. |
| `list_attachments` | List attachments of a message. Use param `id` (not `messageId`). |
| `get_attachment` | Get attachment content. Text types decoded to UTF-8; images/binary returned as base64. |
| `create_draft` | Create a draft email. |
| `update_draft` | Update an existing draft. |
| `send_draft` | Send a draft. ⚠️ **Fails** — `Mail.Send` permission not granted. Users must send manually from Outlook. |
| `move_message` | Move a message to another folder. |
| `list_folders` | List mail folders. |
| `create_folder` | Create a new mail folder. |

### Calendar

| Tool | Description |
|------|-------------|
| `list_events` | List calendar events. Supports date range filters. |
| `create_event` | Create a calendar event. |

### Teams

| Tool | Description |
|------|-------------|
| `list_teams` | List teams the user belongs to. |
| `list_channels` | List channels in a team. |
| `list_channel_messages` | List messages in a team channel. |
| `list_chats` | List personal chats. |
| `list_chat_messages` | List messages in a chat. |
| `send_chat_message` | Send a message to a chat. |

### OneDrive

| Tool | Description |
|------|-------------|
| `list_drive_items` | List files/folders in OneDrive. |
| `get_drive_item_content` | Get content of a file. |
| `search_drive` | Search for files by keyword. |

### SharePoint

| Tool | Description |
|------|-------------|
| `list_sites` | List SharePoint sites. |
| `list_site_drives` | List document libraries in a site. |
| `list_site_drive_items` | List items in a document library. |

## Security Constraints

- Do NOT ask the user to type passwords or tokens in the chat.
- Token is cached at `~/.m365-mcp-token.json` — never read or log this file's content.
- The admin account (used for admin consent) is separate from the daily-use account. Always ensure tools are invoked as the daily-use account after any admin consent flow.
- These tools are used to **verify MCP functionality**, not to handle actual company business. Do not proactively manage company communications.

## Common Patterns

**Read latest inbox:**
```
list_messages(folder: "inbox", top: 10)
```

**Read a message with attachments:**
```
get_message(id: "<message_id>")
list_attachments(id: "<message_id>")
get_attachment(messageId: "<message_id>", attachmentId: "<attachment_id>")
```

**Draft a reply:**
```
create_draft(to: "...", subject: "RE: ...", body: "...")
# User must send manually — send_draft lacks Mail.Send permission
```

**Check today's calendar:**
```
list_events(startDateTime: "2026-06-17T00:00:00", endDateTime: "2026-06-17T23:59:59")
```

**Send a Teams chat message:**
```
list_chats()  # find chat ID
send_chat_message(chatId: "<id>", message: "...")
```

## Authentication

Run `node auth.js` once in the repo directory to authenticate via device-code flow. Token is cached for future sessions. Re-run after any admin consent flow to ensure the daily-use account is active.

If tools return auth errors, the token may have expired — re-run `node auth.js`.
