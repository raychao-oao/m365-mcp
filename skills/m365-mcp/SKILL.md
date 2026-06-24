---
name: m365-mcp
description: This skill should be used when the user asks to "read email", "check mail", "list messages", "draft a reply", "check calendar", "list events", "check Teams", "send a chat message", "list files on OneDrive", "search SharePoint", or any Microsoft 365 task. Provides access to Mail, Calendar, Teams, OneDrive, and SharePoint via Microsoft Graph API.
---

# m365-mcp

Microsoft 365 MCP server providing 33 tools across Mail, Calendar, Teams, OneDrive, and SharePoint. All access is via Microsoft Graph API using delegated permissions.

**Tool prefix**: `mcp__m365-mcp__`

## Tools

### Mail

| Tool | Description |
|------|-------------|
| `list_messages` | List messages in a folder (default: inbox). Supports `search` and `filter` params. |
| `get_message` | Get full message body by ID. |
| `list_attachments` | List attachments of a message. Use param `id` (not `messageId`). |
| `get_attachment` | Get attachment content. Text types decoded to UTF-8; binary files returned as `{ encoding: "base64", content: "..." }` — decode with `Buffer.from(content, 'base64')`. |
| `create_draft` | Create a draft email. |
| `create_reply_draft` | Create a reply draft to an existing message. Supports reply-all. |
| `update_draft` | Update an existing draft. |
| `send_draft` | Send a draft. ⚠️ **Fails** — `Mail.Send` permission not granted. Users must send manually from Outlook. |
| `move_message` | Move a message to another folder. |
| `mark_message` | Mark as read/unread or set/clear follow-up flag. |
| `forward_message` | Forward a message to recipients with optional comment. |
| `list_folders` | List mail folders. |
| `create_folder` | Create a new mail folder. |
| `list_categories` | List all Outlook categories (master category list). |
| `assign_categories` | Assign categories to a message by display name. |
| `get_out_of_office` | Get current automatic replies settings. |
| `set_out_of_office` | Enable, schedule, or disable automatic replies. |

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
| `get_chat_message` | Get a single chat message with full details including attachments. |
| `get_chat_message_images` | Get hosted inline images from a Teams chat message (returns base64). |
| `get_channel_message` | Get a single channel message with full details including attachments. |
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

**Read a Teams message with attachments:**
```
list_chats() or list_teams()/list_channels()  # find IDs
list_chat_messages(chatId: "<id>")            # find message ID
get_chat_message(chatId: "<id>", messageId: "<id>")  # get attachments[]
# Each attachment has: id, name, contentType, contentUrl
# For file references (contentType: "reference"), use SharePoint tools to download
```

## Authentication

Run `node auth.js` once in the repo directory to authenticate via device-code flow. Token is cached for future sessions. Re-run after any admin consent flow to ensure the daily-use account is active.

If tools return auth errors, the token may have expired — re-run `node auth.js`.
