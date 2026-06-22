import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { PublicClientApplication, LogLevel } from '@azure/msal-node';
import { Client } from '@microsoft/microsoft-graph-client';
import { z } from 'zod';
import fs from 'fs';
import path from 'path';
import os from 'os';

const TENANT_ID = process.env.M365_MCP_TENANT_ID;
const CLIENT_ID = process.env.M365_MCP_CLIENT_ID;
if (!TENANT_ID || !CLIENT_ID) {
  process.stderr.write('[m365-mcp] ERROR: M365_MCP_TENANT_ID and M365_MCP_CLIENT_ID must be set\n');
  process.exit(1);
}
// Allow override via env so Claude Code subprocess always finds the right path
const TOKEN_CACHE_PATH =
  process.env.M365_MCP_TOKEN_CACHE ||
  path.join(os.homedir(), '.m365-mcp-token.json');
const SCOPES = [
  'https://graph.microsoft.com/Mail.Read',
  'https://graph.microsoft.com/Mail.ReadWrite',
  'https://graph.microsoft.com/Mail.ReadWrite.Shared',
  'https://graph.microsoft.com/User.Read',
  'https://graph.microsoft.com/Calendars.ReadWrite',
  'https://graph.microsoft.com/Team.ReadBasic.All',
  'https://graph.microsoft.com/Channel.ReadBasic.All',
  'https://graph.microsoft.com/ChannelMessage.Read.All',
  'https://graph.microsoft.com/Chat.ReadWrite',
  'https://graph.microsoft.com/Files.ReadWrite',
  'https://graph.microsoft.com/Sites.Read.All',
];

// ── Token cache ──────────────────────────────────────────────────────────────

class FileTokenCache {
  constructor(cachePath) {
    this.path = cachePath;
  }
  async beforeCacheAccess(ctx) {
    // Always read from disk so tokens written after server start are visible
    const data = fs.existsSync(this.path) ? fs.readFileSync(this.path, 'utf8') : '';
    ctx.tokenCache.deserialize(data);
  }
  async afterCacheAccess(ctx) {
    if (ctx.cacheHasChanged) {
      fs.writeFileSync(this.path, ctx.tokenCache.serialize(), 'utf8');
    }
  }
}

const msalApp = new PublicClientApplication({
  auth: { clientId: CLIENT_ID, authority: `https://login.microsoftonline.com/${TENANT_ID}` },
  cache: { cachePlugin: new FileTokenCache(TOKEN_CACHE_PATH) },
  // Redirect ALL MSAL logs to stderr — stdout is reserved for MCP JSON-RPC
  system: {
    loggerOptions: {
      loggerCallback(_level, message, _containsPii) {
        process.stderr.write(`[msal] ${message}\n`);
      },
      logLevel: LogLevel.Warning,
      piiLoggingEnabled: false,
    },
  },
});

async function getAccessToken() {
  process.stderr.write('[m365-mcp] getAccessToken start\n');
  const accounts = await msalApp.getTokenCache().getAllAccounts();
  process.stderr.write(`[m365-mcp] accounts: ${accounts.length}\n`);
  if (accounts.length > 0) {
    try {
      const r = await msalApp.acquireTokenSilent({ scopes: SCOPES, account: accounts[0] });
      process.stderr.write('[m365-mcp] silent ok\n');
      return r.accessToken;
    } catch (e) {
      process.stderr.write(`[m365-mcp] silent failed: ${e.message}\n`);
    }
  }
  // Never run device-code flow inside an MCP subprocess — it would block forever.
  // Run `node auth.js` once to authenticate, then restart Claude Code.
  throw new Error(
    'Not authenticated. Run `node auth.js` in the m365-mcp repo directory to log in, then restart Claude Code.'
  );
}

function graph(token) {
  return Client.init({ authProvider: (done) => done(null, token) });
}

// ── MCP Server ───────────────────────────────────────────────────────────────

const server = new McpServer({ name: 'm365-mcp', version: '0.1.0' });

server.registerTool('list_messages', {
  description: 'List emails in a mail folder (default: inbox), newest first.',
  inputSchema: {
    folder: z.string().optional().describe('Folder ID or well-known name: inbox, drafts, sentitems, deleteditems. Default: inbox'),
    top:    z.number().min(1).max(50).optional().describe('Max messages (1-50). Default: 20'),
    filter: z.string().optional().describe('OData filter, e.g. "isRead eq false"'),
    search: z.string().optional().describe('Keyword search in subject/body/from'),
  },
}, async ({ folder = 'inbox', top = 20, filter, search }) => {
  process.stderr.write('[m365-mcp] list_messages called\n');
  const token = await getAccessToken();
  process.stderr.write('[m365-mcp] got token, calling graph\n');
  let req = graph(token).api(`/me/mailFolders/${folder}/messages`)
    .select('id,subject,from,receivedDateTime,isRead,bodyPreview,hasAttachments')
    .top(top);
  if (search) {
    req = req.search(`"${search}"`);  // $search is incompatible with $orderby
  } else {
    req = req.orderby('receivedDateTime desc');
    if (filter) req = req.filter(filter);
  }
  const res = await req.get();
  const msgs = res.value.map(m => ({
    id: m.id, subject: m.subject,
    from: m.from?.emailAddress,
    received: m.receivedDateTime,
    isRead: m.isRead,
    preview: m.bodyPreview,
    hasAttachments: m.hasAttachments,
  }));
  return { content: [{ type: 'text', text: JSON.stringify(msgs, null, 2) }] };
});

server.registerTool('get_message', {
  description: 'Get full content of a single email by ID.',
  inputSchema: { id: z.string().describe('Message ID') },
}, async ({ id }) => {
  const token = await getAccessToken();
  const m = await graph(token).api(`/me/messages/${id}`)
    .select('id,subject,from,toRecipients,ccRecipients,receivedDateTime,body,isRead').get();
  return { content: [{ type: 'text', text: JSON.stringify({
    id: m.id, subject: m.subject,
    from: m.from?.emailAddress,
    to: m.toRecipients?.map(r => r.emailAddress),
    cc: m.ccRecipients?.map(r => r.emailAddress),
    received: m.receivedDateTime, isRead: m.isRead,
    body: m.body?.content, bodyType: m.body?.contentType,
  }, null, 2) }] };
});

server.registerTool('list_attachments', {
  description: 'List attachments of an email message.',
  inputSchema: { id: z.string().describe('Message ID') },
}, async ({ id }) => {
  const token = await getAccessToken();
  const res = await graph(token).api(`/me/messages/${id}/attachments`)
    .select('id,name,contentType,size,isInline').get();
  const attachments = res.value.map(a => ({
    id: a.id, name: a.name, contentType: a.contentType,
    size: a.size, isInline: a.isInline,
  }));
  return { content: [{ type: 'text', text: JSON.stringify(attachments, null, 2) }] };
});

server.registerTool('get_attachment', {
  description: 'Get the content of an email attachment. Text-based files are returned as UTF-8 text; binary files are returned as base64 string in the "content" field with "encoding": "base64".',
  inputSchema: {
    messageId:    z.string().describe('Message ID'),
    attachmentId: z.string().describe('Attachment ID from list_attachments'),
  },
}, async ({ messageId, attachmentId }) => {
  const token = await getAccessToken();
  // First fetch without expand to detect attachment type
  const a = await graph(token).api(`/me/messages/${messageId}/attachments/${attachmentId}`).get();
  // Handle item attachments (embedded emails) — no contentBytes
  if (a['@odata.type'] === '#microsoft.graph.itemAttachment') {
    const expanded = await graph(token).api(
      `/me/messages/${messageId}/attachments/${attachmentId}?$expand=microsoft.graph.itemAttachment/item`
    ).get();
    return { content: [{ type: 'text', text: JSON.stringify({
      attachmentType: 'itemAttachment',
      name: expanded.name,
      item: expanded.item,
    }, null, 2) }] };
  }
  const raw = a.contentBytes;
  if (!raw) {
    return { content: [{ type: 'text', text: JSON.stringify({ name: a.name, contentType: a.contentType, size: a.size, error: 'no content' }, null, 2) }] };
  }
  const isText = a.contentType && (
    a.contentType.startsWith('text/') ||
    a.contentType.includes('json') ||
    a.contentType.includes('xml') ||
    a.contentType.includes('csv')
  );
  // Normalize to true base64 string regardless of what the SDK returns
  const b64 = Buffer.isBuffer(raw)
    ? raw.toString('base64')
    : Buffer.from(raw, 'binary').toString('base64');
  if (isText) {
    const text = Buffer.from(b64, 'base64').toString('utf8');
    return { content: [{ type: 'text', text }] };
  }
  return { content: [{ type: 'text', text: JSON.stringify({
    name: a.name, contentType: a.contentType, size: a.size,
    encoding: 'base64',
    content: b64,
  }, null, 2) }] };
});

server.registerTool('create_draft', {
  description: 'Create a new draft email. Does not send it.',
  inputSchema: {
    subject:  z.string().describe('Email subject'),
    to:       z.array(z.string()).optional().describe('Recipient email addresses'),
    cc:       z.array(z.string()).optional().describe('CC email addresses'),
    body:     z.string().optional().describe('Email body content'),
    bodyType: z.enum(['text', 'html']).optional().describe('Body format. Default: text'),
  },
}, async ({ subject, to = [], cc = [], body = '', bodyType = 'text' }) => {
  const token = await getAccessToken();
  const draft = await graph(token).api('/me/messages').post({
    subject,
    body: { contentType: bodyType, content: body },
    toRecipients: to.map(addr => ({ emailAddress: { address: addr } })),
    ccRecipients: cc.map(addr => ({ emailAddress: { address: addr } })),
  });
  return { content: [{ type: 'text', text: JSON.stringify({ id: draft.id, subject: draft.subject, webLink: draft.webLink }, null, 2) }] };
});

server.registerTool('create_reply_draft', {
  description: 'Create a reply draft to an existing message. Preserves reply headers and quoted body.',
  inputSchema: {
    id:      z.string().describe('Message ID to reply to'),
    body:     z.string().optional().describe('Reply body text (plain text). To use HTML, call update_draft after.'),
    replyAll: z.boolean().optional().describe('Reply to all recipients. Default: false'),
  },
}, async ({ id, body, replyAll = false }) => {
  const token = await getAccessToken();
  const endpoint = replyAll ? `/me/messages/${id}/createReplyAll` : `/me/messages/${id}/createReply`;
  const payload = body !== undefined ? { message: {}, comment: body } : {};
  const draft = await graph(token).api(endpoint).post(payload);
  return { content: [{ type: 'text', text: JSON.stringify({ id: draft.id, subject: draft.subject, webLink: draft.webLink }, null, 2) }] };
});

server.registerTool('update_draft', {
  description: 'Update subject, recipients, or body of an existing draft.',
  inputSchema: {
    id:       z.string().describe('Draft message ID'),
    subject:  z.string().optional(),
    to:       z.array(z.string()).optional(),
    cc:       z.array(z.string()).optional(),
    body:     z.string().optional(),
    bodyType: z.enum(['text', 'html']).optional(),
  },
}, async ({ id, subject, to, cc, body, bodyType }) => {
  const token = await getAccessToken();
  const patch = {};
  if (subject !== undefined) patch.subject = subject;
  if (body !== undefined)    patch.body = { contentType: bodyType || 'text', content: body };
  if (to !== undefined)      patch.toRecipients = to.map(addr => ({ emailAddress: { address: addr } }));
  if (cc !== undefined)      patch.ccRecipients = cc.map(addr => ({ emailAddress: { address: addr } }));
  const result = await graph(token).api(`/me/messages/${id}`).patch(patch);
  return { content: [{ type: 'text', text: JSON.stringify({ id: result.id, updated: true }, null, 2) }] };
});

server.registerTool('send_draft', {
  description: 'Send an existing draft email.',
  inputSchema: { id: z.string().describe('Draft message ID to send') },
}, async ({ id }) => {
  const token = await getAccessToken();
  await graph(token).api(`/me/messages/${id}/send`).post({});
  return { content: [{ type: 'text', text: JSON.stringify({ sent: true }, null, 2) }] };
});

server.registerTool('move_message', {
  description: 'Move an email to a different folder.',
  inputSchema: {
    id:                z.string().describe('Message ID'),
    destinationFolder: z.string().describe('Target folder ID or well-known name: inbox, drafts, deleteditems, junkemail'),
  },
}, async ({ id, destinationFolder }) => {
  const token = await getAccessToken();
  const result = await graph(token).api(`/me/messages/${id}/move`).post({ destinationId: destinationFolder });
  return { content: [{ type: 'text', text: JSON.stringify({ newId: result.id }, null, 2) }] };
});

server.registerTool('list_folders', {
  description: 'List all mail folders with total and unread counts.',
  inputSchema: {},
}, async () => {
  const token = await getAccessToken();
  const res = await graph(token).api('/me/mailFolders')
    .select('id,displayName,totalItemCount,unreadItemCount').get();
  const folders = res.value.map(f => ({ id: f.id, name: f.displayName, total: f.totalItemCount, unread: f.unreadItemCount }));
  return { content: [{ type: 'text', text: JSON.stringify(folders, null, 2) }] };
});

server.registerTool('create_folder', {
  description: 'Create a new mail folder.',
  inputSchema: {
    name:           z.string().describe('Folder display name'),
    parentFolderId: z.string().optional().describe('Parent folder ID (omit for top-level folder)'),
  },
}, async ({ name, parentFolderId }) => {
  const token = await getAccessToken();
  const endpoint = parentFolderId ? `/me/mailFolders/${parentFolderId}/childFolders` : '/me/mailFolders';
  const folder = await graph(token).api(endpoint).post({ displayName: name });
  return { content: [{ type: 'text', text: JSON.stringify({ id: folder.id, name: folder.displayName }, null, 2) }] };
});

// ── Calendar ─────────────────────────────────────────────────────────────────

server.registerTool('list_events', {
  description: 'List calendar events. Defaults to upcoming 7 days.',
  inputSchema: {
    startDateTime: z.string().optional().describe('ISO 8601 start (default: now)'),
    endDateTime:   z.string().optional().describe('ISO 8601 end (default: 7 days from now)'),
    top:           z.number().min(1).max(50).optional().describe('Max events (default: 20)'),
    calendarId:    z.string().optional().describe('Calendar ID (omit for primary calendar)'),
  },
}, async ({ startDateTime, endDateTime, top = 20, calendarId }) => {
  const token = await getAccessToken();
  const now = new Date();
  const start = startDateTime || now.toISOString();
  const end = endDateTime || new Date(now.getTime() + 7 * 24 * 60 * 60 * 1000).toISOString();
  const base = calendarId ? `/me/calendars/${calendarId}/calendarView` : '/me/calendarView';
  const res = await graph(token).api(base)
    .query({ startDateTime: start, endDateTime: end })
    .select('id,subject,start,end,location,organizer,isAllDay,bodyPreview,onlineMeeting')
    .top(top).orderby('start/dateTime').get();
  const events = res.value.map(e => ({
    id: e.id,
    subject: e.subject,
    start: e.start,
    end: e.end,
    location: e.location?.displayName,
    organizer: e.organizer?.emailAddress,
    isAllDay: e.isAllDay,
    preview: e.bodyPreview,
    teamsLink: e.onlineMeeting?.joinUrl,
  }));
  return { content: [{ type: 'text', text: JSON.stringify(events, null, 2) }] };
});

server.registerTool('create_event', {
  description: 'Create a calendar event.',
  inputSchema: {
    subject:       z.string().describe('Event title'),
    startDateTime: z.string().describe('ISO 8601 start time, e.g. 2026-06-20T10:00:00'),
    endDateTime:   z.string().describe('ISO 8601 end time'),
    timeZone:      z.string().optional().describe('IANA time zone (default: Asia/Taipei)'),
    location:      z.string().optional().describe('Location name'),
    body:          z.string().optional().describe('Event description'),
    attendees:     z.array(z.string()).optional().describe('Attendee email addresses'),
    isOnline:      z.boolean().optional().describe('Create Teams meeting link'),
  },
}, async ({ subject, startDateTime, endDateTime, timeZone = 'Asia/Taipei', location, body, attendees = [], isOnline }) => {
  const token = await getAccessToken();
  const payload = {
    subject,
    start: { dateTime: startDateTime, timeZone },
    end:   { dateTime: endDateTime,   timeZone },
    ...(location && { location: { displayName: location } }),
    ...(body && { body: { contentType: 'text', content: body } }),
    ...(attendees.length && { attendees: attendees.map(a => ({ emailAddress: { address: a }, type: 'required' })) }),
    ...(isOnline && { isOnlineMeeting: true, onlineMeetingProvider: 'teamsForBusiness' }),
  };
  const event = await graph(token).api('/me/events').post(payload);
  return { content: [{ type: 'text', text: JSON.stringify({
    id: event.id, subject: event.subject,
    start: event.start, end: event.end,
    teamsLink: event.onlineMeeting?.joinUrl,
  }, null, 2) }] };
});

// ── Teams ─────────────────────────────────────────────────────────────────────

server.registerTool('list_teams', {
  description: 'List Teams the current user is a member of.',
  inputSchema: {},
}, async () => {
  const token = await getAccessToken();
  const res = await graph(token).api('/me/joinedTeams')
    .select('id,displayName,description').get();
  const teams = res.value.map(t => ({ id: t.id, name: t.displayName, description: t.description }));
  return { content: [{ type: 'text', text: JSON.stringify(teams, null, 2) }] };
});

server.registerTool('list_channels', {
  description: 'List channels in a Team.',
  inputSchema: { teamId: z.string().describe('Team ID from list_teams') },
}, async ({ teamId }) => {
  const token = await getAccessToken();
  const res = await graph(token).api(`/teams/${teamId}/channels`)
    .select('id,displayName,description').get();
  const channels = res.value.map(c => ({ id: c.id, name: c.displayName, description: c.description }));
  return { content: [{ type: 'text', text: JSON.stringify(channels, null, 2) }] };
});

server.registerTool('list_channel_messages', {
  description: 'List recent messages in a Teams channel.',
  inputSchema: {
    teamId:    z.string().describe('Team ID'),
    channelId: z.string().describe('Channel ID from list_channels'),
    top:       z.number().min(1).max(50).optional().describe('Max messages (default: 20)'),
  },
}, async ({ teamId, channelId, top = 20 }) => {
  const token = await getAccessToken();
  const res = await graph(token).api(`/teams/${teamId}/channels/${channelId}/messages`)
    .top(top).get();
  const msgs = res.value.map(m => ({
    id: m.id,
    from: m.from?.user?.displayName,
    createdAt: m.createdDateTime,
    body: m.body?.content?.replace(/<[^>]+>/g, '').trim(),
  }));
  return { content: [{ type: 'text', text: JSON.stringify(msgs, null, 2) }] };
});

server.registerTool('list_chats', {
  description: 'List recent Teams chats (1-on-1 and group chats).',
  inputSchema: { top: z.number().min(1).max(50).optional().describe('Max chats (default: 20)') },
}, async ({ top = 20 }) => {
  const token = await getAccessToken();
  const res = await graph(token).api('/me/chats')
    .expand('members').top(top).get();
  const chats = res.value.map(c => ({
    id: c.id,
    topic: c.topic,
    chatType: c.chatType,
    members: c.members?.map(m => m.displayName),
  }));
  return { content: [{ type: 'text', text: JSON.stringify(chats, null, 2) }] };
});

server.registerTool('list_chat_messages', {
  description: 'List recent messages in a Teams chat.',
  inputSchema: {
    chatId: z.string().describe('Chat ID from list_chats'),
    top:    z.number().min(1).max(50).optional().describe('Max messages (default: 20)'),
  },
}, async ({ chatId, top = 20 }) => {
  const token = await getAccessToken();
  const res = await graph(token).api(`/me/chats/${chatId}/messages`).top(top).get();
  const msgs = res.value.map(m => ({
    id: m.id,
    from: m.from?.user?.displayName,
    createdAt: m.createdDateTime,
    body: m.body?.content?.replace(/<[^>]+>/g, '').trim(),
  }));
  return { content: [{ type: 'text', text: JSON.stringify(msgs, null, 2) }] };
});

server.registerTool('send_chat_message', {
  description: 'Send a message to a Teams chat (1-on-1 or group).',
  inputSchema: {
    chatId:  z.string().describe('Chat ID from list_chats'),
    message: z.string().describe('Message text to send'),
  },
}, async ({ chatId, message }) => {
  const token = await getAccessToken();
  const result = await graph(token).api(`/chats/${chatId}/messages`).post({
    body: { contentType: 'text', content: message },
  });
  return { content: [{ type: 'text', text: JSON.stringify({ id: result.id, createdAt: result.createdDateTime, sent: true }, null, 2) }] };
});

// ── OneDrive ──────────────────────────────────────────────────────────────────

server.registerTool('list_drive_items', {
  description: 'List files and folders in the user\'s OneDrive. Defaults to root.',
  inputSchema: {
    folderId: z.string().optional().describe('Folder item ID (omit for root)'),
    top:      z.number().min(1).max(100).optional().describe('Max items (default: 30)'),
  },
}, async ({ folderId, top = 30 }) => {
  const token = await getAccessToken();
  const endpoint = folderId
    ? `/me/drive/items/${folderId}/children`
    : '/me/drive/root/children';
  const res = await graph(token).api(endpoint)
    .select('id,name,size,createdDateTime,lastModifiedDateTime,webUrl,folder,file')
    .top(top).get();
  const items = res.value.map(i => ({
    id: i.id,
    name: i.name,
    type: i.folder ? 'folder' : 'file',
    size: i.size,
    modified: i.lastModifiedDateTime,
    webUrl: i.webUrl,
    mimeType: i.file?.mimeType,
  }));
  return { content: [{ type: 'text', text: JSON.stringify(items, null, 2) }] };
});

server.registerTool('get_drive_item_content', {
  description: 'Get the text content of a OneDrive file. Works best for plain text, markdown, and similar formats.',
  inputSchema: {
    itemId: z.string().describe('File item ID from list_drive_items'),
  },
}, async ({ itemId }) => {
  const token = await getAccessToken();
  const info = await graph(token).api(`/me/drive/items/${itemId}`)
    .select('id,name,size,file,webUrl').get();
  if (!info.file) {
    return { content: [{ type: 'text', text: JSON.stringify({ error: 'Item is a folder, not a file' }) }] };
  }
  // Download content via /content endpoint
  const content = await graph(token).api(`/me/drive/items/${itemId}/content`).getStream();
  const chunks = [];
  for await (const chunk of content) chunks.push(chunk);
  const text = Buffer.concat(chunks).toString('utf8');
  return { content: [{ type: 'text', text: text.slice(0, 50000) }] }; // cap at 50k chars
});

server.registerTool('search_drive', {
  description: 'Search files across OneDrive and SharePoint by name or content.',
  inputSchema: {
    query: z.string().describe('Search query'),
    top:   z.number().min(1).max(50).optional().describe('Max results (default: 20)'),
  },
}, async ({ query, top = 20 }) => {
  const token = await getAccessToken();
  const res = await graph(token).api(`/me/drive/root/search(q='${encodeURIComponent(query)}')`)
    .select('id,name,size,webUrl,lastModifiedDateTime,file,parentReference')
    .top(top).get();
  const items = res.value.map(i => ({
    id: i.id,
    name: i.name,
    type: i.file ? 'file' : 'folder',
    path: i.parentReference?.path,
    modified: i.lastModifiedDateTime,
    webUrl: i.webUrl,
    mimeType: i.file?.mimeType,
  }));
  return { content: [{ type: 'text', text: JSON.stringify(items, null, 2) }] };
});

// ── SharePoint ────────────────────────────────────────────────────────────────

server.registerTool('list_sites', {
  description: 'List SharePoint sites the user has access to.',
  inputSchema: {
    search: z.string().optional().describe('Filter by site name (uses $search)'),
    top:    z.number().min(1).max(50).optional().describe('Max results (default: 20)'),
  },
}, async ({ search, top = 20 }) => {
  const token = await getAccessToken();
  let req = graph(token).api('/sites')
    .select('id,displayName,webUrl,description').top(top);
  if (search) req = req.search(search);
  const res = await req.get();
  const sites = res.value.map(s => ({ id: s.id, name: s.displayName, webUrl: s.webUrl, description: s.description }));
  return { content: [{ type: 'text', text: JSON.stringify(sites, null, 2) }] };
});

server.registerTool('list_site_drives', {
  description: 'List document libraries (drives) in a SharePoint site.',
  inputSchema: {
    siteId: z.string().describe('Site ID from list_sites'),
  },
}, async ({ siteId }) => {
  const token = await getAccessToken();
  const res = await graph(token).api(`/sites/${siteId}/drives`)
    .select('id,name,description,webUrl').get();
  const drives = res.value.map(d => ({ id: d.id, name: d.name, description: d.description, webUrl: d.webUrl }));
  return { content: [{ type: 'text', text: JSON.stringify(drives, null, 2) }] };
});

server.registerTool('list_site_drive_items', {
  description: 'List files and folders in a SharePoint document library.',
  inputSchema: {
    siteId:   z.string().describe('Site ID from list_sites'),
    driveId:  z.string().describe('Drive ID from list_site_drives'),
    folderId: z.string().optional().describe('Folder item ID (omit for root)'),
    top:      z.number().min(1).max(100).optional().describe('Max items (default: 30)'),
  },
}, async ({ siteId, driveId, folderId, top = 30 }) => {
  const token = await getAccessToken();
  const endpoint = folderId
    ? `/sites/${siteId}/drives/${driveId}/items/${folderId}/children`
    : `/sites/${siteId}/drives/${driveId}/root/children`;
  const res = await graph(token).api(endpoint)
    .select('id,name,size,lastModifiedDateTime,webUrl,folder,file').top(top).get();
  const items = res.value.map(i => ({
    id: i.id,
    name: i.name,
    type: i.folder ? 'folder' : 'file',
    size: i.size,
    modified: i.lastModifiedDateTime,
    webUrl: i.webUrl,
    mimeType: i.file?.mimeType,
  }));
  return { content: [{ type: 'text', text: JSON.stringify(items, null, 2) }] };
});

const transport = new StdioServerTransport();
await server.connect(transport);
