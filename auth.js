import { PublicClientApplication } from '@azure/msal-node';
import fs from 'fs';
import path from 'path';
import os from 'os';

const TENANT_ID = process.env.M365_MCP_TENANT_ID;
const CLIENT_ID = process.env.M365_MCP_CLIENT_ID;
if (!TENANT_ID || !CLIENT_ID) {
  console.error('ERROR: M365_MCP_TENANT_ID and M365_MCP_CLIENT_ID must be set');
  process.exit(1);
}
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

class FileTokenCache {
  constructor(p) {
    this.path = p;
    this._data = fs.existsSync(p) ? fs.readFileSync(p, 'utf8') : '';
  }
  async beforeCacheAccess(ctx) { ctx.tokenCache.deserialize(this._data); }
  async afterCacheAccess(ctx) {
    if (ctx.cacheHasChanged) {
      this._data = ctx.tokenCache.serialize();
      fs.writeFileSync(this.path, this._data, 'utf8');
    }
  }
}

const app = new PublicClientApplication({
  auth: { clientId: CLIENT_ID, authority: `https://login.microsoftonline.com/${TENANT_ID}` },
  cache: { cachePlugin: new FileTokenCache(TOKEN_CACHE_PATH) },
});

console.log('開始 Microsoft 帳號認證...\n');
const result = await app.acquireTokenByDeviceCode({
  scopes: SCOPES,
  deviceCodeCallback: (resp) => {
    console.log(`請開瀏覽器前往: ${resp.verificationUri}`);
    console.log(`輸入代碼: ${resp.userCode}\n`);
    console.log('（等待認證完成...）');
  },
});
console.log(`\n認證成功！帳號: ${result.account.username}`);
console.log(`Token 已存至: ${TOKEN_CACHE_PATH}`);
