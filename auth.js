import { PublicClientApplication } from '@azure/msal-node';
import fs from 'fs';
import path from 'path';
import os from 'os';

const TENANT_ID = '3a4da202-0099-478c-929f-caa39aa19edb';
const CLIENT_ID = '7b6a5a61-c73d-47c9-8664-2a2cf0133c5b';
const TOKEN_CACHE_PATH = path.join(os.homedir(), '.outlook-mcp-token.json');
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
