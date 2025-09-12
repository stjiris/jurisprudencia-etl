import { ClientSecretCredential } from '@azure/identity';
import { Client, PageCollection } from '@microsoft/microsoft-graph-client';
import { TokenCredentialAuthenticationProvider } from
  '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';
import { envOrFail } from '../utils/aux';
import { DriveItem } from '@microsoft/microsoft-graph-types';

export async function initializeGraphForAppOnlyAuth(): Promise<{
  client: Client;
}> {
  const tenantId = envOrFail('TENANT_ID');
  const clientId = envOrFail('CLIENT_ID');
  const clientSecret = envOrFail('CLIENT_SECRET');

  const credential = new ClientSecretCredential(tenantId, clientId, clientSecret);

  const authProvider = new TokenCredentialAuthenticationProvider(credential, {
    scopes: ['https://graph.microsoft.com/.default'],
  });

  const client = Client.initWithMiddleware({
    authProvider,
  });

  return { client };
}

export async function getAppOnlyTokenAsync(clientSecretCredential: ClientSecretCredential): Promise<string> {
  const response = await clientSecretCredential.getToken([
    'https://graph.microsoft.com/.default',
  ]);
  return response.token;
}

export type FileYield = {
  item: DriveItem;
  downloadUrl: string;
};

export async function* allDriveFiles(
  client: Client,
  siteId: string | undefined,
  driveId: string | undefined
): AsyncGenerator<FileYield, void, unknown> {
  if (!siteId) throw new Error('siteId required');
  if (!driveId) throw new Error('driveId required');

  const visited = new Set<string>();
  const folderStack: string[] = ['root'];

  async function* fetchChildren(folderId: string) {
    let url = `/sites/${siteId}/drives/${driveId}/items/${folderId}/children`;
    while (url) {
      const res: any = await client.api(url).get();
      const items = Array.isArray(res?.value) ? (res.value as DriveItem[]) : null;
      if (!items) {
        return [];
      }
      for (const it of items) {
        yield it;
      }
      url = res['@odata.nextLink'] ?? null;
    }
  }

  while (folderStack.length > 0) {
    const folderId = folderStack.shift()!;

    for await (const child of (async function* () { yield* fetchChildren(folderId); })()) {
      const key = child.id ?? child.webUrl ?? `${driveId}:${folderId}:${child.name}`;

      if (visited.has(key)) {
        continue;
      }
      visited.add(key);

      if (child.folder) {
        if (child.id) folderStack.push(child.id);
        continue;
      }
      const contentUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives/${driveId}/items/${child.id}/content`;
      yield {
        item: child,
        downloadUrl: contentUrl,
      };
    }
  }
}



