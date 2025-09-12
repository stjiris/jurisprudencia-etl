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

export async function listAllItems(client: Client, siteId: string, driveId: string = "", folderId: string = 'root') {
  const res = await client.api(`/sites/${siteId}/drives/${driveId}/items/${folderId}/children`).get();
  for (const item of res.value as DriveItem[]) {
    console.log(item.name, item.folder ? "Folder" : "File");

    if (item.folder) {
      await listAllItems(client, siteId, driveId, item.id);
    }
  }
}
