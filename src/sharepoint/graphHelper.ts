import { ClientSecretCredential } from '@azure/identity';
import { Client, PageCollection } from '@microsoft/microsoft-graph-client';
import { TokenCredentialAuthenticationProvider } from
  '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials';
import { envOrFail } from '../utils/aux';

export async function initializeGraphForAppOnlyAuth(): Promise<{
  credential: ClientSecretCredential;
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

  return { credential, client };
}

export async function getAppOnlyTokenAsync(clientSecretCredential: ClientSecretCredential): Promise<string> {
  const response = await clientSecretCredential.getToken([
    'https://graph.microsoft.com/.default',
  ]);
  return response.token;
}
