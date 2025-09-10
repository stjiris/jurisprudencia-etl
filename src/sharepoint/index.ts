import 'dotenv/config';
import { getAppOnlyTokenAsync, initializeGraphForAppOnlyAuth } from "./graphHelper";

async function main() {
  const { credential, client } = await initializeGraphForAppOnlyAuth();

  const token = await getAppOnlyTokenAsync(credential);
  console.log(`App-only token: ${token}`);
  console.log("here")
  const site = await client.api('/sites/stjpt.sharepoint.com').get();
  console.log('site id:', site.id);
  console.log('webUrl:', site.webUrl);
  
}

main().catch(e => console.error(e));