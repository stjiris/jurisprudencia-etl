import 'dotenv/config';
import { getAppOnlyTokenAsync, initializeGraphForAppOnlyAuth, allDriveFiles } from "./graphHelper";
import { Drive, DriveItem } from '@microsoft/microsoft-graph-types';

async function main() {
  const { client } = await initializeGraphForAppOnlyAuth();

  const hostname = "stjpt.sharepoint.com/sites/FileShare";
  
  const init_url = "/sites/" + hostname
  console.log(init_url)
  const site_response = await client.api(`/sites?search=FileShare/`).get();
  if (!Array.isArray(site_response.value) || site_response.value.length === 0) {
    throw new Error('Site not found by search');
  }
  const site = site_response.value[0];
  console.log('Found site:', site.webUrl, site.id);
  
  const drivesResponse: any = await client.api(`/sites/${site.id}/drives`).get();
  const drives = Array.isArray(drivesResponse?.value) ? (drivesResponse.value as Drive[]) : [];
  const drive = drives[0]
  console.log(`Drive: ${drive.name}, id: ${drive.id}`);
  for await (const f of allDriveFiles(client, site.id, drive.id)) {
    console.log(f.item.name/*, f.downloadUrl*/);
  }
}

main().catch(e => console.error(e));