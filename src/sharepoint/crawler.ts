import 'dotenv/config';
import { getAppOnlyTokenAsync, initializeGraphForAppOnlyAuth, listAllItems } from "./graphHelper";
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
  
  const drivesResponse = await client.api(`/sites/${site.id}/drives`).get();

  (drivesResponse.value as Drive[]).forEach((d: Drive) => {
    console.log(`Drive: ${d.name}, id: ${d.id}`);
  });

  for (const drive of drivesResponse.value as Drive[]) {
    const itemsResponse = await client.api(`/sites/${site.id}/drives/${drive.id}/root/children`).get();
    console.log(`Drive name: ${drive.name}, id: ${drive.id}`);

    (itemsResponse.value as DriveItem[]).forEach((item: DriveItem) => {
      console.log(item.name, item.folder ? "Folder" : "File");
      listAllItems(client, site.id, drive.id, item.id)
    });
  }
}

main().catch(e => console.error(e));