import 'dotenv/config';
import { getAppOnlyTokenAsync, initializeGraphForAppOnlyAuth, allDriveFiles } from "./graphHelper";
import { Drive, DriveItem } from '@microsoft/microsoft-graph-types';
import { ResponseType } from '@microsoft/microsoft-graph-client';
import { Readable } from 'stream';
import { envOrFail, envOrFailArray } from '../utils/aux';

const drive_files: { driveId: string; item: DriveItem }[] = [];


async function main() {
  const { client } = await initializeGraphForAppOnlyAuth();
  const drive_names = envOrFailArray("DRIVES");
  const site_id = envOrFail("SITE_ID");
  console.log(site_id);
  const site = await client.api(`/sites/${site_id}`).get();
  if (!site || !site.id) {
    throw new Error(`Site not found with id ${site_id}`);
  }
  console.log('Found site:', site.webUrl, site.id);
  
  const drivesResponse: any = await client.api(`/sites/${site.id}/drives`).get();
  const drives = Array.isArray(drivesResponse?.value) ? (drivesResponse.value as Drive[]) : [];


  console.log(drive_names)
  for (const drive of drives) {
    console.log(`Drive: ${drive.name}`);
    if (drive.name && !(drive_names.includes(drive.name))) {
      continue
    }
    let i = 0;
    for await (const f of allDriveFiles(client, site.id, drive.id)) {
      drive_files.push({ driveId: drive.id!, item: f.item });
      console.log(i++);
    }
  }

  if (drive_files.length === 0) {
    console.log('No files found.');
    return;
  }

  
}

main().catch(e => console.error(e));