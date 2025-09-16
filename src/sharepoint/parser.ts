/* import { ResponseType } from '@microsoft/microsoft-graph-client';
import { Drive, DriveItem } from '@microsoft/microsoft-graph-types';

async function download_local(drives: Drive[], files: { driveId: string; item: DriveItem }[]) {
    const { promises: fsp } = await import('fs');
  const path = await import('path');

  const driveNameById = new Map<string,string>();
  for (const d of drives) {
    if (d.id) driveNameById.set(d.id, d.name ?? d.id);
  }

  for (const f of files) {
    const driveId = f.driveId;
    const driveName = driveNameById.get(driveId) ?? driveId;
    const folderPath = path.join(process.cwd(), "drives", driveName);

    await fsp.mkdir(folderPath, { recursive: true });

    const filename = f.item.name ?? `${f.item.id}.bin`;
    const safeFilename = filename.replace(/[<>:"/\\|?*\x00-\x1F]/g, '_');
    const outPath = path.join(folderPath, safeFilename);

    try {
      await fsp.stat(outPath);
      console.log(`Skipping "${filename}" — already exists at ${outPath}`);
      continue;
    } catch {
    }

    console.log(`Downloading "${filename}" -> ${outPath}`);

    const contentPath = `/sites/${site.id}/drives/${driveId}/items/${f.item.id}/content`;
    const arrayBuffer = await client.api(contentPath).responseType(ResponseType.ARRAYBUFFER).get();
    const buffer = Buffer.from(arrayBuffer as ArrayBuffer);

    await fsp.writeFile(outPath, Readable.from(buffer));
  }

  console.log('All files downloaded.');
}
 */