import { initializeGraphForAppOnlyAuth, allDriveFiles, FileYield } from "./crawler";
import { Drive, DriveItem } from '@microsoft/microsoft-graph-types';
import { ResponseType } from '@microsoft/microsoft-graph-client';
import { Readable } from 'stream';
import { envOrFail, envOrFailArray } from '../utils/aux';
import { Report, report } from "../communication/report";
import { JurisprudenciaVersion } from '@stjiris/jurisprudencia-document';
import { client } from "../client";
import { WriteResponseBase } from '@elastic/elasticsearch/lib/api/types';
import { indexedUrlId, indexJurisprudenciaDocumentFromSharepointFile, updateJurisprudenciaDocumentFromSharepointFile } from './parser';

const FLAG_FULL_UPDATE = process.argv.some(arg => arg === "-f" || arg === "--full");

const FLAG_HELP = process.argv.some(arg => arg === "-h" || arg === "--help");

const all_files: FileYield[] = [];

function showHelp(code: number, error?: string) {
    if (error) {
        process.stderr.write(`Error: ${error}\n\n`);
    }
    process.stdout.write(`Usage: ${process.argv0} ${__filename} [OPTIONS]\n`)

    process.stdout.write(`Populate Jurisprudencia index. (${JurisprudenciaVersion})\n`)
    process.stdout.write(`Use ES_URL, ES_USER and ES_PASS environment variables to setup the elasticsearch client\n`)
    process.stdout.write(`Options:\n`)
    process.stdout.write(`\t--full, -f\tWork in progress. Should update every document already indexed and check if there are deletions\n`);
    process.stdout.write(`\t--help, -h\tshow this help\n`)
    process.exit(code);
}

async function main() {
  if (FLAG_HELP) return showHelp(0);
  let info: Report = {
          created: 0,
          dateEnd: new Date(),
          dateStart: new Date(),
          deleted: 0,
          skiped: 0,
          soft: !FLAG_FULL_UPDATE,
          target: JurisprudenciaVersion,
          updated: 0
      }

  process.once("SIGINT", () => {
      info.dateEnd = new Date();
      console.log("Terminado a pedido do utilizador");
      report(info).then(() => process.exit(0));
  })
  // if elastic search client doesn't exist fail
  /* let existsR = await client.indices.exists({ index: JurisprudenciaVersion }, { ignore: [404] });
  if (!existsR) {
      return showHelp(1, `${JurisprudenciaVersion} not found`);
  } */

  // setup microsoft graph client
  const graphClient = await initializeGraphForAppOnlyAuth();
  const drive_names = envOrFailArray("DRIVES");
  const site_id = envOrFail("SITE_ID");

  // connect to the site
  const site = await graphClient.api(`/sites/${site_id}`).get();
  if (!site || !site.id) {
    throw new Error(`Site not found with id ${site_id}`);
  }
  
  // retrieve the all drives
  const drivesResponse: any = await graphClient.api(`/sites/${site.id}/drives`).get();
  const drives = Array.isArray(drivesResponse?.value) ? (drivesResponse.value as Drive[]) : [];

  for (const drive of drives) {
    console.log(`Drive: ${drive.name}`);

    // only crawl selected drives
    if (drive.name && !(drive_names.includes(drive.name))) {
      continue
    }
    // this is just a dev feature to see how many files are found in the crawl
    let i = 0;
    // get all files
    for await (const f of allDriveFiles(graphClient, site.id, drive.id)) {
        let id = await indexedUrlId(f.downloadURL);
        i++;
        if (id && !FLAG_FULL_UPDATE) {
            info.skiped++;
            continue;
        };

        let r: WriteResponseBase | undefined = undefined;
        if (id) {
            r = await updateJurisprudenciaDocumentFromSharepointFile(id, f, graphClient);
        }
        else {
            r = await indexJurisprudenciaDocumentFromSharepointFile(f, graphClient);
        }
        switch (r?.result) {
            case "created":
                info.created++;
                break;
            case "deleted":
                info.deleted++;
                break;
            case "updated":
                info.updated++;
                break;
            case "noop":
            case "not_found":
            default:
                info.skiped++;
                break;
        }
    }

  }
  info.dateEnd = new Date()
  //await report(info)
}

main().catch(e => console.error(e));