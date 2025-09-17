import { calculateHASH, calculateUUID, JurisprudenciaDocument, JurisprudenciaVersion, PartialJurisprudenciaDocument } from "@stjiris/jurisprudencia-document";
import { FileYield } from "./crawler";
import { client } from "../client";
import { IndexResponse, UpdateResponse } from "@elastic/elasticsearch/lib/api/types";
import { Client, ResponseType } from "@microsoft/microsoft-graph-client";
import mammoth from 'mammoth';

export async function updateJurisprudenciaDocumentFromSharepointFile(id: string, file: FileYield, graphClient: Client): Promise<UpdateResponse | undefined> {
  return;
}

export async function indexJurisprudenciaDocumentFromSharepointFile(file: FileYield, graphClient: Client): Promise<IndexResponse | undefined> {
    let obj = await createJurisprudenciaDocumentFromURL(file, graphClient);
    console.log("Object: ", obj);
    if (obj) {
        return client.index({
            index: JurisprudenciaVersion,
            body: obj
        })
    }
}

async function downloadDocxAsLines(
  client: Client,
  downloadURL: string
): Promise<string[]> {
  const arrayBuffer: ArrayBuffer = await client
    .api(downloadURL)
    .responseType(ResponseType.ARRAYBUFFER)
    .get();

  const buffer = Buffer.from(arrayBuffer);

  const result = await mammoth.extractRawText({ buffer });
  const text = result.value;

  const lines = text.split(/\r?\n/).map(line => line.trim()).filter(Boolean);

  return lines;
}

export async function createJurisprudenciaDocumentFromURL(file:FileYield, graphClient: Client): Promise<Partial<JurisprudenciaDocument> | undefined> {
  let Original: JurisprudenciaDocument["Original"] = {};
  let CONTENT: JurisprudenciaDocument["CONTENT"] = [];
  let numProc: JurisprudenciaDocument["Número de Processo"] = file.item.name;
  let DataAcordao: JurisprudenciaDocument["Data"] = null;

  if (file.pathSegments.length !== 0) {
    DataAcordao = file.pathSegments[1];
    // transform date from DD-MM-YYYY to DD/MM/YYYY
    DataAcordao = DataAcordao.replace(/(\d{2})\-(\d{2})\-(\d{4})/, "$1/$2/$3");
  }
  let Data: JurisprudenciaDocument["Data"] = DataAcordao || "01/01/1900";
  const todas_linhas = await downloadDocxAsLines(graphClient, file.downloadURL);
  for (const line of todas_linhas) {
    CONTENT.push(line);
  }

  let obj: PartialJurisprudenciaDocument = {
      "Original": Original,
      "CONTENT": CONTENT,
      "Data": Data,
      "Número de Processo": numProc,
      "Fonte": "STJ (Sharepoint)",
      "URL": file.item.webUrl,
      "Jurisprudência": { Index: ["Simples"], Original: ["Simples"], Show: ["Simples"] },
      "STATE": "privado",
  }
  
  obj.Texto = todas_linhas.map(line => `<p><font>${line}</font><br>`).join('');;
  
  addSeccaoAndArea(obj, file)
  obj["HASH"] = calculateHASH({
      ...obj,
      Original: obj.Original,
      "Número de Processo": obj["Número de Processo"] || "",
      Sumário: obj.Sumário || "",
      Texto: obj.Texto || "",
  })

  obj["UUID"] = calculateUUID(obj["HASH"])
  return obj;
}

export async function indexedUrlId(url: string): Promise<string | null> {
  return client.search({
          index: JurisprudenciaVersion,
          query: {
              term: {
                  "URL": url
              }
          },
          _source: false,
          size: 1
      }).then(r => r.hits.hits[0] ? r.hits.hits[0]._id : null)
  
}

function addSeccaoAndArea(obj: Partial<JurisprudenciaDocument>, file: FileYield) {

}

