import { JurisprudenciaDocument, JurisprudenciaVersion, PartialJurisprudenciaDocument } from "@stjiris/jurisprudencia-document";
import { FileYield } from "./crawler";
import { client } from "../client";
import { IndexResponse, UpdateResponse } from "@elastic/elasticsearch/lib/api/types";
import { Client } from "@microsoft/microsoft-graph-client";

export async function updateJurisprudenciaDocumentFromSharepointFile(id: string, file: FileYield, graphClient: Client): Promise<UpdateResponse | undefined> {
  return;
}

export async function indexJurisprudenciaDocumentFromSharepointFile(file:FileYield, graphClient: Client): Promise<IndexResponse | undefined> {
  let indexable_file = formatFileForJurisIndex(file, graphClient);
  console.log("this was the file contents")
  /* let Data: JurisprudenciaDocument["Data"] = DataToUse || DataAcordao || "01/01/1900";
  
  let obj: PartialJurisprudenciaDocument = {
      "Original": Original,
      "CONTENT": CONTENT,
      "Data": Data,
      "Número de Processo": numProc,
      "Fonte": "STJ (Sharepoint)",
      "URL": file.downloadURL,
      "Jurisprudência": { Index: ["Simples"], Original: ["Simples"], Show: ["Simples"] },
      "STATE": "privado",
  } */
 return;
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

function formatFileForJurisIndex(file: FileYield, graphClient: Client) {
  const fileContent = graphClient.api(file.downloadURL);
  console.log(fileContent);
}
