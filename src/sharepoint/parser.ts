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

  if (file.pathSegments.length >= 2) {
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
    const Secções = {
        SECÇÃO_1: "1.ª Secção (Cível)",
        SECÇÃO_2: "2.ª Secção (Cível)",
        SECÇÃO_3: "3.ª Secção (Criminal)",
        SECÇÃO_4: "4.ª Secção (Social)",
        SECÇÃO_5: "5.ª Secção (Criminal)",
        SECÇÃO_6: "6.ª Secção (Cível)",
        SECÇÃO_7: "7.ª Secção (Cível)",
        CONTENCIOSO: "Contencioso",
        CONFLITOS: "Conflitos"
    };
    const Áreas = {
        SECÇÃO_1: "Área Cível",
        SECÇÃO_2: "Área Cível",
        SECÇÃO_3: "Área Criminal",
        SECÇÃO_4: "Área Social",
        SECÇÃO_5: "Área Criminal",
        SECÇÃO_6: "Área Cível",
        SECÇÃO_7: "Área Cível",
        CONTENCIOSO: "Contencioso",
        CONFLITOS: "Conflitos"
    }
    if (file.pathSegments.length === 0) {
        return;
    }

    const seg = (file.pathSegments[0] || '').trim();

    const match = seg.match(/(\d)|CONTENCIOSO/i);
    if (!match) return;

    const digitOrText = match[0].toUpperCase();

    if (digitOrText === "CONTENCIOSO" || digitOrText === "8") {
        const sec = Secções.CONTENCIOSO;
        const area = Áreas.CONTENCIOSO;
        obj.Secção = { Index: [sec], Original: [sec], Show: [sec] };
        obj.Área   = { Index: [area], Original: [area], Show: [area] };
    } else {
        const key = `SECÇÃO_${digitOrText}` as keyof typeof Secções;
        const sec = Secções[key];
        const area = Áreas[key];
        if (sec && area) {
            obj.Secção = { Index: [sec], Original: [sec], Show: [sec] };
            obj.Área   = { Index: [area], Original: [area], Show: [area] };
        }
    }

}

