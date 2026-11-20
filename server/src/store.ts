import * as lancedb from "@lancedb/lancedb";
import { Document } from "@langchain/core/documents";
import { PATHS, TABLE_NAME, RAG } from "./config.js";
import { getEmbeddings } from "./embeddings.js";

export interface Hit {
  doc: Document;
  /** Cosine similarity in [-1, 1]. Higher means a closer match. */
  similarity: number;
}

/**
 * Columns LanceDB injects into every query result. Anything else is metadata
 * we wrote at ingest time.
 */
const RESERVED_COLUMNS = new Set([
  "vector",
  "text",
  "_distance",
  "_rowid",
  "_score",
  "score",
]);

/** True once a knowledge index has been built on disk. */
export async function indexExists(): Promise<boolean> {
  try {
    const db = await lancedb.connect(PATHS.vectorDir);
    return (await db.tableNames()).includes(TABLE_NAME);
  } catch {
    return false;
  }
}

/**
 * Embed the question and return the `k` closest chunks.
 *
 * We query LanceDB directly instead of going through the LangChain wrapper's
 * `similaritySearchWithScore`, because that wrapper reads a `score` column
 * while LanceDB actually emits `_distance` — so it hands back `undefined`.
 * Embeddings and the store are still LangChain's.
 */
export async function search(question: string, k: number = RAG.topK): Promise<Hit[]> {
  const db = await lancedb.connect(PATHS.vectorDir);
  const table = await db.openTable(TABLE_NAME);
  const vector = await getEmbeddings().embedQuery(question);

  const rows = await table
    .query()
    .nearestTo(vector)
    .distanceType("cosine")
    .limit(k)
    .toArray();

  return rows.map((row) => {
    const record = row as Record<string, unknown>;
    const distance = typeof record._distance === "number" ? record._distance : 1;

    const metadata: Record<string, string | number | boolean> = {};
    for (const [key, value] of Object.entries(record)) {
      if (RESERVED_COLUMNS.has(key)) continue;
      if (typeof value === "string" || typeof value === "number" || typeof value === "boolean") {
        metadata[key] = value;
      }
    }

    return {
      doc: new Document({
        pageContent: String(record.text ?? ""),
        metadata,
      }),
      // Cosine distance is 1 - cos(a, b), so similarity is just 1 - distance.
      similarity: Math.max(-1, Math.min(1, 1 - distance)),
    };
  });
}
