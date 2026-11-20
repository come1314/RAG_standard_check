import fs from "node:fs/promises";
import path from "node:path";
import { RecursiveCharacterTextSplitter } from "@langchain/textsplitters";
import { DocxLoader } from "@langchain/community/document_loaders/fs/docx";
import { LanceDB } from "@langchain/community/vectorstores/lancedb";
import { Document } from "@langchain/core/documents";
import { RAG, PATHS, TABLE_NAME, isSupportedFile } from "./config.js";
import { getEmbeddings } from "./embeddings.js";

/** Extensions we know how to read out of `knowledgeFile/` live in config.ts,
 *  shared with the upload endpoint so the two can never disagree. */

export interface FileFingerprint {
  name: string;
  mtimeMs: number;
  size: number;
}

export interface IndexConfig {
  embeddingModel: string;
  queryPrefix: string;
  passagePrefix: string;
  chunkSize: number;
  chunkOverlap: number;
  /** Whether chunks were stored with a `《文档》 · 小节标题` prefix. */
  contextPrefix: boolean;
}

export interface IngestResult {
  files: string[];
  fingerprints: FileFingerprint[];
  chunkCount: number;
  embeddingModel: string;
  config: IndexConfig;
  durationMs: number;
  builtAt: string;
}

export interface LoadedFile {
  name: string;
  chunks: number;
}

/**
 * Splitter tuned for Chinese prose: sentence-ending punctuation is listed
 * before the generic whitespace separators so chunks break on natural
 * boundaries instead of mid-sentence.
 */
function createSplitter(): RecursiveCharacterTextSplitter {
  return new RecursiveCharacterTextSplitter({
    chunkSize: RAG.chunkSize,
    chunkOverlap: RAG.chunkOverlap,
    separators: ["\n\n", "\n", "。", "！", "？", "；", "…", "，", "、", " ", ""],
  });
}

/** Every readable document currently sitting in `knowledgeFile/`. */
export async function listKnowledgeFiles(): Promise<string[]> {
  let entries;
  try {
    entries = await fs.readdir(PATHS.knowledgeDir, { withFileTypes: true });
  } catch {
    return [];
  }
  return entries
    .filter((entry) => entry.isFile() && isSupportedFile(entry.name))
    .map((entry) => entry.name)
    .sort();
}

interface Heading {
  at: number;
  text: string;
}

/**
 * Section headings with their character offset, used to label each chunk with
 * the section it belongs to. Only short lines that look like a numbering
 * (`2.2 异常处理`) or `第X章/节` are treated as headings — a normal sentence
 * that happens to start with a digit would otherwise mislabel everything after
 * it.
 */
function headingsOf(raw: string): Heading[] {
  const out: Heading[] = [];
  let pos = 0;
  for (const line of raw.split("\n")) {
    const text = line.trim();
    const looksLikeHeading =
      text.length > 0 &&
      text.length <= 40 &&
      (/^\d+(\.\d+){0,3}[\s、.．]/.test(text) || /^第[一二三四五六七八九十百]+[章节部分]/.test(text));
    if (looksLikeHeading) out.push({ at: pos, text });
    pos += line.length + 1;
  }
  return out;
}

function titleOf(raw: string, fallback: string): string {
  const first = raw.split("\n").map((line) => line.trim()).find(Boolean) ?? "";
  return first.slice(0, 24) || fallback.replace(/\.[^.]+$/, "");
}

/** The heading in effect at `offset`, or undefined if none precedes it. */
function headingAt(headings: Heading[], offset: number): Heading | undefined {
  for (let i = headings.length - 1; i >= 0; i -= 1) {
    if (headings[i].at <= offset) return headings[i];
  }
  return undefined;
}

/**
 * Prepend `《文档名》 · 小节标题` to a chunk before it is embedded.
 *
 * A bare chunk carries no clue which document or section it came from, which
 * is how a passage from an unrelated file outscored the section that actually
 * answered the question. Matching on a 60-character probe rather than the whole
 * chunk keeps this cheap; when the probe isn't found we still label the
 * document, which is the part that mattered most.
 */
function withContext(
  text: string,
  fullText: string,
  headings: Heading[],
  title: string,
): string {
  const probe = text.slice(0, 60);
  const offset = fullText.indexOf(probe);
  const heading = offset >= 0 ? headingAt(headings, offset) : undefined;
  return `《${title}》${heading ? ` · ${heading.text}` : ""}\n${text}`;
}

async function fingerprint(): Promise<FileFingerprint[]> {
  const names = await listKnowledgeFiles();
  const out: FileFingerprint[] = [];
  for (const name of names) {
    const stat = await fs.stat(path.join(PATHS.knowledgeDir, name));
    out.push({ name, mtimeMs: stat.mtimeMs, size: stat.size });
  }
  return out;
}

/**
 * Read every supported file and split it into overlapping chunks.
 *
 * Metadata is rebuilt as flat primitives (`source`, `chunk_index`) rather than
 * copied from the loader — LanceDB stores metadata as table columns, and the
 * nested objects the splitter produces don't map onto Arrow columns.
 */
export async function loadAndSplit(): Promise<{ docs: Document[]; loaded: LoadedFile[] }> {
  const files = await listKnowledgeFiles();
  const splitter = createSplitter();
  const docs: Document[] = [];
  const loaded: LoadedFile[] = [];

  for (const file of files) {
    const fullPath = path.join(PATHS.knowledgeDir, file);
    const raw = await new DocxLoader(fullPath, { type: "docx" }).load();

    // The un-split text is what lets each chunk be matched back to its section
    // heading; the loader hands back one document per file.
    const fullText = raw.map((doc) => doc.pageContent).join("\n");
    const headings = headingsOf(fullText);
    const title = titleOf(fullText, file);

    const chunks = await splitter.splitDocuments(raw);

    chunks.forEach((chunk, index) => {
      const pageContent = RAG.contextPrefix
        ? withContext(chunk.pageContent, fullText, headings, title)
        : chunk.pageContent;

      docs.push(
        new Document({
          pageContent,
          metadata: { source: file, chunk_index: index },
        }),
      );
    });
    loaded.push({ name: file, chunks: chunks.length });
  }

  return { docs, loaded };
}

/** Index the whole knowledge directory, replacing whatever index existed. */
export async function ingest(): Promise<IngestResult> {
  const started = Date.now();
  const files = await listKnowledgeFiles();

  if (files.length === 0) {
    throw new Error(`No supported files found in ${PATHS.knowledgeDir}`);
  }

  console.log(`\n[ingest] reading ${files.length} file(s) from knowledgeFile/`);
  const { docs, loaded } = await loadAndSplit();

  for (const file of loaded) {
    console.log(`  · ${file.name}: ${file.chunks} chunks`);
  }

  if (docs.length === 0) {
    throw new Error("Documents loaded but produced no chunks.");
  }

  console.log(`[ingest] embedding ${docs.length} chunks locally…`);
  await fs.mkdir(PATHS.vectorDir, { recursive: true });

  // `mode: "overwrite"` recreates the table, so a rebuild never duplicates rows.
  await LanceDB.fromDocuments(docs, getEmbeddings(), {
    uri: PATHS.vectorDir,
    tableName: TABLE_NAME,
    mode: "overwrite",
  });

  const result: IngestResult = {
    files,
    fingerprints: await fingerprint(),
    chunkCount: docs.length,
    embeddingModel: RAG.embeddingModel,
    config: {
      embeddingModel: RAG.embeddingModel,
      queryPrefix: RAG.embeddingQueryPrefix,
      passagePrefix: RAG.embeddingPassagePrefix,
      chunkSize: RAG.chunkSize,
      chunkOverlap: RAG.chunkOverlap,
      contextPrefix: RAG.contextPrefix,
    },
    durationMs: Date.now() - started,
    builtAt: new Date().toISOString(),
  };

  await fs.mkdir(PATHS.dataDir, { recursive: true });
  await fs.writeFile(PATHS.metaFile, JSON.stringify(result, null, 2), "utf8");

  console.log(
    `[ingest] done — ${result.chunkCount} chunks in ${(result.durationMs / 1000).toFixed(1)}s`,
  );
  return result;
}

/** Read the sidecar written by the last successful ingest, if any. */
export async function readIndexMeta(): Promise<IngestResult | null> {
  try {
    const parsed = JSON.parse(await fs.readFile(PATHS.metaFile, "utf8"));
    return typeof parsed?.chunkCount === "number" ? (parsed as IngestResult) : null;
  } catch {
    return null;
  }
}

/**
 * The index is stale when the set of files on disk no longer matches what was
 * indexed — added, removed, or modified files all invalidate it.
 *
 * The index config is compared too, and that part matters more than it looks:
 * switching embedding models changes what the vectors mean, but the old vectors
 * are perfectly readable, so without this check a restart would silently search
 * a stale index built under a different model.
 */
export async function isIndexStale(): Promise<{ stale: boolean; reason: string }> {
  const meta = await readIndexMeta();
  if (!meta) return { stale: true, reason: "索引尚未构建" };

  const current: IndexConfig = {
    embeddingModel: RAG.embeddingModel,
    queryPrefix: RAG.embeddingQueryPrefix,
    passagePrefix: RAG.embeddingPassagePrefix,
    chunkSize: RAG.chunkSize,
    chunkOverlap: RAG.chunkOverlap,
    contextPrefix: RAG.contextPrefix,
  };
  const previous = meta.config;
  if (!previous) return { stale: true, reason: "索引缺少配置信息（旧版本构建）" };

  const configChanged = (Object.keys(current) as Array<keyof IndexConfig>).filter(
    (key) => current[key] !== previous[key],
  );
  if (configChanged.length > 0) {
    return { stale: true, reason: `索引配置已变更：${configChanged.join("、")}` };
  }

  const files = await fingerprint();
  const a = files.map((f) => `${f.name}:${f.size}:${Math.round(f.mtimeMs)}`).sort();
  const b = (meta.fingerprints ?? [])
    .map((f) => `${f.name}:${f.size}:${Math.round(f.mtimeMs)}`)
    .sort();

  if (a.length !== b.length) {
    return { stale: true, reason: `文件数量变化（索引 ${b.length} 个，当前 ${a.length} 个）` };
  }
  const changed = a.filter((entry, i) => entry !== b[i]);
  if (changed.length > 0) {
    return { stale: true, reason: `文件已变更：${changed.map((c) => c.split(":")[0]).join("、")}` };
  }
  return { stale: false, reason: "索引与 knowledgeFile/ 一致" };
}
