import path from "node:path";
import { fileURLToPath } from "node:url";
import dotenv from "dotenv";

const here = path.dirname(fileURLToPath(import.meta.url));

/** `server/` */
export const SERVER_ROOT = path.resolve(here, "..");
/** repository root — the parent of `server/` and `web/` */
export const PROJECT_ROOT = path.resolve(SERVER_ROOT, "..");

// Load `server/.env` explicitly so it works no matter which directory the
// process was launched from.
dotenv.config({ path: path.join(SERVER_ROOT, ".env") });

export const PATHS = {
  /** Where the source documents live. */
  knowledgeDir: path.join(PROJECT_ROOT, "knowledgeFile"),
  /** Everything we generate lives under `server/data/`. */
  dataDir: path.join(SERVER_ROOT, "data"),
  /** LanceDB database directory. */
  vectorDir: path.join(SERVER_ROOT, "data", "lancedb"),
  /** Sidecar file describing when the index was built and from what. */
  metaFile: path.join(SERVER_ROOT, "data", "index-meta.json"),
  /** HuggingFace download cache, so the embedding model is fetched once. */
  hfCache: path.join(SERVER_ROOT, "data", "hf-cache"),
};

function int(value: string | undefined, fallback: number): number {
  const parsed = Number(value);
  return Number.isInteger(parsed) && parsed > 0 ? parsed : fallback;
}

/** Which chat-model backend generates the answers. */
export type LlmProvider = "anthropic" | "openai";

function currentProvider(): string {
  return (process.env.LLM_PROVIDER ?? "anthropic").trim().toLowerCase();
}

function currentModel(): string {
  const provider = currentProvider();
  return (
    process.env.LLM_MODEL ?? (provider === "openai" ? "gpt-4o-mini" : "claude-opus-5-5")
  );
}

export const RAG = {
  /**
   * Retrieval-trained sentence model. E5 expects a role prefix on every input
   * (`query: ` for questions, `passage: ` for indexed chunks) — set the two
   * prefix vars to an empty string if you switch to a model that doesn't want
   * them (e.g. paraphrase-multilingual-MiniLM-L12-v2).
   */
  embeddingModel:
    process.env.EMBEDDING_MODEL ?? "Xenova/multilingual-e5-small",
  embeddingQueryPrefix: process.env.EMBEDDING_QUERY_PREFIX ?? "query: ",
  embeddingPassagePrefix: process.env.EMBEDDING_PASSAGE_PREFIX ?? "passage: ",
  /**
   * huggingface.co is unreachable from this network; hf-mirror.com is a full
   * mirror of it. Override with HF_ENDPOINT if you have direct access.
   */
  hfEndpoint: process.env.HF_ENDPOINT ?? "https://hf-mirror.com",
  /**
   * "anthropic" (Claude) or "openai" (any OpenAI-compatible endpoint:
   * DeepSeek, 通义千问, 智谱 GLM, Ollama, vLLM...).
   *
   * Both are getters, not plain values: `settings.ts` rewrites
   * LLM_PROVIDER/LLM_MODEL on every save, so reading them on access is what
   * lets the settings panel take effect without restarting the server.
   * Kept as strings rather than validated here so a typo can't stop
   * `npm run ingest`, which never calls a model — `llm.ts` rejects it on the
   * first chat request instead.
   */
  get llmProvider(): string {
    return currentProvider();
  },
  get llmModel(): string {
    return currentModel();
  },
  /**
   * Chunks returned per question. 250-char chunks are shorter than the old
   * 400, so 8 of them cost the same prompt as 5 did before.
   *
   * Measured on the bundled eval plus one hand-checked failure case: at
   * chunkSize 250 with context prefixes, a query about 异常处理 moves from
   * ranking the preface #1 to ranking 《Java开发规范 · 2.2 异常处理》#1, and
   * drops an unrelated document out of the results entirely.
   */
  topK: int(process.env.TOP_K, 8),
  /**
   * CHUNK_SIZE is in characters, but the embedding models cap at 512 *tokens*
   * (`max_position_embeddings`) and silently truncate past that — for Chinese
   * roughly one character per token, so 800-character chunks threw away the
   * tail of every chunk before it was ever embedded. 250 keeps every chunk well
   * inside the limit and makes each one topically homogeneous, which is what
   * stops a long generic passage (a preface) from outscoring the section that
   * actually answers the question.
   */
  chunkSize: int(process.env.CHUNK_SIZE, 250),
  chunkOverlap: int(process.env.CHUNK_OVERLAP, 50),
  /**
   * Prepend `《文档名》 · 小节标题` to every chunk before embedding. Without it
   * a chunk carries no clue which document or section it belongs to, so a
   * passage from an unrelated file can score higher than the right one.
   * Measured effect on the 异常处理 case: irrelevant-document hits 2 -> 0 and
   * on-topic chunks in the top 8 2 -> 5.
   */
  contextPrefix: (process.env.CHUNK_CONTEXT_PREFIX ?? "true").toLowerCase() !== "false",
  port: int(process.env.PORT, 3001),
  /**
   * Loopback by default — the settings endpoint accepts API keys over this
   * port, so it should not be reachable from the LAN unless you opt in.
   * Set HOST=0.0.0.0 to expose it.
   */
  host: process.env.HOST ?? "127.0.0.1",
};

// transformers.js reads HF_ENDPOINT while it loads, and that module is imported
// after this one — so publish the mirror endpoint before anything else can see it.
process.env.HF_ENDPOINT ??= RAG.hfEndpoint;

/** The LanceDB table holding our chunks. */
export const TABLE_NAME = "knowledge_chunks";

/** Document extensions the knowledge base can read, checked in one place so
 *  the loader and the upload endpoint can never disagree. */
export const SUPPORTED_EXTENSIONS = [".docx"] as const;

export function isSupportedFile(fileName: string): boolean {
  return (SUPPORTED_EXTENSIONS as readonly string[]).includes(
    path.extname(fileName).toLowerCase(),
  );
}
