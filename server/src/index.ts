import express, { type NextFunction, type Request, type Response } from "express";
import multer from "multer";
import { RAG, PATHS } from "./config.js";
import { hasLlmCredentials } from "./llm.js";
import { ingest, readIndexMeta, isIndexStale, listKnowledgeFiles } from "./ingest.js";
import { indexExists } from "./store.js";
import { answer, answerStream, type AbortFlag, type Turn } from "./answer.js";
import { saveUploads } from "./upload.js";
import { loadSettings, publicSettings, updateSettings, PROVIDERS, type PartialUpdate } from "./settings.js";

type Handler = (req: Request, res: Response, next: NextFunction) => Promise<void> | void;

const wrap =
  (handler: Handler) =>
  (req: Request, res: Response, next: NextFunction): void => {
    void Promise.resolve(handler(req, res, next)).catch(next);
  };

const app = express();
app.use(express.json({ limit: "1mb" }));

/**
 * Files are buffered in memory rather than streamed to disk first: they have
 * to be validated (extension, ZIP magic bytes, name collision) before a single
 * byte lands in `knowledgeFile/`, and the real documents here are ~1.4 MB.
 *
 * `defParamCharset` matters: multer defaults it to 'latin1', but browsers put
 * raw UTF-8 bytes in `filename="…"`. Left alone, a Chinese filename arrives as
 * mojibake ("中文文件名测试.docx" -> "ä¸­ææç¨…") and gets written to disk that
 * way. Verified against a WHATWG FormData upload.
 */
const upload = multer({
  storage: multer.memoryStorage(),
  limits: { fileSize: 15 * 1024 * 1024, files: 8 },
  defParamCharset: "utf8",
});

/** Tracks the background build so the UI can show progress and we can reject
 *  chat requests that arrive before there is anything to retrieve from. */
let building: Promise<void> | null = null;
let buildError: string | null = null;

function ensureIndex(): Promise<void> {
  if (building) return building;

  building = (async () => {
    buildError = null;
    try {
      const stale = await isIndexStale();
      const exists = await indexExists();
      if (!exists || stale.stale) {
        console.log(`[index] ${stale.reason} — rebuilding…`);
        await ingest();
      } else {
        console.log("[index] up to date, skipping rebuild");
      }
    } catch (error) {
      buildError = error instanceof Error ? error.message : String(error);
      console.error("[index] build failed:", buildError);
      throw error;
    } finally {
      building = null;
    }
  })();

  return building;
}

/**
 * Request body shared by the JSON and streaming chat endpoints.
 * Returns null when there is no usable question.
 */
function readChatRequest(body: unknown): { question: string; history: Turn[] } | null {
  const b = (body ?? {}) as { question?: unknown; history?: unknown };
  const question = typeof b.question === "string" ? b.question.trim() : "";
  if (!question) return null;

  const history: Turn[] = Array.isArray(b.history)
    ? (b.history as unknown[])
        .filter(
          (t): t is Turn =>
            Boolean(t) &&
            typeof (t as Turn).question === "string" &&
            typeof (t as Turn).answer === "string",
        )
        .slice(-6)
    : [];

  return { question, history };
}

/** Reason the store can't answer yet, or null when it can. */
async function readyForChat(): Promise<string | null> {
  if (building) return "知识库索引尚未就绪，请稍候重试";
  if (!(await indexExists())) return "知识库尚未构建，请先建立索引";
  return null;
}

app.get(
  "/api/status",
  wrap(async (_req, res) => {
    const [meta, stale, exists] = await Promise.all([
      readIndexMeta(),
      isIndexStale(),
      indexExists(),
    ]);

    res.json({
      // Usable as soon as a table exists. Staleness is reported separately and
      // shown as a warning — a file added mid-run shouldn't block questions that
      // the already-built index can answer.
      ready: exists && !building,
      indexing: Boolean(building),
      buildError,
      stale,
      files: await listKnowledgeFiles(),
      knowledgeDir: PATHS.knowledgeDir,
      chunkCount: meta?.chunkCount ?? 0,
      builtAt: meta?.builtAt ?? null,
      embeddingModel: RAG.embeddingModel,
      llmModel: RAG.llmModel,
      llmMode: hasLlmCredentials() ? "llm" : "extractive",
      llmProvider: RAG.llmProvider,
      topK: RAG.topK,
      chunkSize: RAG.chunkSize,
      chunkOverlap: RAG.chunkOverlap,
    });
  }),
);

app.post(
  "/api/ingest",
  wrap(async (_req, res) => {
    if (building) {
      res.status(409).json({ error: "索引正在构建中，请稍候" });
      return;
    }
    const result = await ingest();
    res.json(result);
  }),
);

app.get(
  "/api/settings",
  wrap(async (_req, res) => {
    res.json(await publicSettings());
  }),
);

/**
 * Partial update: any field left out of the body is left alone, which is what
 * lets the UI never echo back a key it was never shown.
 */
app.post(
  "/api/settings",
  wrap(async (req, res) => {
    const body = req.body ?? {};
    const update: PartialUpdate = {};

    if (typeof body.llmProvider === "string") update.llmProvider = body.llmProvider.trim();

    if (body.providers && typeof body.providers === "object") {
      const providers: Record<string, { model?: string; baseUrl?: string; apiKey?: string }> = {};
      for (const [name, value] of Object.entries(body.providers as Record<string, unknown>)) {
        if (!PROVIDERS.includes(name) || typeof value !== "object" || value === null) continue;
        const fields = value as Record<string, unknown>;
        const entry: { model?: string; baseUrl?: string; apiKey?: string } = {};
        // Every field is optional; a missing key means "keep what's stored".
        if (typeof fields.model === "string") entry.model = fields.model.trim();
        if (typeof fields.baseUrl === "string") entry.baseUrl = fields.baseUrl.trim();
        if (typeof fields.apiKey === "string") entry.apiKey = fields.apiKey; // not trimmed: keys are exact
        providers[name] = entry;
      }
      update.providers = providers;
    }

    if (update.llmProvider !== undefined && !PROVIDERS.includes(update.llmProvider)) {
      // Client input error, so 400 — not the 500 the generic handler would give.
      res.status(400).json({
        error: `不支持的 LLM_PROVIDER="${update.llmProvider}"，可选值：${PROVIDERS.join(" / ")}`,
      });
      return;
    }

    const saved = await updateSettings(update);
    res.json({
      ...(await publicSettings()),
      // Confirms it applied without restarting.
      applied: { llmProvider: saved.llmProvider, credentials: hasLlmCredentials() },
    });
  }),
);

/**
 * Multipart upload into `knowledgeFile/`, field name `files`.
 * Succeeded files make the index stale, so a rebuild is kicked off in the
 * background and the existing /api/status polling reports progress.
 */
app.post(
  "/api/upload",
  upload.array("files", 8),
  wrap(async (req, res) => {
    const files = (req.files as Express.Multer.File[] | undefined) ?? [];
    if (files.length === 0) {
      res.status(400).json({
        error: "没有收到文件（multipart 表单字段名应为 files）",
        saved: [],
        failures: [],
      });
      return;
    }

    const outcome = await saveUploads(files);
    const rebuilding = outcome.saved.length > 0;
    if (rebuilding) {
      console.log(`[upload] ${outcome.saved.length} file(s) added, rebuilding index…`);
      // Deliberately not awaited: the response returns immediately and the
      // frontend follows progress through /api/status.
      void ensureIndex().catch(() => {
        /* failure is already recorded in buildError */
      });
    }

    res.json({ ...outcome, rebuilding });
  }),
);

app.post(
  "/api/chat",
  wrap(async (req, res) => {
    const parsed = readChatRequest(req.body);
    if (!parsed) {
      res.status(400).json({ error: "question 不能为空" });
      return;
    }

    const blocked = await readyForChat();
    if (blocked) {
      res.status(503).json({ error: blocked });
      return;
    }

    res.json(await answer(parsed.question, parsed.history));
  }),
);

/**
 * Streaming chat: same request body as /api/chat, but the answer arrives as
 * Server-Sent Events instead of one JSON payload.
 *
 * Frame order: `sources` (retrieved cards, sent before any text) → `delta`* →
 * `done` | `error`. Each frame's payload is the full StreamEvent, so the client
 * can switch on `type` and ignore the SSE `event:` field entirely.
 */
app.post(
  "/api/chat/stream",
  wrap(async (req, res) => {
    const parsed = readChatRequest(req.body);
    if (!parsed) {
      res.status(400).json({ error: "question 不能为空" });
      return;
    }

    const blocked = await readyForChat();
    if (blocked) {
      res.status(503).json({ error: blocked });
      return;
    }

    res.writeHead(200, {
      "Content-Type": "text/event-stream; charset=utf-8",
      // `no-transform` keeps any proxy from buffering or gzipping the stream,
      // either of which would turn live output into one big delayed blob.
      "Cache-Control": "no-cache, no-transform",
      Connection: "keep-alive",
      "X-Accel-Buffering": "no",
    });

    const abort: AbortFlag = { aborted: false };
    res.on("close", () => {
      abort.aborted = true;
    });

    const send = (event: string, data: unknown): void => {
      if (res.writableEnded) return;
      res.write(`event: ${event}\ndata: ${JSON.stringify(data)}\n\n`);
    };

    // Errors are handled here rather than bubbling to the generic handler:
    // by this point headers are already sent, so a res.json() from the error
    // middleware would crash against an open stream.
    try {
      for await (const event of answerStream(parsed.question, parsed.history, abort)) {
        send(event.type, event);
      }
    } catch (error) {
      send("error", {
        type: "error",
        error: error instanceof Error ? error.message : String(error),
      });
    }
    res.end();
  }),
);

app.use("/api", (_req, res) => {
  res.status(404).json({ error: "接口不存在" });
});

app.use((error: unknown, _req: Request, res: Response, _next: NextFunction) => {
  // multer rejects before the route handler runs, so its errors land here.
  if (error instanceof multer.MulterError) {
    const status = error.code === "LIMIT_FILE_SIZE" ? 413 : 400;
    const hint = error.code === "LIMIT_FILE_SIZE" ? "（单个文件上限 15MB）" : "";
    res.status(status).json({ error: `上传失败：${error.message}${hint}` });
    return;
  }

  const message = error instanceof Error ? error.message : String(error);
  console.error("[http]", message);
  res.status(500).json({ error: message });
});

async function main(): Promise<void> {
  // Settings must land in the environment before the first request, or a key
  // saved from the panel wouldn't be live until the process restarted.
  await loadSettings();

  app.listen(RAG.port, RAG.host, () => {
    console.log(`\n  RAG API listening on http://${RAG.host}:${RAG.port}`);
    console.log(`  knowledge dir : ${PATHS.knowledgeDir}`);
    console.log(`  vector store  : ${PATHS.vectorDir}`);
    console.log(`  embeddings    : ${RAG.embeddingModel} (local)`);
    console.log(
      `  generation    : ${
        hasLlmCredentials()
          ? `${RAG.llmProvider} / ${RAG.llmModel}`
          : `extractive fallback (${RAG.llmProvider} has no credentials)`
      }\n`,
    );
    void ensureIndex();
  });
}

void main().catch((error: unknown) => {
  console.error("Failed to start:", error instanceof Error ? error.message : error);
  process.exit(1);
});
