export interface StaleInfo {
  stale: boolean;
  reason: string;
}

export interface IndexStatus {
  ready: boolean;
  indexing: boolean;
  buildError: string | null;
  stale: StaleInfo;
  files: string[];
  knowledgeDir: string;
  chunkCount: number;
  builtAt: string | null;
  embeddingModel: string;
  llmModel: string;
  llmProvider: string;
  llmMode: "llm" | "extractive";
  topK: number;
  chunkSize: number;
  chunkOverlap: number;
}

export interface IngestResult {
  files: string[];
  chunkCount: number;
  embeddingModel: string;
  durationMs: number;
  builtAt: string;
}

export interface Source {
  index: number;
  file: string;
  similarity: number;
  snippet: string;
}

export interface Turn {
  question: string;
  answer: string;
}

export interface ChatResult {
  answer: string;
  sources: Source[];
  mode: "llm" | "extractive";
  model: string | null;
  latencyMs: number;
}

/** Per-provider settings as the server reports them — never a raw key. */
export interface PublicProvider {
  model: string;
  baseUrl: string;
  apiKeySet: boolean;
  apiKeyMasked: string | null;
}

export interface Settings {
  llmProvider: string;
  providers: Record<string, PublicProvider>;
  embeddingModel: string;
}

export interface SavedSettings extends Settings {
  applied: { llmProvider: string; credentials: boolean };
}

export interface SettingsUpdate {
  llmProvider?: string;
  providers?: Record<string, { model?: string; baseUrl?: string; apiKey?: string }>;
}

export interface UploadResult {
  saved: string[];
  failures: Array<{ name: string; reason: string }>;
  rebuilding: boolean;
}

async function request<T>(path: string, init: RequestInit = {}): Promise<T> {
  const response = await fetch(path, {
    ...init,
    headers:
      init.body instanceof FormData
        ? init.headers // browser must set the multipart boundary itself
        : { "Content-Type": "application/json", ...init.headers },
  });

  const body = (await response.json().catch(() => ({}))) as { error?: string };
  if (!response.ok) {
    throw new Error(body.error ?? `请求失败（HTTP ${response.status}）`);
  }
  return body as T;
}

export function getStatus(): Promise<IndexStatus> {
  return request<IndexStatus>("/api/status");
}

export function rebuildIndex(): Promise<IngestResult> {
  return request<IngestResult>("/api/ingest", { method: "POST" });
}

export function ask(question: string, history: Turn[]): Promise<ChatResult> {
  return request<ChatResult>("/api/chat", {
    method: "POST",
    body: JSON.stringify({ question, history }),
  });
}

/** Mirrors the server's StreamEvent — see server/src/answer.ts. */
export type StreamEvent =
  | { type: "sources"; sources: Source[]; mode: "llm" | "extractive"; model: string | null }
  | { type: "delta"; text: string }
  | { type: "done"; latencyMs: number }
  | { type: "error"; error: string };

export interface StreamHandlers {
  onSources: (info: {
    sources: Source[];
    mode: "llm" | "extractive";
    model: string | null;
  }) => void;
  onDelta: (text: string) => void;
  onDone: (latencyMs: number) => void;
}

/** One SSE frame: `event: <name>` / `data: <json>` separated by a blank line. */
function parseFrame(frame: string): StreamEvent | null {
  let data = "";
  for (const line of frame.split("\n")) {
    if (line.startsWith("data:")) data += line.slice(5).replace(/^ /, "");
  }
  if (!data) return null;
  try {
    return JSON.parse(data) as StreamEvent;
  } catch {
    return null;
  }
}

/**
 * Ask a question and receive the answer as it is generated.
 *
 * Transport is Server-Sent Events over fetch rather than EventSource because we
 * need a POST body. `TextDecoder` is used with `{stream: true}` — Chinese text
 * is multi-byte UTF-8, and a chunk boundary can split a character in half,
 * which would otherwise surface as a replacement character in the output.
 */
export async function askStream(
  question: string,
  history: Turn[],
  handlers: StreamHandlers,
  signal?: AbortSignal,
): Promise<void> {
  const response = await fetch("/api/chat/stream", {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ question, history }),
    signal,
  });

  if (!response.ok) {
    const body = (await response.json().catch(() => ({}))) as { error?: string };
    throw new Error(body.error ?? `请求失败（HTTP ${response.status}）`);
  }
  if (!response.body) throw new Error("当前浏览器不支持流式响应");

  const reader = response.body.getReader();
  const decoder = new TextDecoder();
  let buffer = "";

  for (;;) {
    const { done, value } = await reader.read();
    if (done) break;
    buffer += decoder.decode(value, { stream: true });

    for (;;) {
      const separator = buffer.indexOf("\n\n");
      if (separator === -1) break;
      const frame = buffer.slice(0, separator);
      buffer = buffer.slice(separator + 2);

      const event = parseFrame(frame);
      if (!event) continue;

      switch (event.type) {
        case "sources":
          handlers.onSources(event);
          break;
        case "delta":
          if (event.text) handlers.onDelta(event.text);
          break;
        case "done":
          handlers.onDone(event.latencyMs);
          break;
        case "error":
          throw new Error(event.error);
      }
    }
  }
}

export function getSettings(): Promise<Settings> {
  return request<Settings>("/api/settings");
}

/** Only the fields present in `update` change — omitted ones stay as stored. */
export function saveSettings(update: SettingsUpdate): Promise<SavedSettings> {
  return request<SavedSettings>("/api/settings", {
    method: "POST",
    body: JSON.stringify(update),
  });
}

export function uploadFiles(files: File[]): Promise<UploadResult> {
  const form = new FormData();
  for (const file of files) form.append("files", file, file.name);
  return request<UploadResult>("/api/upload", { method: "POST", body: form });
}
