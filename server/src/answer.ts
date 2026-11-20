import { AIMessage, HumanMessage, SystemMessage, type BaseMessage } from "@langchain/core/messages";
import { RAG } from "./config.js";
import { createChatModel, hasLlmCredentials } from "./llm.js";
import { search, type Hit } from "./store.js";

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
  /** "llm" when Claude wrote the answer, "extractive" when we only retrieved. */
  mode: "llm" | "extractive";
  model: string | null;
  latencyMs: number;
}

/**
 * Events emitted by `answerStream`, in the order they are produced.
 *
 * `sources` always comes first so the UI can show the retrieved cards before a
 * single word of the answer exists — the retrieval is local and finishes in
 * milliseconds, while the model takes seconds.
 */
export type StreamEvent =
  | { type: "sources"; sources: Source[]; mode: "llm" | "extractive"; model: string | null }
  | { type: "delta"; text: string }
  | { type: "done"; latencyMs: number }
  | { type: "error"; error: string };

/** Set by the HTTP layer when the client goes away, so we stop generating. */
export interface AbortFlag {
  aborted: boolean;
}

const SYSTEM_PROMPT = `你是知识库问答助手，负责回答关于所提供文档的问题。

严格遵守以下规则：
1. 只依据下面提供的「参考资料」作答，不要编造参考资料中不存在的内容。
2. 引用资料时标注来源编号，例如 [1]、[2]。
3. 如果参考资料不足以回答问题，直接说明「参考资料中没有找到相关信息」，不要猜测或用常识补充。
4. 用与提问相同的语言回答，先给结论，再给依据，保持简洁。
5. 如果多份资料之间存在冲突，指明冲突并分别说明各自的出处。`;

/** How much of the conversation to replay. Older turns are dropped. */
const HISTORY_TURNS = 6;

function sourceLabel(i: number, hit: Hit): string {
  const file = String(hit.doc.metadata.source ?? "未知文件");
  return `[${i + 1}] 来源：${file}`;
}

function buildContext(hits: Hit[]): string {
  if (hits.length === 0) return "（没有检索到任何参考资料）";
  return hits.map((hit, i) => `${sourceLabel(i, hit)}\n${hit.doc.pageContent}`).join("\n\n---\n\n");
}

function buildSources(hits: Hit[]): Source[] {
  return hits.map((hit, i) => ({
    index: i + 1,
    file: String(hit.doc.metadata.source ?? "未知文件"),
    similarity: Number(hit.similarity.toFixed(4)),
    snippet: hit.doc.pageContent.slice(0, 160),
  }));
}

/** Concatenated visible text of an AI message, ignoring thinking/tool blocks. */
function textOf(content: unknown): string {
  if (typeof content === "string") return content;
  if (!Array.isArray(content)) return String(content ?? "");
  return content
    .filter((block): block is { type: "text"; text: string } =>
      Boolean(block) && typeof block === "object" && (block as { type?: string }).type === "text",
    )
    .map((block) => block.text)
    .join("");
}

/**
 * Text of a *streaming* chunk. Deliberately looser than `textOf`: providers
 * disagree on block naming mid-stream (OpenAI sends `{type:"text",text}`,
 * Anthropic sends `{type:"text_delta",text}`), and thinking blocks carry their
 * payload under `thinking` so they still fall through.
 */
function chunkText(content: unknown): string {
  if (typeof content === "string") return content;
  if (!Array.isArray(content)) return content == null ? "" : String(content);

  let out = "";
  for (const block of content) {
    if (typeof block === "string") {
      out += block;
      continue;
    }
    if (block && typeof block === "object") {
      const text = (block as { text?: unknown }).text;
      if (typeof text === "string") out += text;
    }
  }
  return out;
}

/**
 * Build the message array sent to the model.
 *
 * History is passed as real messages rather than template variables: generated
 * answers routinely contain braces, which ChatPromptTemplate would try to
 * interpolate and reject. A system-only template would also leave the request
 * with an empty `messages` array, which the API rejects outright.
 */
function buildMessages(question: string, hits: Hit[], history: Turn[]): BaseMessage[] {
  const recent = history.slice(-HISTORY_TURNS);
  return [
    new SystemMessage(SYSTEM_PROMPT),
    ...recent.flatMap((turn) => [new HumanMessage(turn.question), new AIMessage(turn.answer)]),
    new HumanMessage(`参考资料：\n\n${buildContext(hits)}\n\n用户问题：${question}`),
  ];
}

/**
 * Without credentials we can still retrieve — the whole ingest + vector-search
 * half of the pipeline runs offline. We return the matched chunks verbatim and
 * label the response `extractive` so the UI never passes it off as an answer.
 */
function extractiveAnswer(question: string, hits: Hit[]): string {
  if (hits.length === 0) {
    return `「${question}」没有在知识库中检索到相关内容。`;
  }
  const body = hits
    .map((hit, i) => `${sourceLabel(i, hit)}\n${hit.doc.pageContent}`)
    .join("\n\n");
  return [
    "当前未配置 Claude 凭据，以下是知识库中检索到的原始片段（未经模型归纳）：",
    "",
    body,
  ].join("\n");
}

/**
 * Answer one question: retrieve the top-k chunks, then ask Claude to answer
 * from those chunks only.
 *
 * No temperature/top_k/top_p is set on purpose — `claude-opus-5-5` rejects
 * sampling params (it runs adaptive thinking), and LangChain already strips
 * them for this model family.
 */
export async function answer(
  question: string,
  history: Turn[] = [],
): Promise<ChatResult> {
  const started = Date.now();
  const hits = await search(question, RAG.topK);
  const sources = buildSources(hits);

  if (!hasLlmCredentials()) {
    return {
      answer: extractiveAnswer(question, hits),
      sources,
      mode: "extractive",
      model: null,
      latencyMs: Date.now() - started,
    };
  }

  // Backend-specific construction lives in llm.ts — this function only cares
  // that it gets something it can .invoke().
  const model = createChatModel();
  const response = await model.invoke(buildMessages(question, hits, history));

  return {
    answer: textOf(response.content),
    sources,
    mode: "llm",
    model: RAG.llmModel,
    latencyMs: Date.now() - started,
  };
}

/**
 * Streaming variant of `answer()`, yielding the response as it is generated.
 *
 * Emitting `sources` before any text is the whole point: retrieval is local and
 * takes milliseconds, generation takes seconds. Without it the user stares at a
 * blank bubble for the entire latency, which is the failure mode this replaces.
 *
 * Errors are yielded as an `error` event rather than thrown, so the transport
 * can report them inside the stream instead of tearing down a half-written
 * answer the user has already been reading.
 */
export async function* answerStream(
  question: string,
  history: Turn[] = [],
  abort?: AbortFlag,
): AsyncGenerator<StreamEvent> {
  const started = Date.now();
  const hits = await search(question, RAG.topK);
  const sources = buildSources(hits);
  const stopped = (): boolean => abort?.aborted === true;

  // No credentials: still stream, so the fallback gets the same typewriter
  // treatment instead of popping in as one block.
  if (!hasLlmCredentials()) {
    yield { type: "sources", sources, mode: "extractive", model: null };
    if (!stopped()) {
      yield { type: "delta", text: extractiveAnswer(question, hits) };
    }
    if (!stopped()) yield { type: "done", latencyMs: Date.now() - started };
    return;
  }

  const model = createChatModel();
  yield { type: "sources", sources, mode: "llm", model: RAG.llmModel };

  try {
    const stream = await model.stream(buildMessages(question, hits, history));
    for await (const chunk of stream) {
      if (stopped()) return;
      const text = chunkText(chunk.content);
      if (text) yield { type: "delta", text };
    }
    if (!stopped()) yield { type: "done", latencyMs: Date.now() - started };
  } catch (error) {
    if (stopped()) return;
    yield { type: "error", error: error instanceof Error ? error.message : String(error) };
  }
}
