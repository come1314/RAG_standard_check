<script setup lang="ts">
import { computed, nextTick, ref, watch } from "vue";
import { askStream, type Source, type Turn } from "../api";
import { renderMarkdown } from "../markdown";

interface Message {
  id: number;
  role: "user" | "assistant";
  /** Full text as received — also what gets replayed as conversation history. */
  content: string;
  /** Assistant answers only: sanitized Markdown HTML of the *revealed* text. */
  html: string;
  sources: Source[];
  /** Whether this answer's source cards are showing; collapsed by default. */
  sourcesOpen: boolean;
  mode: "llm" | "extractive";
  model: string | null;
  latencyMs: number;
  /** True while the typewriter is still draining into this message. */
  streaming: boolean;
}

const props = defineProps<{
  blocked: boolean;
  blockedReason: string | null;
}>();

const messages = ref<Message[]>([]);
const input = ref("");
const busy = ref(false);
const error = ref<string | null>(null);
const scroller = ref<HTMLElement | null>(null);

let nextId = 1;

/**
 * Typewriter buffer.
 *
 * The network delivers text in token-sized bursts, which reads as choppy — and
 * on a slow endpoint a whole paragraph can land at once. So received text goes
 * into `pending`, and a requestAnimationFrame loop drains it into `revealed` at
 * a steady rate; `revealed` is the only thing rendered. Motion stays smooth no
 * matter how the deltas arrive.
 *
 * These are deliberately plain variables rather than refs: only this loop reads
 * them. What the template needs (message.html / message.streaming) is written
 * through `messages.value.find(...)`, which is the reactive proxy — mutating a
 * raw object held on the side would never trigger a re-render in Vue 3.
 */
let liveMessageId = 0;
let pending = "";
let revealed = "";
let finished = true;
let typingRaf = 0;
let lastRender = 0;
/**
 * Retrieved chunks, held back until the typewriter settles.
 *
 * The server sends `sources` first — retrieval is milliseconds, generation is
 * seconds — but that is a transport detail, not a viewing order. Cards appearing
 * above an empty bubble pull the eye away from the text as it types, so they are
 * parked here and attached once every character is on screen.
 */
let stashedSources: Source[] = [];

/**
 * Reveal rate, in characters per frame (requestAnimationFrame, ~60fps).
 *
 * `CHARS_PER_FRAME` is the floor so a slow stream still visibly types;
 * `MAX_CHARS_PER_FRAME` is the ceiling so one big payload (the extractive
 * fallback hands over the whole answer at once) reads as typing rather than
 * popping in; `CATCH_UP_DIVISOR` in between lets a backlog close quickly
 * instead of lagging behind the network.
 */
const CHARS_PER_FRAME = 3;
const MAX_CHARS_PER_FRAME = 10;
const CATCH_UP_DIVISOR = 18;

function messageById(id: number): Message | undefined {
  return messages.value.find((m) => m.id === id);
}

/** Distinct source file names, shown on the collapsed summary bar. */
function sourceFiles(sources: Source[]): string {
  return [...new Set(sources.map((s) => s.file.replace(/\.docx$/, "")))].join("、");
}

function scheduleTyping(): void {
  if (!typingRaf) typingRaf = requestAnimationFrame(tickTypewriter);
}

function tickTypewriter(): void {
  typingRaf = 0;

  if (pending.length > 0) {
    const step = Math.max(
      CHARS_PER_FRAME,
      Math.min(MAX_CHARS_PER_FRAME, Math.ceil(pending.length / CATCH_UP_DIVISOR)),
    );
    revealed += pending.slice(0, step);
    pending = pending.slice(step);
  }

  const settled = finished && pending.length === 0;
  const now = performance.now();

  // Re-rendering Markdown every frame is wasted work while the drain is steady;
  // ~30fps of paint is plenty for text. Always paint on the final frame.
  if (settled || now - lastRender > 32) {
    lastRender = now;
    const target = messageById(liveMessageId);
    if (target) {
      target.html = renderMarkdown(revealed);
      if (settled) {
        target.streaming = false;
        // Last character is on screen — now the source cards are allowed in.
        if (stashedSources.length > 0) {
          target.sources = stashedSources;
          stashedSources = [];
        }
      }
    }
  }

  if (settled) return;
  scheduleTyping();
}

const suggestions = [
  "问题处理流程是什么？",
  "配置管理有哪些要求？",
  "Java 异常处理有什么要求？",
];

/** Completed question/answer pairs, replayed as context on each turn. */
const history = computed<Turn[]>(() => {
  const turns: Turn[] = [];
  for (let i = 0; i < messages.value.length - 1; i += 1) {
    const current = messages.value[i];
    const next = messages.value[i + 1];
    if (current.role === "user" && next.role === "assistant") {
      turns.push({ question: current.content, answer: next.content });
    }
  }
  return turns;
});

/**
 * True from send until the first character is revealed — the retrieval window,
 * the only stretch where a progress indicator tells the user anything.
 */
const waitingForFirstToken = computed(() => {
  const last = messages.value[messages.value.length - 1];
  return busy.value && last?.role === "assistant" && last.content === "";
});

async function send(preset?: string): Promise<void> {
  const question = (preset ?? input.value).trim();
  if (!question || busy.value || props.blocked) return;

  // Snapshot history before appending this turn.
  const prior = history.value;

  error.value = null;
  input.value = "";
  messages.value.push({
    id: nextId++,
    role: "user",
    content: question,
    html: "",
    sources: [],
    sourcesOpen: false,
    mode: "llm",
    model: null,
    latencyMs: 0,
    streaming: false,
  });
  busy.value = true;

  // The answer bubble exists from the start so there is somewhere for text to
  // appear; it stays empty until the first delta, which is the retrieval window.
  const replyId = nextId++;
  messages.value.push({
    id: replyId,
    role: "assistant",
    content: "",
    html: "",
    sources: [],
    sourcesOpen: false,
    mode: "llm",
    model: null,
    latencyMs: 0,
    streaming: true,
  });

  liveMessageId = replyId;
  pending = "";
  revealed = "";
  finished = false;
  lastRender = 0;
  stashedSources = [];
  scheduleTyping();

  try {
    await askStream(question, prior, {
      onSources: (info) => {
        const target = messageById(replyId);
        if (!target) return;
        // Status is useful immediately — in particular `extractive` is a warning
        // that no model was called, and hiding it behind the animation would
        // mislead. Only the cards themselves wait.
        target.mode = info.mode;
        target.model = info.model;
        stashedSources = info.sources;
      },
      onDelta: (text) => {
        // `content` accumulates the full text for history replay; `pending`
        // feeds the typewriter. The two converge once the drain finishes.
        const target = messageById(replyId);
        if (target) target.content += text;
        pending += text;
        scheduleTyping();
      },
      onDone: (latencyMs) => {
        const target = messageById(replyId);
        if (target) target.latencyMs = latencyMs;
        finished = true;
        scheduleTyping();
      },
    });
  } catch (cause) {
    // Keep whatever already made it to screen — a half answer plus a clear error
    // beats discarding thirty seconds of generation. Only roll back when there
    // is nothing at all to show, and then put the question back in the composer.
    const target = messageById(replyId);
    if (!target?.content) {
      messages.value = messages.value.filter((m) => m.id !== replyId);
      messages.value.pop();
      input.value = question;
      liveMessageId = 0;
      stashedSources = [];
    } else if (target) {
      // Partial answer: flush what we have, cards included.
      if (stashedSources.length > 0) {
        target.sources = stashedSources;
        stashedSources = [];
      }
      target.streaming = false;
    }
    error.value = cause instanceof Error ? cause.message : String(cause);
  } finally {
    // Covers a stream that ends without a `done` frame (proxy cut off, etc.).
    finished = true;
    scheduleTyping();
    busy.value = false;
  }
}

function onKeydown(event: KeyboardEvent): void {
  if (event.key === "Enter" && !event.shiftKey) {
    event.preventDefault();
    void send();
  }
}

/**
 * Clicking a `[1]` badge scrolls to that source card and flashes it, so the
 * numbers in the answer actually lead somewhere instead of being decoration.
 * The card is expanded first — landing on a collapsed header would show why it
 * was cited only after another click, which defeats the point.
 */
function onCitationClick(event: MouseEvent): void {
  const badge = (event.target as HTMLElement | null)?.closest<HTMLElement>("[data-cite]");
  if (!badge?.dataset.cite) return;

  const wrap = event.currentTarget as HTMLElement | null;
  const card = wrap?.querySelector<HTMLElement>(`.source[data-index="${badge.dataset.cite}"]`);
  if (!wrap || !card) return;

  const message = messageById(Number(wrap.dataset.messageId));
  if (message) message.sourcesOpen = true;

  // Wait for the expand to paint before scrolling, or the card moves under the
  // scroll target and the highlight lands in the wrong place.
  void nextTick(() => {
    card.scrollIntoView({ behavior: "smooth", block: "center" });
    card.classList.remove("flash");
    // Restart the CSS animation even if the same citation is clicked twice.
    void card.offsetWidth;
    card.classList.add("flash");
    window.setTimeout(() => card.classList.remove("flash"), 1600);
  });
}

watch(
  [messages, busy],
  () => {
    void nextTick(() => {
      // Smooth scrolling at typewriter rate stacks animations and lags behind;
      // jump instead while text is streaming, smooth otherwise.
      scroller.value?.scrollTo({
        top: scroller.value.scrollHeight,
        behavior: !finished ? "auto" : "smooth",
      });
    });
  },
  { deep: true },
);
</script>

<template>
  <main class="chat">
    <div ref="scroller" class="transcript">
      <div v-if="messages.length === 0" class="empty">
        <h3>基于 knowledgeFile/ 的文档问答</h3>
        <p>
          每个问题会先在本地向量库中检索最相关的文档片段，再交由模型仅依据这些片段作答。
        </p>
        <ul>
          <li v-for="s in suggestions" :key="s">
            <a href="#" @click.prevent="send(s)">{{ s }}</a>
          </li>
        </ul>
      </div>

      <template v-for="message in messages" :key="message.id">
        <div
          :class="['msg', message.role]"
          :data-message-id="message.id"
          @click="message.role === 'assistant' && onCitationClick($event)"
        >
          <div v-if="message.role === 'user'" class="bubble user-text">{{ message.content }}</div>
          <!-- eslint-disable-next-line vue/no-v-html -- sanitized in renderMarkdown() -->
          <div
            v-else
            :class="['bubble', 'answer', { typing: message.streaming }]"
            v-html="message.html"
          ></div>

          <div v-if="message.role === 'assistant'" class="meta">
            <span v-if="message.streaming" class="spinner"></span>
            <span v-if="message.model" class="meta-pill">{{ message.model }}</span>
            <span v-if="!message.streaming" class="meta-pill">{{ (message.latencyMs / 1000).toFixed(1) }}s</span>
            <span v-if="!message.streaming" class="meta-pill">召回 {{ message.sources.length }} 条</span>
            <span
              v-if="message.mode === 'extractive'"
              class="badge warn"
              title="未配置模型凭证，返回的是检索到的原始片段"
            >
              未调用模型
            </span>
          </div>

          <template v-if="message.sources.length">
            <div
              class="sources-toggle"
              role="button"
              tabindex="0"
              :aria-expanded="!!message.sourcesOpen"
              @click="message.sourcesOpen = !message.sourcesOpen"
              @keydown.enter.prevent="message.sourcesOpen = !message.sourcesOpen"
              @keydown.space.prevent="message.sourcesOpen = !message.sourcesOpen"
            >
              <span class="sources-count">来源 {{ message.sources.length }} 条</span>
              <span class="sources-files">{{ sourceFiles(message.sources) }}</span>
              <span class="chev" aria-hidden="true"></span>
            </div>

            <div class="sources-wrap" :class="{ open: message.sourcesOpen }">
              <div class="sources">
                <div
                  v-for="source in message.sources"
                  :key="source.index"
                  class="source"
                  :data-index="source.index"
                  :style="{ animationDelay: (source.index - 1) * 70 + 'ms' }"
                >
                  <div class="source-head">
                    <span class="source-num">{{ source.index }}</span>
                    <span class="source-file">{{ source.file.replace(/\.docx$/, "") }}</span>
                    <span class="score" title="余弦相似度，越接近 1 越相关">
                      {{ source.similarity.toFixed(3) }}
                      <span class="score-bar" aria-hidden="true">
                        <span
                          class="score-fill"
                          :style="{ width: Math.round(source.similarity * 100) + '%' }"
                        ></span>
                      </span>
                    </span>
                  </div>
                  <div class="snippet">{{ source.snippet }}</div>
                </div>
              </div>
            </div>
          </template>
        </div>
      </template>

      <div v-if="waitingForFirstToken" class="msg assistant">
        <div class="bubble">
          <span class="spinner"></span>
          正在检索知识库…
        </div>
      </div>
    </div>

    <div class="composer">
      <div v-if="blocked || error" :class="error ? 'error-banner' : 'hint'" style="margin-bottom: 10px">
        {{ error ?? blockedReason }}
      </div>

      <div class="composer-inner">
        <textarea
          v-model="input"
          :disabled="busy || blocked"
          rows="1"
          placeholder="输入你的问题，Enter 发送，Shift+Enter 换行"
          @keydown="onKeydown"
        ></textarea>
        <button
          class="send"
          :disabled="busy || blocked || !input.trim()"
          @click="send()"
        >
          {{ busy ? "…" : "发送" }}
        </button>
      </div>
      <p class="footer-note">
        回答仅依据检索到的文档片段生成，可能遗漏未收录的内容；请以原始规范文件为准。
      </p>
    </div>
  </main>
</template>

<style scoped>
.empty a {
  color: var(--accent);
  text-decoration: none;
}

.empty a:hover {
  text-decoration: underline;
}

/* Typewriter caret — sits after the revealed text, before the next character
   shows up. Removed by dropping the .typing class when the drain finishes. */
.answer.typing::after {
  content: "";
  display: inline-block;
  width: 2px;
  height: 1.05em;
  margin-left: 3px;
  vertical-align: -0.15em;
  background: var(--accent);
  animation: caret-blink 1s steps(2, start) infinite;
}

@keyframes caret-blink {
  to {
    visibility: hidden;
  }
}

/*
 * Source cards enter only after the last character is on screen, one after
 * another. `both` fill holds each card at its `from` state through its
 * animation-delay, which is what produces the stagger.
 */
@keyframes source-in {
  from {
    opacity: 0;
    transform: translateY(10px);
  }
  to {
    opacity: 1;
    transform: none;
  }
}

.sources .source {
  animation: source-in 0.4s ease both;
}
</style>
