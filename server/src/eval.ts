import { search } from "./store.js";
import { readIndexMeta } from "./ingest.js";
import { RAG } from "./config.js";

/**
 * Retrieval-only smoke test:
 *
 *   npm run eval --prefix server
 *
 * For each question, check whether the file that actually contains the answer
 * shows up, and at what rank. This measures the vector store alone — no LLM is
 * called, so it's free to run as often as you like.
 *
 * Tune CHUNK_SIZE / TOP_K / EMBEDDING_MODEL against these numbers rather than
 * judging from a handful of chat answers.
 */
const CASES: Array<[query: string, expectedFile: string]> = [
  ["问题单的处理流程是什么", "问题处理流程规范.docx"],
  ["Bug 严重程度如何分级", "问题处理流程规范.docx"],
  ["问题单在什么情况下可以关闭", "问题处理流程规范.docx"],
  ["git 分支命名有什么规范", "配置管理规范.docx"],
  ["提交代码前要做哪些配置管理动作", "配置管理规范.docx"],
  ["版本发布需要走什么流程", "配置管理规范.docx"],
  ["Java 中异常处理有什么强制要求", "Java开发规范.docx"],
  ["日志打印有什么规范", "Java开发规范.docx"],
  ["集合和数组的使用有什么要求", "Java开发规范.docx"],
];

const K = RAG.topK;

/**
 * Chunk-level checks.
 *
 * The file-level cases above cannot catch a chunk-ranking failure: for the
 * 异常处理 query, 4 of 5 results came from the correct document, so file recall
 * reported 100% while the actual top hit was that document's *preface* and the
 * section that answered the question ranked #6 — outside K entirely.
 *
 * `mustMatch` is checked against the text of the #1 chunk only. These are
 * hand-picked to admit the on-topic section and reject the passages that used
 * to win (a preface, a code-cleanup note, a different document).
 */
const ON_TOPIC: Array<{ query: string; mustMatch: RegExp }> = [
  {
    query: "Java 开发规范对异常处理有什么要求？",
    mustMatch: /捕获异常|Throwable|预检查|RuntimeException|catch 时/,
  },
  { query: "问题单在什么情况下可以关闭？", mustMatch: /关闭/ },
  { query: "git 分支命名有什么规范？", mustMatch: /分支/ },
  { query: "日志打印有什么规范？", mustMatch: /日志|logger/ },
  { query: "Bug 严重程度如何分级？", mustMatch: /严重程度/ },
];

async function scoreOnTopic(): Promise<number> {
  let pass = 0;
  const rows: string[] = [];

  for (const c of ON_TOPIC) {
    const hits = await search(c.query, 1);
    const top = hits[0]?.doc.pageContent ?? "";
    const ok = c.mustMatch.test(top);
    if (ok) pass += 1;
    rows.push(
      `  ${ok ? "PASS" : "FAIL"}  ${c.query}  →  ${top.replace(/\s+/g, " ").slice(0, 62)}…`,
    );
  }

  console.log(rows.join("\n"));
  console.log(`on-topic@1 : ${pass}/${ON_TOPIC.length} (${((pass / ON_TOPIC.length) * 100).toFixed(0)}%)`);
  return pass / ON_TOPIC.length;
}

async function main(): Promise<void> {
  const meta = await readIndexMeta();
  const rows: string[] = [];
  let reciprocalRankSum = 0;
  let top1 = 0;
  let inTopK = 0;

  for (const [query, expected] of CASES) {
    const hits = await search(query, K);
    const rank = hits.findIndex((hit) => String(hit.doc.metadata.source) === expected);
    const reciprocalRank = rank === -1 ? 0 : 1 / (rank + 1);

    reciprocalRankSum += reciprocalRank;
    if (rank === 0) top1 += 1;
    if (rank !== -1) inTopK += 1;

    rows.push(
      `${rank === -1 ? "MISS" : `#${rank + 1}`}  RR=${reciprocalRank.toFixed(2)}  ${query}  →  ${expected.replace(/\.docx$/, "")}`,
    );
  }

  console.log(rows.join("\n"));
  console.log(
    `\nchunks=${meta?.chunkCount ?? 0}  model=${meta?.embeddingModel ?? "?"}  prefixes=${JSON.stringify(
      [RAG.embeddingQueryPrefix, RAG.embeddingPassagePrefix],
    )}`,
  );
  console.log(`top-1 accuracy : ${top1}/${CASES.length} (${((top1 / CASES.length) * 100).toFixed(0)}%)`);
  console.log(`top-${K} recall  : ${inTopK}/${CASES.length} (${((inTopK / CASES.length) * 100).toFixed(0)}%)`);
  console.log(`MRR@${K}        : ${(reciprocalRankSum / CASES.length).toFixed(3)}`);

  console.log("\nchunk-level (hand-checked; file-level cannot see this failure)");
  const onTopic = await scoreOnTopic();

  // Non-zero exit on a regression so this can gate a build.
  if (onTopic < 1) {
    process.exitCode = 1;
  }
}

main().catch((error: unknown) => {
  console.error("Eval failed:", error instanceof Error ? error.message : error);
  process.exit(1);
});
