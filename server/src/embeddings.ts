import { env } from "@huggingface/transformers";
import type { PretrainedModelOptions } from "@huggingface/transformers";
import { HuggingFaceTransformersEmbeddings } from "@langchain/community/embeddings/huggingface_transformers";
import { RAG, PATHS } from "./config.js";

// transformers.js resolves these lazily, at download time, so assigning them
// here is enough — but it must happen before the first embed() call.
env.remoteHost = RAG.hfEndpoint;
env.cacheDir = PATHS.hfCache;

/**
 * q8 is ~1/4 the size of fp32 and fast enough on CPU. It happens to also be
 * transformers.js's default for the wasm device — setting it explicitly just
 * makes the choice visible and silences the "dtype not specified" warning.
 *
 * Typed as the wider `PretrainedModelOptions` rather than the `PretrainedOptions`
 * LangChain declares, because `dtype` lives on the former.
 */
const pretrainedOptions: PretrainedModelOptions = { dtype: "q8" };

/**
 * Adds the role prefixes some retrieval models require (E5, BGE, GTE).
 *
 * Both sides of the comparison must be prefixed the same way or every score is
 * meaningless — questions are `query: `-prefixed, indexed chunks `passage: `-prefixed.
 * Since ingest and query share this one instance, the pairing can't drift.
 */
class PrefixedEmbeddings extends HuggingFaceTransformersEmbeddings {
  private readonly queryPrefix: string;
  private readonly passagePrefix: string;

  constructor(fields: ConstructorParameters<typeof HuggingFaceTransformersEmbeddings>[0]) {
    super(fields);
    this.queryPrefix = RAG.embeddingQueryPrefix;
    this.passagePrefix = RAG.embeddingPassagePrefix;
  }

  override embedDocuments(texts: string[]): Promise<number[][]> {
    return super.embedDocuments(texts.map((text) => `${this.passagePrefix}${text}`));
  }

  override embedQuery(text: string): Promise<number[]> {
    return super.embedQuery(`${this.queryPrefix}${text}`);
  }
}

let instance: PrefixedEmbeddings | null = null;

/**
 * Local ONNX embeddings, shared across ingest and query time.
 * Indexing and retrieval must use the same model or the vectors won't align,
 * so we hold a single instance for the process lifetime.
 */
export function getEmbeddings(): PrefixedEmbeddings {
  instance ??= new PrefixedEmbeddings({
    model: RAG.embeddingModel,
    pretrainedOptions,
  });
  return instance;
}
