import { ChatAnthropic } from "@langchain/anthropic";
import { ChatOpenAI } from "@langchain/openai";
import Anthropic from "@anthropic-ai/sdk";
import type { BaseChatModel } from "@langchain/core/language_models/chat_models";
import { RAG } from "./config.js";

/**
 * Whether the configured backend actually has credentials.
 *
 * Both backends degrade the same way when they don't: retrieval keeps working
 * and answers fall back to raw chunks, so the pipeline never hard-fails on a
 * missing key.
 *
 * For OpenAI-compatible gateways a key is still expected. Local endpoints
 * (Ollama, vLLM, LM Studio) don't check it, so set a placeholder:
 * `OPENAI_API_KEY=ollama`.
 */
export function hasLlmCredentials(): boolean {
  switch (RAG.llmProvider) {
    case "anthropic":
      return Boolean(process.env.ANTHROPIC_API_KEY || process.env.ANTHROPIC_AUTH_TOKEN);
    case "openai":
      return Boolean(process.env.OPENAI_API_KEY);
    default:
      return false;
  }
}

/**
 * Build the LangChain chat model for the configured backend.
 *
 * Only the Anthropic branch carries Anthropic-specific options — the refusal
 * fallback betas and the `createClient` workaround are meaningless on OpenAI,
 * and sending them there would just confuse the endpoint.
 */
export function createChatModel(): BaseChatModel {
  switch (RAG.llmProvider) {
    case "openai":
      // ChatOpenAI already reads OPENAI_API_KEY and OPENAI_BASE_URL itself, so
      // pointing this at DeepSeek / 通义千问 / GLM / Ollama needs no code change —
      // just those two env vars. Nothing is hardcoded here on purpose.
      //
      // maxTokens is left at 4096 rather than Claude's 16000: OpenAI-compatible
      // models cap out far lower per request (deepseek-chat and most Qwen
      // variants top out around 8K), and 4096 is plenty for a RAG answer.
      // No temperature — several compatible endpoints reject unknown params.
      return new ChatOpenAI({
        model: RAG.llmModel,
        maxTokens: 4096,
      });

    case "anthropic":
      return new ChatAnthropic({
        model: RAG.llmModel,
        maxTokens: 16000,
        // ChatAnthropic's constructor hard-requires ANTHROPIC_API_KEY and throws
        // without it, but the official SDK resolves the whole documented chain
        // (ANTHROPIC_API_KEY → ANTHROPIC_AUTH_TOKEN → profile). Supplying a client
        // factory skips that guard and lets the SDK do its own credential lookup —
        // which is what this environment needs, since it authenticates with
        // ANTHROPIC_AUTH_TOKEN against a local proxy via ANTHROPIC_BASE_URL.
        createClient: (options) => new Anthropic(options),
        // Opt into server-side refusal fallbacks: if the primary model declines a
        // request, the API re-runs it on a fallback model within the same call.
        betas: ["server-side-fallback-2026-07-01"],
        invocationKwargs: { fallbacks: "default" },
      });

    default:
      // Unreachable while hasLlmCredentials() gates every request, but a clear
      // error beats a confusing one if that ever changes.
      throw new Error(
        `不支持的 LLM_PROVIDER="${RAG.llmProvider}"，可选值：anthropic / openai`,
      );
  }
}
