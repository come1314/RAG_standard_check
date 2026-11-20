import fs from "node:fs/promises";
import path from "node:path";
import { PATHS, RAG } from "./config.js";

/**
 * Runtime model/key configuration, editable from the frontend.
 *
 * Values persist to `server/data/settings.json` and are pushed into
 * `process.env` on load and on every save. That indirection is deliberate:
 * `config.ts` reads the environment through getters and `llm.ts` builds the
 * model per request, so a save takes effect on the very next question — no
 * restart.
 *
 * Merge rule: a field present in the file wins; a field absent from it leaves
 * the environment untouched. Someone who never opens the UI keeps whatever
 * their shell already exported.
 */

export interface ProviderSettings {
  model?: string;
  baseUrl?: string;
  /** Never leaves the server — `publicSettings()` strips it before responding. */
  apiKey?: string;
}

export interface AppSettings {
  llmProvider: string;
  providers: Record<string, ProviderSettings>;
}

/** Which environment variables each provider's fields map onto. */
const ENV_KEYS: Record<string, { apiKey?: string; baseUrl?: string }> = {
  anthropic: { apiKey: "ANTHROPIC_API_KEY", baseUrl: "ANTHROPIC_BASE_URL" },
  openai: { apiKey: "OPENAI_API_KEY", baseUrl: "OPENAI_BASE_URL" },
};

/** Model id used when a provider has never been configured. */
const DEFAULT_MODEL: Record<string, string> = {
  anthropic: "claude-opus-5-5",
  openai: "gpt-4o-mini",
};

export const PROVIDERS = Object.keys(ENV_KEYS);

const SETTINGS_FILE = path.join(PATHS.dataDir, "settings.json");

function setOrDelete(key: string, value: string): void {
  if (value === "") delete process.env[key];
  else process.env[key] = value;
}

function isProvider(name: string): boolean {
  return Object.hasOwn(ENV_KEYS, name);
}

let cache: AppSettings | null = null;

/** Read the settings file. Missing or malformed means "not configured". */
async function read(): Promise<AppSettings> {
  if (cache) return cache;
  try {
    const parsed = JSON.parse(await fs.readFile(SETTINGS_FILE, "utf8"));
    cache = {
      llmProvider:
        typeof parsed?.llmProvider === "string" && isProvider(parsed.llmProvider)
          ? parsed.llmProvider
          : (process.env.LLM_PROVIDER ?? "anthropic"),
      providers:
        parsed && typeof parsed.providers === "object" && parsed.providers !== null
          ? (parsed.providers as Record<string, ProviderSettings>)
          : {},
    };
  } catch {
    cache = { llmProvider: process.env.LLM_PROVIDER ?? "anthropic", providers: {} };
  }
  return cache;
}

/** Push settings into the environment so the next request picks them up. */
function apply(settings: AppSettings): void {
  process.env.LLM_PROVIDER = settings.llmProvider;

  for (const [name, fields] of Object.entries(settings.providers ?? {})) {
    if (!isProvider(name)) continue;
    const keys = ENV_KEYS[name];
    // Only touch a variable the file actually has an opinion about — otherwise
    // an untouched provider keeps its shell-level credentials.
    if (fields.apiKey !== undefined && keys.apiKey) setOrDelete(keys.apiKey, fields.apiKey);
    if (fields.baseUrl !== undefined && keys.baseUrl) setOrDelete(keys.baseUrl, fields.baseUrl);
  }

  // Only set LLM_MODEL when a model was saved for the *active* provider.
  // Deleting it unconditionally would clobber a shell-exported LLM_MODEL for
  // anyone who has never opened the settings panel.
  const activeModel = settings.providers?.[settings.llmProvider]?.model;
  if (activeModel !== undefined) setOrDelete("LLM_MODEL", activeModel);
}

/** Called once at startup so a saved key is live before the first request. */
export async function loadSettings(): Promise<AppSettings> {
  const settings = await read();
  apply(settings);
  return settings;
}

async function persist(settings: AppSettings): Promise<void> {
  cache = settings;
  apply(settings);
  await fs.mkdir(PATHS.dataDir, { recursive: true });
  await fs.writeFile(SETTINGS_FILE, JSON.stringify(settings, null, 2), "utf8");
  try {
    await fs.chmod(SETTINGS_FILE, 0o600);
  } catch {
    /* Windows has no POSIX mode bits; `server/data/` is gitignored anyway. */
  }
}

export interface PartialUpdate {
  llmProvider?: string;
  providers?: Record<string, ProviderSettings>;
}

/**
 * Merge a partial update. An absent field means "unchanged", so the frontend
 * never has to echo back a secret it was never given in the first place.
 */
export async function updateSettings(update: PartialUpdate): Promise<AppSettings> {
  const current = await read();
  const provider = update.llmProvider ?? current.llmProvider;

  if (!isProvider(provider)) {
    throw new Error(`不支持的 LLM_PROVIDER="${provider}"，可选值：${PROVIDERS.join(" / ")}`);
  }

  const providers: Record<string, ProviderSettings> = { ...current.providers };
  for (const [name, fields] of Object.entries(update.providers ?? {})) {
    if (!isProvider(name)) continue;
    providers[name] = { ...providers[name], ...fields };
  }

  // Invariant: after any save the active provider has an explicit model, so a
  // later switch can never inherit the previous provider's model id.
  if (!providers[provider]?.model) {
    providers[provider] = { ...providers[provider], model: DEFAULT_MODEL[provider] ?? "" };
  }

  const next: AppSettings = { llmProvider: provider, providers };
  await persist(next);
  return next;
}

function mask(value: string | undefined): string | null {
  if (!value || value.length < 16) return null;
  return `${value.slice(0, 3)}…${value.slice(-4)}`;
}

export interface PublicProvider {
  model: string;
  baseUrl: string;
  apiKeySet: boolean;
  apiKeyMasked: string | null;
}

/**
 * What the frontend sees. Deliberately no raw key: only whether one exists and,
 * for long keys, enough of a handle to tell two of them apart.
 */
export async function publicSettings(): Promise<{
  llmProvider: string;
  providers: Record<string, PublicProvider>;
  embeddingModel: string;
}> {
  const settings = await read();
  const active = settings.llmProvider;

  const providers: Record<string, PublicProvider> = {};
  for (const name of PROVIDERS) {
    const saved = settings.providers?.[name] ?? {};
    const envKeys = ENV_KEYS[name];

    // Saved value wins; otherwise fall through to whatever the shell exported.
    const apiKey =
      saved.apiKey !== undefined ? saved.apiKey : (envKeys.apiKey ? process.env[envKeys.apiKey] : "") ?? "";
    const baseUrl = saved.baseUrl ?? (envKeys.baseUrl ? process.env[envKeys.baseUrl] : "") ?? "";
    // The environment's LLM_MODEL describes whichever provider was selected
    // when it was exported, so it must not be shown for the other one.
    const model =
      saved.model ?? (name === active ? process.env.LLM_MODEL : undefined) ?? DEFAULT_MODEL[name] ?? "";

    providers[name] = {
      model,
      baseUrl,
      apiKeySet: Boolean(apiKey),
      apiKeyMasked: mask(apiKey),
    };
  }

  return { llmProvider: active, providers, embeddingModel: RAG.embeddingModel };
}
