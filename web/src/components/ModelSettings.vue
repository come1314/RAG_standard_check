<script setup lang="ts">
import { computed, onMounted, ref } from "vue";
import { getSettings, saveSettings, type Settings, type SettingsUpdate } from "../api";

const emit = defineEmits<{ changed: [] }>();

interface ProviderForm {
  model: string;
  baseUrl: string;
  /** Edit box only. Cleared after every save so a raw key never lingers in the page. */
  apiKey: string;
  apiKeySet: boolean;
  apiKeyMasked: string | null;
}

const PROVIDER_LABELS: Record<string, string> = {
  anthropic: "Anthropic（Claude）",
  openai: "OpenAI 兼容接口",
};

const providers = ref<Record<string, ProviderForm>>({});
const active = ref("anthropic");
const loading = ref(true);
const saving = ref(false);
const error = ref<string | null>(null);
const notice = ref<string | null>(null);

const current = computed<ProviderForm | undefined>(() => providers.value[active.value]);
const keyHint = computed(() => {
  const form = current.value;
  if (!form) return "";
  if (!form.apiKeySet) return "未配置";
  return form.apiKeyMasked ? `已配置 ${form.apiKeyMasked}` : "已配置";
});

async function load(): Promise<void> {
  loading.value = true;
  error.value = null;
  try {
    const settings: Settings = await getSettings();
    active.value = settings.llmProvider;
    const next: Record<string, ProviderForm> = {};
    for (const [name, p] of Object.entries(settings.providers)) {
      next[name] = {
        model: p.model,
        baseUrl: p.baseUrl,
        apiKey: "",
        apiKeySet: p.apiKeySet,
        apiKeyMasked: p.apiKeyMasked,
      };
    }
    providers.value = next;
  } catch (cause) {
    error.value = cause instanceof Error ? cause.message : String(cause);
  } finally {
    loading.value = false;
  }
}

/**
 * Send only what changed. `apiKey` is included solely when the user typed
 * something new — leaving the box blank keeps whatever is already stored, so
 * the form never has to echo a secret back.
 */
async function submit(fields: {
  model?: string;
  baseUrl?: string;
  apiKey?: string;
}): Promise<void> {
  saving.value = true;
  error.value = null;
  notice.value = null;
  try {
    const update: SettingsUpdate = {
      llmProvider: active.value,
      providers: { [active.value]: fields },
    };
    const saved = await saveSettings(update);

    const form = providers.value[active.value];
    if (form) {
      form.apiKey = ""; // do not keep the secret in the DOM or in memory
      const reported = saved.providers[active.value];
      form.apiKeySet = reported?.apiKeySet ?? false;
      form.apiKeyMasked = reported?.apiKeyMasked ?? null;
      if (fields.model !== undefined && reported) form.model = reported.model;
      if (fields.baseUrl !== undefined && reported) form.baseUrl = reported.baseUrl;
    }

    notice.value = saved.applied.credentials
      ? `已保存并生效：${saved.llmProvider} / ${saved.providers[active.value]?.model ?? ""}`
      : "已保存，但该提供商仍没有可用凭证，回答会退化为「仅检索」。";
    emit("changed");
  } catch (cause) {
    error.value = cause instanceof Error ? cause.message : String(cause);
  } finally {
    saving.value = false;
  }
}

function save(): void {
  const form = current.value;
  if (!form) return;
  const model = form.model.trim();
  if (!model) {
    error.value = "模型名不能为空";
    return;
  }
  const fields: { model: string; baseUrl: string; apiKey?: string } = {
    model,
    baseUrl: form.baseUrl.trim(),
  };
  const typed = form.apiKey.trim();
  if (typed) fields.apiKey = typed;
  void submit(fields);
}

/** An empty apiKey in the payload means "delete the stored one". */
function clearKey(): void {
  void submit({ apiKey: "" });
}

onMounted(load);
</script>

<template>
  <section class="panel">
    <h2>模型设置</h2>

    <p v-if="loading" class="hint">读取中…</p>

    <template v-else-if="current">
      <div class="field">
        <label for="provider">提供商</label>
        <select id="provider" v-model="active" :disabled="saving">
          <option v-for="name in Object.keys(providers)" :key="name" :value="name">
            {{ PROVIDER_LABELS[name] ?? name }}
          </option>
        </select>
      </div>

      <div class="field">
        <label for="model">模型名</label>
        <input
          id="model"
          v-model="current.model"
          type="text"
          :disabled="saving"
          placeholder="例如 claude-opus-5-5 / deepseek-chat"
        />
      </div>

      <div class="field">
        <label for="baseUrl">接口地址</label>
        <input
          id="baseUrl"
          v-model="current.baseUrl"
          type="text"
          :disabled="saving"
          placeholder="留空使用官方默认地址"
        />
      </div>

      <div class="field">
        <label for="apiKey">API Key</label>
        <input
          id="apiKey"
          v-model="current.apiKey"
          type="password"
          autocomplete="off"
          :disabled="saving"
          :placeholder="current.apiKey ? '输入新 key 覆盖，留空则保持不变' : keyHint"
        />
        <div class="field-actions">
          <span class="hint">{{ keyHint }}</span>
          <button
            v-if="current.apiKeySet"
            class="link"
            type="button"
            :disabled="saving"
            @click="clearKey"
          >
            清除
          </button>
        </div>
      </div>

      <p v-if="error" class="form-error">{{ error }}</p>
      <p v-else-if="notice" class="form-ok">{{ notice }}</p>

      <button class="btn primary" type="button" :disabled="saving || loading" @click="save">
        <span v-if="saving" class="spinner"></span>
        {{ saving ? " 保存中…" : "保存设置" }}
      </button>
      <p class="hint">保存后立即生效，无需重启。Key 只存在本机 server/data/settings.json。</p>
    </template>
  </section>
</template>
