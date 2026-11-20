<script setup lang="ts">
import { ref } from "vue";
import { uploadFiles, type IndexStatus, type UploadResult } from "../api";
import ModelSettings from "./ModelSettings.vue";

defineProps<{
  status: IndexStatus | null;
  rebuilding: boolean;
}>();

const emit = defineEmits<{ rebuild: []; changed: [] }>();

const fileInput = ref<HTMLInputElement | null>(null);
const uploading = ref(false);
const uploadError = ref<string | null>(null);
const uploadResult = ref<UploadResult | null>(null);

function openPicker(): void {
  fileInput.value?.click();
}

async function onFilesSelected(event: Event): Promise<void> {
  const input = event.target as HTMLInputElement;
  const files = Array.from(input.files ?? []);
  // Reset first so choosing the same file again still fires a change event.
  input.value = "";
  if (files.length === 0) return;

  uploading.value = true;
  uploadError.value = null;
  uploadResult.value = null;
  try {
    uploadResult.value = await uploadFiles(files);
    // The file list, chunk count and staleness all moved — let the parent refetch.
    emit("changed");
  } catch (cause) {
    uploadError.value = cause instanceof Error ? cause.message : String(cause);
  } finally {
    uploading.value = false;
  }
}

function formatBuiltAt(value: string | null): string {
  if (!value) return "—";
  const date = new Date(value);
  return Number.isNaN(date.getTime()) ? "—" : date.toLocaleString("zh-CN");
}
</script>

<template>
  <aside class="sidebar">
    <div class="brand">
      <h1>知识库问答</h1>
      <p>Vue 3 + LangChain RAG · 本地向量检索</p>
    </div>

    <section class="panel">
      <h2>索引状态</h2>

      <div class="status-row">
        <span class="label">状态</span>
        <span class="value">
          <span v-if="status?.indexing" class="badge warn">构建中…</span>
          <span v-else-if="status?.ready" class="badge ok">就绪</span>
          <span v-else-if="status?.buildError" class="badge error">构建失败</span>
          <span v-else class="badge warn">未就绪</span>
        </span>
      </div>

      <template v-if="status">
        <div class="status-row">
          <span class="label">文档</span>
          <span class="value">{{ status.files.length }} 个</span>
        </div>
        <div class="status-row">
          <span class="label">分块数</span>
          <span class="value">{{ status.chunkCount }}</span>
        </div>
        <div class="status-row">
          <span class="label">构建时间</span>
          <span class="value">{{ formatBuiltAt(status.builtAt) }}</span>
        </div>
        <div class="status-row">
          <span class="label">回答方式</span>
          <span class="value">
            <span :class="['badge', status.llmMode === 'llm' ? 'ok' : 'warn']">
              {{ status.llmMode === "llm"
                ? `${status.llmProvider} · ${status.llmModel}`
                : `仅检索（${status.llmProvider} 未配置凭证）` }}
            </span>
          </span>
        </div>
        <div class="status-row">
          <span class="label">一致性</span>
          <span class="value">{{ status.stale.reason }}</span>
        </div>
      </template>

      <p v-if="status?.buildError" class="hint" style="color: var(--error)">
        {{ status.buildError }}
      </p>

      <button
        class="btn"
        style="margin-top: 10px"
        :disabled="rebuilding || status?.indexing"
        @click="emit('rebuild')"
      >
        <span v-if="rebuilding || status?.indexing" class="spinner"></span>
        {{ rebuilding || status?.indexing ? " 重建中…" : " 重建索引" }}
      </button>
      <p class="hint">knowledgeFile/ 中的文件变更后，服务重启时会自动重建，也可手动触发。</p>
    </section>

    <section class="panel">
      <h2>已收录文件</h2>
      <ul class="file-list">
        <li v-for="file in status?.files ?? []" :key="file">
          <span class="dot"></span>
          <span>{{ file }}</span>
        </li>
        <li v-if="!status || status.files.length === 0">
          <span class="dot" style="background: var(--border)"></span>
          <span style="color: var(--text-dim)">knowledgeFile/ 暂无 .docx 文件</span>
        </li>
      </ul>

      <input
        ref="fileInput"
        class="visually-hidden"
        type="file"
        accept=".docx"
        multiple
        @change="onFilesSelected"
      />
      <button
        class="btn"
        style="margin-top: 10px"
        type="button"
        :disabled="uploading"
        @click="openPicker"
      >
        <span v-if="uploading" class="spinner"></span>
        {{ uploading ? " 上传中…" : " 上传 .docx" }}
      </button>

      <p v-if="uploadError" class="form-error">{{ uploadError }}</p>
      <div v-else-if="uploadResult" class="upload-result">
        <p v-if="uploadResult.saved.length" class="form-ok">
          已加入 {{ uploadResult.saved.length }} 个文件，正在重建索引…
        </p>
        <p
          v-for="failure in uploadResult.failures"
          :key="failure.name"
          class="form-error"
        >
          {{ failure.name }}：{{ failure.reason }}
        </p>
        <p v-if="!uploadResult.saved.length && !uploadResult.failures.length" class="form-error">
          没有文件被保存
        </p>
      </div>

      <p class="hint">仅支持 .docx，单个上限 15MB。同名文件不会被覆盖。</p>
    </section>

    <ModelSettings @changed="emit('changed')" />

    <section class="panel">
      <h2>检索参数</h2>
      <div class="status-row">
        <span class="label">向量数据库</span>
        <span class="value">LanceDB（本地）</span>
      </div>
      <div class="status-row">
        <span class="label">召回条数 K</span>
        <span class="value">{{ status?.topK ?? "—" }}</span>
      </div>
      <div class="status-row">
        <span class="label">分块 / 重叠</span>
        <span class="value">
          {{ status ? `${status.chunkSize} / ${status.chunkOverlap}` : "—" }}
        </span>
      </div>
      <div class="status-row">
        <span class="label">嵌入模型</span>
        <span class="value">{{ status?.embeddingModel ?? "—" }}</span>
      </div>
    </section>
  </aside>
</template>
