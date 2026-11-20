<script setup lang="ts">
import { computed, onBeforeUnmount, onMounted, ref } from "vue";
import { getStatus, rebuildIndex, type IndexStatus } from "./api";
import IndexSidebar from "./components/IndexSidebar.vue";
import ChatPane from "./components/ChatPane.vue";

const status = ref<IndexStatus | null>(null);
const statusError = ref<string | null>(null);
const rebuilding = ref(false);

let poller: number | null = null;

async function refresh(): Promise<void> {
  try {
    status.value = await getStatus();
    statusError.value = null;
  } catch (cause) {
    statusError.value = cause instanceof Error ? cause.message : String(cause);
  }
}

async function rebuild(): Promise<void> {
  rebuilding.value = true;
  try {
    await rebuildIndex();
    statusError.value = null;
  } catch (cause) {
    statusError.value = cause instanceof Error ? cause.message : String(cause);
  } finally {
    rebuilding.value = false;
    await refresh();
  }
}

const blocked = computed(() => {
  if (statusError.value) return true;
  if (!status.value) return true;
  if (status.value.indexing) return true;
  if (status.value.buildError) return true;
  return status.value.chunkCount === 0;
});

const blockedReason = computed(() => {
  if (statusError.value) return `无法连接知识库服务：${statusError.value}`;
  if (!status.value) return "正在读取索引状态…";
  if (status.value.indexing) return "知识库索引正在构建，完成后即可提问。";
  if (status.value.buildError) return `索引构建失败：${status.value.buildError}`;
  return "知识库中还没有内容，请先在 knowledgeFile/ 放入文档并重建索引。";
});

onMounted(() => {
  void refresh();
  // Poll only while something is in flight, so the sidebar flips to ready
  // without anyone clicking anything.
  poller = window.setInterval(() => {
    if (!status.value || status.value.indexing || rebuilding.value) void refresh();
  }, 2000);
});

onBeforeUnmount(() => {
  if (poller !== null) window.clearInterval(poller);
});
</script>

<template>
  <div class="layout">
    <IndexSidebar
      :status="status"
      :rebuilding="rebuilding"
      @rebuild="rebuild"
      @changed="refresh"
    />
    <ChatPane :blocked="blocked" :blocked-reason="blockedReason" />
  </div>
</template>
