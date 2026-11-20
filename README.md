# 知识库问答 · Vue3 + LangChain RAG

一个把 RAG 流程完整跑通的示例：读取 `knowledgeFile/` 中的 `.docx` 文档，切分并向量化，
在本地向量库中检索，最后由 Claude 依据检索到的片段回答问题。

前端 Vue 3，后端 Node + LangChain，向量库 LanceDB（纯本地、单进程、无需独立服务），
嵌入模型在本地用 ONNX 跑，**不依赖任何在线向量服务**。

```
knowledgeFile/*.docx
      │
      ▼  DocxLoader (mammoth)          读取
   Document[]
      │
      ▼  RecursiveCharacterTextSplitter 切分（400 字 / 重叠 80）
   chunks[]
      │
      ▼  HuggingFaceTransformersEmbeddings 本地向量化（E5 + passage 前缀）
      │        └── server/data/hf-cache/  模型缓存，首次下载后离线可用
      ▼
 ┌─────────────────────────────┐
 │  LanceDB（本地向量库）        │  server/data/lancedb/
 └─────────────────────────────┘
      ▲        │
      │        ▼  query + "query: " 前缀 → cosine top-5
 用户提问 ──────────────────────► 检索到的片段
                                     │
                                     ▼  注入提示词（含来源编号）
                               Claude 生成回答 + 引用
```

## 快速开始

```bash
# 1. 安装依赖（根目录）
npm install
npm run install:all

# 2. 准备文档：examples/ 里有三份脱敏样例，复制进 knowledgeFile/ 后即可跑通全流程
cp examples/*.docx knowledgeFile/

# 3. 构建索引（首次会下载嵌入模型，约 118MB）
npm run ingest

# 4. 同时启动后端和前端
npm run dev
```

> Windows PowerShell 没有 `cp` 时可用：`Copy-Item examples/*.docx knowledgeFile/`。
> `knowledgeFile/` 被 `.gitignore` 忽略，你放进去的真实文档不会进版本库。

打开 <http://localhost:5173> 即可提问。后端单独跑在 <http://localhost:3001>。

只想跑后端或前端时：

```bash
npm run dev:server     # http://localhost:3001
npm run dev:web        # http://localhost:5173，/api 已代理到 3001
```

> 索引会在服务启动时自动检查：`knowledgeFile/` 的文件或索引配置有变化时自动重建。
> 也可以在页面左侧点「重建索引」，或执行 `npm run ingest`。

### 配置

复制 `server/.env.example` 为 `server/.env` 后按需修改。**不建 `.env` 也能直接跑**，
所有值都有可用的默认值。

关键项：

| 变量 | 默认值 | 说明 |
| --- | --- | --- |
| `LLM_PROVIDER` | `anthropic` | `anthropic`（Claude）或 `openai`（任何 OpenAI 兼容接口） |
| `LLM_MODEL` | 随 provider 变 | anthropic→`claude-opus-5-5`，openai→`gpt-4o-mini` |
| `ANTHROPIC_API_KEY` / `ANTHROPIC_AUTH_TOKEN` | — | anthropic 的凭证，两者都没有时仍可检索，回答退化为「仅返回原始片段」 |
| `ANTHROPIC_BASE_URL` | — | 走本地代理时用 |
| `OPENAI_API_KEY` | — | openai 的凭证，本地端点（Ollama/vLLM）填占位值即可 |
| `OPENAI_BASE_URL` | — | 指向 DeepSeek / 通义千问 / GLM / Ollama 的地址 |
| `EMBEDDING_MODEL` | `Xenova/multilingual-e5-small` | 本地嵌入模型，检索训练、支持中英文 |
| `HF_ENDPOINT` | `https://hf-mirror.com` | 默认走镜像；直连 huggingface.co 时改成 `https://huggingface.co` |
| `TOP_K` | `8` | 每次召回的片段数 |
| `CHUNK_SIZE` / `CHUNK_OVERLAP` | `250` / `50` | 切分参数，改完需 `npm run ingest` |
| `CHUNK_CONTEXT_PREFIX` | `true` | 给每块加 `《文档》· 小节标题` 前缀，改完需重建 |

### 换成自己的模型

**只有「回答生成」这一半需要 key**，读取、切分、向量化、向量检索全部本地跑，不碰网络。

#### 用 OpenAI 兼容接口（DeepSeek / 通义千问 / 智谱 GLM / Ollama / vLLM）

`server/.env`：

```bash
LLM_PROVIDER=openai
LLM_MODEL=deepseek-chat            # 各家模型名不同，必须填对
OPENAI_API_KEY=sk-xxxx
OPENAI_BASE_URL=https://api.deepseek.com/v1   # 用官方 OpenAI 就省略这行
```

`ChatOpenAI` 本身就会读 `OPENAI_API_KEY` 和 `OPENAI_BASE_URL`，所以**换供应商不需要改代码**，
改环境变量即可。本地 Ollama/vLLM 不校验 key，填 `OPENAI_API_KEY=ollama` 这类占位值就行。

#### 用 Anthropic 官方 key

```bash
LLM_PROVIDER=anthropic
LLM_MODEL=claude-opus-5-5
ANTHROPIC_API_KEY=sk-ant-xxxx      # 或改用 ANTHROPIC_AUTH_TOKEN + ANTHROPIC_BASE_URL
```

改完**重启服务**才生效。

> ⚠️ `dotenv` 默认**不会覆盖** shell 里已存在的同名变量。如果你的终端里已经 export 了
> `ANTHROPIC_AUTH_TOKEN` 之类，`.env` 里的值会被忽略 —— 换个新终端跑，或改到操作系统
> 环境变量里（Windows：设置 → 系统 → 高级系统设置 → 环境变量 → 用户变量）。


## 接口

| 方法 | 路径 | 说明 |
| --- | --- | --- |
| `GET` | `/api/status` | 索引状态、文件列表、分块数、当前参数 |
| `POST` | `/api/ingest` | 重建索引 |
| `POST` | `/api/chat` | `{ question, history? }` → `{ answer, sources, mode, latencyMs }` |
| `POST` | `/api/chat/stream` | 同样的请求体，改为 **SSE 流式**返回（前端默认走这条） |
| `GET` | `/api/settings` | 当前模型配置（**不含明文 key**，只有 `apiKeySet` 和掩码） |
| `POST` | `/api/settings` | 部分更新模型/key，**保存后立即生效，无需重启** |
| `POST` | `/api/upload` | multipart 字段名 `files`，传入 `.docx` |

`sources[]` 返回每条召回片段的来源文件、余弦相似度与内容摘要，前端会直接展示——
方便确认「回答确实来自这些文档」。

### 流式输出

`/api/chat/stream` 用 Server-Sent Events 返回，帧顺序是：

```
event: sources    检索结果（先到，不等模型）
event: delta      × N   每帧一小段正文
event: done       latencyMs
event: error      出错时
```

**`sources` 先于任何正文发出**是关键：检索是本地的、毫秒级，生成要几秒到几十秒。
先让来源卡片出现，用户在等待期就能看到「依据找对了没有」，而不是盯着空白气泡。

前端收到 `delta` 后不直接渲染，而是进一个缓冲队列，由 `requestAnimationFrame`
按固定速率吐出来 —— 网络是按 token 一小块一小块给的，直接渲染会很跳；
匀速吐字才像打字机。详见 `web/src/components/ChatPane.vue` 的 `tickTypewriter`。

客户端断开会通知服务端停止生成（`res.on("close")`），不会在后台继续烧 token。

### 前端能做什么

侧边栏里有两块是直接操作上面接口的：

- **模型设置** —— 切换 `anthropic` / `openai` 提供商，改模型名、接口地址、API Key。
  Key 只在你输入时才会发送，留空表示保持不变；保存后立刻生效。
- **上传文件** —— 选一个或多个 `.docx` 推到 `knowledgeFile/`，成功后自动触发重建索引，
  侧边栏的状态会跟着轮询更新。

上传的几道防线（`upload.ts`）：只收 `.docx`；文件名里的路径分隔符会被**清除而不是拒绝**
（所以 `..\..\x.docx` 落地时只会变成 `x.docx`）；用 ZIP 魔数校验内容，改扩展名的假文件会被挡下；
**同名文件不覆盖** —— 这些是原始资料，静默替换没有撤销的余地。


## 项目结构

```
examples/                   脱敏样例 .docx（复制进 knowledgeFile/ 后使用）
knowledgeFile/              待索引的 .docx 文档（不入库）
server/                     后端
  src/config.ts             路径与参数
  src/settings.ts           运行时模型/key 配置（前端可改，落盘到 data/settings.json）
  src/llm.ts                选择并构造聊天模型（anthropic / openai）
  src/embeddings.ts         本地嵌入（含 query/passage 前缀）
  src/ingest.ts             读取 → 切分 → 入库，以及过期检测
  src/upload.ts             上传文件的校验与落盘
  src/store.ts              LanceDB 检索
  src/answer.ts             检索 + 生成回答
  src/index.ts              Express 接口
  src/cli-ingest.ts         npm run ingest
  src/eval.ts               npm run eval
  data/                     向量库、模型缓存与 settings.json（不入库）
web/                        前端
  src/App.vue               布局与状态轮询
  src/components/ModelSettings.vue   模型 / key 设置面板
  src/components/IndexSidebar.vue    索引状态 + 文件列表 + 上传
  src/components/ChatPane.vue
```

## 效果评估

```bash
npm run eval    # 只跑检索，不调用模型，免费
```

两组指标，**缺一不可**：

**文件级**（9 条问答对，看「正确文档」排第几）

| 指标 | 结果 |
| --- | --- |
| top-1 命中 | 8/9（89%） |
| top-8 召回 | 9/9（100%） |
| MRR@8 | 0.944 |

**块级**（5 条人工标注，看 top-1 的**内容**是否切题）

| 指标 | 结果 |
| --- | --- |
| on-topic@1 | 5/5（100%） |

> 为什么需要第二组：查「Java 开发规范对异常处理有什么要求」时，召回的 5 条里 4 条都来自
> 正确文档，文件级指标报 **100% 满分**，但 top-1 实际是那份手册的**前言**，真正回答问题的
> 「2.2 异常处理」排在 **#6**、被 `TOP_K` 截掉了。文件级指标测不出这种失败。
> 所以 `eval.ts` 里加了 `ON_TOPIC` 用例，不通过时进程以退出码 1 结束，可以直接卡 CI。

演进过程（改参数后重新 `npm run ingest` 再 `npm run eval` 即可复现）：

| 版本 | MRR | 说明 |
| --- | --- | --- |
| paraphrase 模型 + 800 字 | 0.704 | 初版，top-1 仅 44% |
| E5 + 800 字 | 0.889 | 换成检索训练模型（需 `query:`/`passage:` 前缀） |
| E5 + 400 字 | 0.911 | 解决 512 token 静默截断 |
| **E5 + 250 字 + 小节前缀 + K=8** | **0.944** | 解决「泛化块霸榜、异文档混入」 |

## 几个踩过的坑

**嵌入模型上限是 512 token，不是字符。** `CHUNK_SIZE` 按字符计，但两个候选模型的
`max_position_embeddings` 都是 512，而中文约 1 字 1 token。800 字的分块**会静默截断**，
每块尾部内容根本没参与向量化。

**长而杂的段落会霸榜。** 分块越大越「平均化」：一份手册的**前言**（错误码、表结构、鉴权
什么都提一句）向量落在语料中心，跟谁都不远，于是查询「异常处理」时它反而拿了最高分，
而真正切题的章节排第 6 被截掉。两个针对性改动：分块降到 250 字让每块更同质；再给每块加
`《文档名》 · 小节标题` 前缀，让块自己知道出处。实测该查询下异文档混入 2→0、top-8 内
切题块 2→5。前缀开关见 `CHUNK_CONTEXT_PREFIX`。

**`ChatAnthropic` 只认 `ANTHROPIC_API_KEY`（仅 anthropic 分支）。** 它的构造函数没有
`ANTHROPIC_API_KEY` 就直接抛错，不理会 `ANTHROPIC_AUTH_TOKEN`。而官方 SDK 的凭证链是
`ANTHROPIC_API_KEY → ANTHROPIC_AUTH_TOKEN → 配置文件`。解法是传一个 `createClient`，
跳过它那道检查，把凭证解析交还给 SDK（本项目通过本地代理 + `ANTHROPIC_AUTH_TOKEN` 认证，
必须走这条路）。`openai` 分支没有这个问题 —— `ChatOpenAI` 不强制要求 key，
所以 Ollama 这类无鉴权端点能直接用。

**`claude-opus-5-5` 不能传 temperature。** 2026-10-21 起，传 `temperature`/`top_k`/`top_p`
会直接 400。`@langchain/anthropic` 已经知道这点，会自动剥离——所以这里**不要**手动设置。

**LangChain 的 LanceDB 封装读不到分数。** 它读 `score` 列，而 LanceDB 实际输出的是
`_distance`，于是 `similaritySearchWithScore` 永远返回 `undefined`。`store.ts` 因此直接
用 LanceDB 查询并读取 `_distance`，转成相似度返回；写入侧仍走 LangChain 的封装。

**拒答回退已开启（仅 anthropic 分支）。** `claude-opus-5-5` 的请求带了
`betas: ["server-side-fallback-2026-07-01"]` + `fallbacks: "default"`：模型如果触发安全
拒答，API 会在同一次调用里改用回退模型重试。不需要的话删掉 `llm.ts` 里这两行即可。
`openai` 分支不带这些参数 —— 那是 Anthropic 专有的，发给兼容端点只会添乱。

**镜像下载可能被截断。** hf-mirror 下载的 ONNX 偶尔会中途断掉，表现为
`Protobuf parsing failed`。比对 `Content-Length` 与本地文件大小，不一致就删掉重下
（`server/data/hf-cache/`）。

## 未做的事

- **回答是逐轮独立检索的**：每轮只把新的问题拿去召回，历史对话作为上下文传给模型，
  不会为历史轮次重新检索。对多数问答够用；如果要做多轮深度追问，需要按轮次分别召回。
- **未做重排序（rerank）**：召回 top-5 后没有二次精排。语料变大后加一个 reranker 收益明显。

> README 中的评估数值来自项目早期使用的真实语料，仓库自带的样例文档仅用于跑通流程，
> 不代表相同指标。换用自己的文档后请重新 `npm run ingest` 与 `npm run eval`。

## 贡献

见 [CONTRIBUTING.md](CONTRIBUTING.md)。

## 许可证

[MIT](LICENSE) © 2026 jousjd
