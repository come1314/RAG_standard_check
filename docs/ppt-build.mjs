import pptxgen from "pptxgenjs";
import { fileURLToPath } from "node:url";
import path from "node:path";

// ---------------------------------------------------------------------------
// 知识库问答智能体 · 汇报 PPT
// 设计参数取自 dataviz 参考调色板（已用 validate_palette.js 校验：
// #2a78d6 / #eb6834 / #1baf7a 在 #ffffff 表面上全部通过 CVD 与明度检查）。
// 单系列图表只用蓝 #2a78d6，不需要分类调色板；文字一律走文本色阶，不穿数据色。
// ---------------------------------------------------------------------------

// 输出到脚本所在目录；PPT_OUT 可覆盖（构建流程用它把结果送到别处）
const HERE = path.dirname(fileURLToPath(import.meta.url));
const OUT = process.env.PPT_OUT ?? path.join(HERE, "智能体汇报.pptx");

// 文本 / 图面色阶
const INK = "0B0B0B";      // primary ink
const INK2 = "52514E";     // secondary ink
const MUTED = "898781";    // 轴、刻度、次要标签
const LINE = "E1E0D9";     // hairline 网格 / 分隔线
const BASE = "C3C2B7";     // 基线 / 轴
const PLANE = "F9F9F7";    // 代码块底
const CARD = "FFFFFF";
const ACCENT = "2A78D6";   // 数据蓝（slot 1）
const ACCENT2 = "EB6834";  // 橙（slot 2，仅用于图示，不用于单系列数据）

const FONT = "Microsoft YaHei";
const MONO = "Consolas";

const pptx = new pptxgen();
pptx.layout = "LAYOUT_WIDE"; // 13.333 x 7.5
pptx.author = "standard_check";
pptx.company = "standard_check";
pptx.title = "知识库问答智能体 · 汇报材料";
pptx.subject = "RAG 知识库问答智能体：用法 / 逻辑 / 实现";
pptx.theme = { headFontFace: FONT, bodyFontFace: FONT };

const M = 0.75;
const W = 13.333 - M * 2; // 11.833
const BODY_Y = 1.78;

// --------------------------------------------------------------------------- helpers

function rect(slide, x, y, w, h, fill, lineColor) {
  slide.addShape(pptx.ShapeType.rect, {
    x, y, w, h,
    fill: { color: fill },
    line: lineColor ? { color: lineColor, width: 0.75 } : { type: "none" },
  });
}

function hairline(slide, x, y, w, color = LINE, width = 0.75) {
  slide.addShape(pptx.ShapeType.line, {
    x, y, w, h: 0,
    line: { color, width },
  });
}

/** 标准内页：眉标 + 标题 + 强调短线 + 页脚 */
function base(kicker, title, subtitle) {
  const slide = pptx.addSlide();
  slide.background = { color: "FFFFFF" };

  slide.addText(kicker.toUpperCase(), {
    x: M, y: 0.42, w: W, h: 0.26,
    fontFace: FONT, fontSize: 10, color: MUTED, charSpacing: 2, bold: true,
    margin: 0, valign: "middle",
  });

  slide.addText(title, {
    x: M, y: 0.7, w: W, h: 0.5,
    fontFace: FONT, fontSize: 25, color: INK, bold: true,
    margin: 0, valign: "middle",
  });

  rect(slide, M, 1.3, 0.72, 0.045, ACCENT);

  if (subtitle) {
    slide.addText(subtitle, {
      x: M, y: 1.38, w: W, h: 0.32,
      fontFace: FONT, fontSize: 11.5, color: INK2,
      margin: 0, valign: "middle",
    });
  }

  return slide;
}

function footer(slide, n) {
  hairline(slide, M, 7.06, W, LINE);
  slide.addText("知识库问答智能体 · 汇报材料", {
    x: M, y: 7.12, w: 5, h: 0.26,
    fontFace: FONT, fontSize: 8.5, color: MUTED, margin: 0, valign: "middle",
  });
  slide.addText(String(n).padStart(2, "0"), {
    x: M + W - 1, y: 7.12, w: 1, h: 0.26,
    fontFace: FONT, fontSize: 8.5, color: MUTED, align: "right", margin: 0, valign: "middle",
  });
}

/** 单段落项目符号；lead 加粗（主墨色），rest 常规（次墨色） */
function bullet(slide, x, y, w, h, lead, rest, opts = {}) {
  const runs = [];
  if (lead) runs.push({ text: lead, options: { bold: true, color: INK } });
  if (rest) runs.push({ text: rest, options: { bold: false, color: opts.restColor ?? INK2 } });
  slide.addText(runs, {
    x, y, w, h,
    fontFace: FONT, fontSize: opts.size ?? 12.5, valign: "top",
    margin: 0, lineSpacingMultiple: 1.22,
    bullet: { code: "25AA", indent: 14 },
    fit: "shrink",
  });
}

/** 无项目符号的普通段落 */
function para(slide, x, y, w, h, text, opts = {}) {
  slide.addText(text, {
    x, y, w, h,
    fontFace: FONT, fontSize: opts.size ?? 12, color: opts.color ?? INK2,
    bold: opts.bold ?? false, valign: opts.valign ?? "top",
    margin: 0, lineSpacingMultiple: opts.ls ?? 1.28, align: opts.align ?? "left",
    fit: "shrink",
  });
}

/** 浅底代码块 */
function code(slide, x, y, w, h, source) {
  rect(slide, x, y, w, h, PLANE);
  rect(slide, x, y, 0.035, h, ACCENT);
  slide.addText(source, {
    x: x + 0.18, y: y + 0.1, w: w - 0.32, h: h - 0.2,
    fontFace: MONO, fontSize: 9.5, color: "2C2C2A", valign: "top",
    margin: 0, lineSpacingMultiple: 1.18, fit: "shrink",
  });
}

/** 小标题条 */
function band(slide, x, y, w, text, color = ACCENT) {
  rect(slide, x, y + 0.03, 0.055, 0.2, color);
  slide.addText(text, {
    x: x + 0.17, y, w: w - 0.17, h: 0.28,
    fontFace: FONT, fontSize: 13, color: INK, bold: true,
    margin: 0, valign: "middle",
  });
}

/** 卡片（细环描边） */
function card(slide, x, y, w, h, fill = CARD) {
  rect(slide, x, y, w, h, fill, LINE);
}

function table(slide, rows, x, y, w, colW, opts = {}) {
  slide.addTable(rows, {
    x, y, w, colW,
    fontFace: FONT, fontSize: opts.size ?? 11,
    color: INK2, valign: "middle", align: "left",
    border: { type: "solid", color: LINE, pt: 0.75 },
    margin: [0.07, 0.11, 0.07, 0.11],
    rowH: opts.rowH ?? 0.36,
    autoPage: false,
  });
}

const th = (text) => ({
  text,
  options: { bold: true, color: INK, fill: { color: PLANE }, fontSize: 10.5 },
});

// ===========================================================================
// 01 封面
// ===========================================================================
{
  const slide = pptx.addSlide();
  slide.background = { color: "FFFFFF" };

  rect(slide, 0, 0, 13.333, 0.14, ACCENT);
  rect(slide, 0, 6.42, 13.333, 1.08, PLANE);

  slide.addText("项目汇报材料 · 2026-11-18", {
    x: 1.1, y: 1.72, w: 10, h: 0.3,
    fontFace: FONT, fontSize: 12, color: MUTED, charSpacing: 2, margin: 0,
  });

  slide.addText("知识库问答智能体", {
    x: 1.1, y: 2.16, w: 11, h: 0.92,
    fontFace: FONT, fontSize: 44, color: INK, bold: true, margin: 0, valign: "middle",
  });

  rect(slide, 1.1, 3.28, 1.5, 0.055, ACCENT);

  slide.addText("Vue 3 + LangChain RAG —— 用法 · 逻辑 · 实现", {
    x: 1.1, y: 3.56, w: 11, h: 0.42,
    fontFace: FONT, fontSize: 19, color: INK2, margin: 0, valign: "middle",
  });

  slide.addText(
    "读取 knowledgeFile/ 中的 .docx 规范文档，切分并向量化后存入本地向量库，\n" +
      "检索相关片段交给大模型归纳，答案标注 [1] [2] 引用，可点击跳回原文。",
    {
      x: 1.1, y: 4.24, w: 10.2, h: 0.9,
      fontFace: FONT, fontSize: 13, color: INK2, margin: 0, lineSpacingMultiple: 1.4,
    },
  );

  const chips = ["本地向量库 LanceDB", "本地 ONNX 嵌入 · 无需在线向量服务", "模型 / Key 前端可热切换"];
  chips.forEach((text, i) => {
    const x = 1.1 + i * 3.62;
    rect(slide, x, 5.62, 3.4, 0.46, CARD, LINE);
    slide.addText(text, {
      x: x + 0.16, y: 5.62, w: 3.1, h: 0.46,
      fontFace: FONT, fontSize: 10.5, color: INK2, margin: 0, valign: "middle",
    });
  });

  slide.addText("standard_check", {
    x: 1.1, y: 6.72, w: 6, h: 0.42,
    fontFace: FONT, fontSize: 11, color: MUTED, margin: 0, valign: "middle",
  });
}

// ===========================================================================
// 02 项目定位
// ===========================================================================
{
  const slide = base("01 / 项目定位", "一个可追溯的规范知识库问答智能体", null);

  card(slide, M, BODY_Y, W, 1.02, PLANE);
  slide.addText(
    [
      { text: "用户用自然语言提问 → 从规范 .docx 中检索相关片段 → 大模型归纳成答案 → 标注 ", options: { color: INK2 } },
      { text: "[1] [2]", options: { color: ACCENT, bold: true } },
      { text: " 引用编号 → 点击编号跳到对应原文片段。", options: { color: INK2 } },
    ],
    {
      x: M + 0.32, y: BODY_Y, w: W - 0.64, h: 1.02,
      fontFace: FONT, fontSize: 15, margin: 0, valign: "middle", lineSpacingMultiple: 1.35,
    },
  );

  const cols = [
    {
      title: "可追溯",
      body: "答案必须带引用编号。前端展示每条召回片段的来源文件、余弦相似度与内容摘要 —— 用来确认「回答确实来自这些文档」。",
    },
    {
      title: "本地可跑",
      body: "读取、切分、向量化、向量检索全部在本机执行，不依赖任何在线向量服务。嵌入模型首次下载后完全离线可用。",
    },
    {
      title: "流程优先",
      body: "先把端到端链路打通，再逐步优化检索质量。改动的好坏用可量化的指标判断，而不是看几个问答的主观感受。",
    },
  ];
  const cw = (W - 0.56) / 3;
  cols.forEach((c, i) => {
    const x = M + i * (cw + 0.28);
    card(slide, x, 3.14, cw, 2.12);
    rect(slide, x, 3.14, cw, 0.05, ACCENT);
    slide.addText(c.title, {
      x: x + 0.26, y: 3.36, w: cw - 0.52, h: 0.36,
      fontFace: FONT, fontSize: 16, color: INK, bold: true, margin: 0, valign: "middle",
    });
    para(slide, x + 0.26, 3.82, cw - 0.52, 1.28, c.body, { size: 11.5 });
  });

  card(slide, M, 5.52, W, 1.14, PLANE);
  rect(slide, M, 5.52, 0.055, 1.14, ACCENT2);
  slide.addText(
    [
      { text: "只有「回答生成」这半段需要 API Key。", options: { bold: true, color: INK } },
      {
        text: "  没有 key 时系统不报错，而是退化为「返回检索到的原始片段」，标注 mode: extractive，前端明确提示未经模型归纳 —— 检索那一半照常工作。",
        options: { color: INK2 },
      },
    ],
    {
      x: M + 0.3, y: 5.52, w: W - 0.6, h: 1.14,
      fontFace: FONT, fontSize: 12.5, margin: 0, valign: "middle", lineSpacingMultiple: 1.32,
    },
  );

  footer(slide, 2);
}

// ===========================================================================
// 03 整体架构
// ===========================================================================
{
  const slide = base("02 / 整体架构", "两趟流水线：离线索引 · 在线问答", null);

  // 5 个盒子铺满正文宽度，输入项写进 band 里 —— 侧边标签会把盒子挤出右边界
  const boxW = 2.12;
  const gap = 0.308;
  const boxH = 0.86;
  const startX = M;

  function pipeline(y, stages, accent) {
    stages.forEach((s, i) => {
      const x = startX + i * (boxW + gap);
      card(slide, x, y, boxW, boxH);
      rect(slide, x, y, boxW, 0.045, accent);
      slide.addText(
        [
          { text: s.k + "\n", options: { bold: true, color: INK, fontSize: 11.5 } },
          { text: s.v, options: { color: INK2, fontSize: 9.5 } },
        ],
        {
          x: x + 0.1, y: y + 0.06, w: boxW - 0.2, h: boxH - 0.12,
          fontFace: FONT, margin: 0, valign: "middle", align: "center",
          lineSpacingMultiple: 1.18, fit: "shrink",
        },
      );

      if (i < stages.length - 1) {
        slide.addShape(pptx.ShapeType.line, {
          x: x + boxW + 0.05, y: y + boxH / 2, w: gap - 0.1, h: 0,
          line: { color: BASE, width: 1.5, endArrowType: "triangle" },
        });
      }
    });
  }

  band(slide, M, BODY_Y + 0.02, W, "① 离线 · 索引构建　　输入　knowledgeFile/*.docx");
  pipeline(BODY_Y + 0.44, [
    { k: "读取", v: "DocxLoader (mammoth)" },
    { k: "切分", v: "250 字 / 重叠 50" },
    { k: "加前缀", v: "《文档》· 小节标题" },
    { k: "向量化", v: "本地 ONNX · passage: " },
    { k: "入库", v: "LanceDB 表 knowledge_chunks" },
  ], ACCENT);

  band(slide, M, 3.42, W, "② 在线 · 问答　　输入　用户提问");
  pipeline(3.84, [
    { k: "向量化", v: "query: 前缀 · 同一模型" },
    { k: "检索", v: "cosine top-K（K=8）" },
    { k: "组装提示词", v: "系统规则 + 历史 + 参考资料" },
    { k: "生成", v: "LLM 归纳 + [n] 引用" },
    { k: "渲染", v: "marked → 角标 → DOMPurify" },
  ], ACCENT2);

  card(slide, M, 5.32, W, 1.44, PLANE);
  slide.addText(
    [
      { text: "离线边界很清楚：", options: { bold: true, color: INK } },
      {
        text: "①～⑤ 全部在本机完成，只有第 ⑥ 步「生成」需要调用模型接口。\n",
        options: { color: INK2 },
      },
      { text: "同一套嵌入模型贯穿两趟流水线", options: { bold: true, color: INK } },
      {
        text: "（单例 PrefixedEmbeddings）：索引侧和查询侧的向量必须落在同一个空间，换模型就必须重建索引 —— 这一点由过期检测强制保证。",
        options: { color: INK2 },
      },
    ],
    {
      x: M + 0.3, y: 5.32, w: W - 0.6, h: 1.44,
      fontFace: FONT, fontSize: 12, margin: 0, valign: "middle", lineSpacingMultiple: 1.34,
    },
  );

  footer(slide, 3);
}

// ===========================================================================
// 04 技术选型
// ===========================================================================
{
  const slide = base("03 / 技术选型", "每一层都选了「本地可跑、不绑死一家」的方案", null);

  const rows = [
    [th("层"), th("选型"), th("版本"), th("为什么")],
    ["前端", "Vue 3 + Vite", "3.5 / 6.0", "组件化拆分侧边栏与对话区，开发体验好"],
    ["后端", "Node + Express + TypeScript", "4.21 / 5.7", "与前端同语言，LangChain.js 生态完整"],
    ["编排", "LangChain.js", "v1.x", "文档加载器、切分器、模型抽象统一到一个接口"],
    ["向量库", "LanceDB", "0.19.1", "嵌入式、单进程、无需独立服务，Windows 有预编译二进制"],
    ["嵌入模型", "Xenova/multilingual-e5-small", "—", "检索任务训练、支持中英文、ONNX 本地推理，约 118MB"],
    ["文档解析", "mammoth", "1.13", "读 .docx，正文与表格都能取到"],
    ["生成模型", "Claude / 任意 OpenAI 兼容接口", "可切换", "前端可热切换供应商，不绑死一家"],
  ];
  table(slide, rows, M, BODY_Y + 0.12, W, [1.5, 3.25, 1.25, 5.833], { rowH: 0.48, size: 11.5 });

  card(slide, M, 5.86, W, 0.94, PLANE);
  rect(slide, M, 5.86, 0.055, 0.94, ACCENT);
  slide.addText(
    [
      { text: "换成自己的模型不需要改代码：", options: { bold: true, color: INK } },
      {
        text: "ChatOpenAI 本身就读 OPENAI_API_KEY 和 OPENAI_BASE_URL，改环境变量就能指向 DeepSeek / 通义千问 / GLM / Ollama / vLLM。本地端点不校验 key，填占位值即可。",
        options: { color: INK2 },
      },
    ],
    {
      x: M + 0.3, y: 5.86, w: W - 0.6, h: 0.94,
      fontFace: FONT, fontSize: 12, margin: 0, valign: "middle", lineSpacingMultiple: 1.3,
    },
  );

  footer(slide, 4);
}

// ===========================================================================
// 05 用法：启动与配置
// ===========================================================================
{
  const slide = base("04 / 用法", "三条命令启动", null);

  band(slide, M, BODY_Y, 6.4, "启动");
  code(slide, M, BODY_Y + 0.42, 6.4, 1.32,
    "npm install && npm run install:all   # 安装依赖（根目录）\n" +
      "npm run ingest                       # 构建索引（首次下载嵌入模型 ~118MB）\n" +
      "npm run dev                          # 同时启动后端 3001 与前端 5173");

  para(slide, M, BODY_Y + 1.92, 6.4, 0.62,
    "打开 http://localhost:5173 即可提问。只想跑其中一半：npm run dev:server / npm run dev:web。",
    { size: 11.5 });

  band(slide, M, BODY_Y + 2.68, 6.4, "索引自动过期检测");
  bullet(slide, M, BODY_Y + 3.1, 6.4, 0.62, "文件变化：", "knowledgeFile/ 增删改任一文件都会触发重建。");
  bullet(slide, M, BODY_Y + 3.72, 6.4, 0.62, "配置变化：", "嵌入模型 / 分块参数 / 前缀开关改动同样触发重建。");

  const rows = [
    [th("变量"), th("默认值"), th("说明")],
    ["LLM_PROVIDER", "anthropic", "anthropic / openai 兼容"],
    ["LLM_MODEL", "随 provider 变", "opus-5-5 / gpt-4o-mini"],
    ["OPENAI_BASE_URL", "—", "DeepSeek / Qwen / GLM / Ollama"],
    ["EMBEDDING_MODEL", "e5-small", "本地嵌入模型，不走网络"],
    ["TOP_K", "8", "每次召回片段数"],
    ["CHUNK_SIZE / OVERLAP", "250 / 50", "切分参数，改完需重建"],
    ["CHUNK_CONTEXT_PREFIX", "true", "加《文档》· 小节标题前缀"],
  ];
  table(slide, rows, M + 6.85, BODY_Y + 0.42, W - 6.85, [1.85, 1.15, 1.983], { rowH: 0.46, size: 9 });

  para(slide, M + 6.85, BODY_Y + 4.3, W - 6.85, 0.62,
    "完整配置见 server/.env.example —— 所有值都有可用默认值，不建 .env 也能直接跑。",
    { size: 10, color: MUTED });

  footer(slide, 5);
}

// ===========================================================================
// 06 用法：前端能力
// ===========================================================================
{
  const slide = base("05 / 用法", "前端能做什么", "侧边栏直接操作后端接口，不需要碰命令行");

  const items = [
    {
      t: "切换模型与 Key",
      d: "选 anthropic / openai，改模型名、接口地址、API Key。保存后立即生效，无需重启服务。Key 只在输入时才发送，留空表示保持不变；输入框是密码型，保存后立即从内存清除。",
    },
    {
      t: "上传新文件",
      d: "选一个或多个 .docx 推到 knowledgeFile/，成功后自动触发重建索引，侧边栏状态轮询更新 —— 整个过程不需要重启或执行命令。",
    },
    {
      t: "点击引用跳转原文",
      d: "答案里的 [1] 是可点击角标，点击后滚动到对应来源卡片并高亮。卡片显示来源文件、余弦相似度与内容摘要。",
    },
    {
      t: "Markdown 排版",
      d: "模型输出的粗体、列表、表格、代码块渲染成富文本，而不是一堆 ** 星号；引用角标单独处理成上标徽章。",
    },
  ];

  const cw = (W - 0.32) / 2;
  items.forEach((it, i) => {
    const x = M + (i % 2) * (cw + 0.32);
    const y = BODY_Y + 0.18 + Math.floor(i / 2) * 2.32;
    card(slide, x, y, cw, 2.06);
    rect(slide, x, y, 0.05, 2.06, i % 2 === 0 ? ACCENT : ACCENT2);
    slide.addText(`${i + 1}`, {
      x: x + 0.3, y: y + 0.26, w: 0.5, h: 0.42,
      fontFace: FONT, fontSize: 20, color: i % 2 === 0 ? ACCENT : ACCENT2, bold: true,
      margin: 0, valign: "middle",
    });
    slide.addText(it.t, {
      x: x + 0.88, y: y + 0.26, w: cw - 1.16, h: 0.42,
      fontFace: FONT, fontSize: 16, color: INK, bold: true, margin: 0, valign: "middle",
    });
    para(slide, x + 0.3, y + 0.86, cw - 0.6, 1.02, it.d, { size: 11.5 });
  });

  footer(slide, 6);
}

// ===========================================================================
// 07 接口一览
// ===========================================================================
{
  const slide = base("06 / 接口", "六个接口，前后端各司其职", null);

  const rows = [
    [th("方法"), th("路径"), th("说明")],
    ["GET", "/api/status", "索引状态、文件列表、分块数、当前参数"],
    ["POST", "/api/ingest", "重建索引"],
    ["POST", "/api/chat", "{ question, history? } → { answer, sources, mode, latencyMs }"],
    ["GET", "/api/settings", "当前模型配置（不含明文 key，只有 apiKeySet 和掩码）"],
    ["POST", "/api/settings", "部分更新模型 / key，保存后立即生效，无需重启"],
    ["POST", "/api/upload", "multipart 字段名 files，传入 .docx"],
  ];
  table(slide, rows, M, BODY_Y + 0.2, W, [1.15, 2.55, 8.133], { rowH: 0.48, size: 11.5 });

  band(slide, M, 5.52, W, "为什么 settings 是「部分更新」");
  para(slide, M, 5.96, W, 0.9,
    "字段缺失 = 保持不变。前端从没拿到过 key，就不用回传 key —— 也就不存在「不小心把空字符串覆盖掉真 key」的问题。" +
      "GET /api/settings 只返回 apiKeySet 和掩码（前 3 + 后 4），明文只落在 gitignore 的 server/data/settings.json 里。",
    { size: 12 });

  footer(slide, 7);
}

// ===========================================================================
// 08 核心逻辑：索引构建
// ===========================================================================
{
  const slide = base("07 / 核心逻辑", "索引构建：从 .docx 到向量", null);

  band(slide, M, BODY_Y, 6.5, "构建流水线");
  code(slide, M, BODY_Y + 0.42, 6.5, 2.34,
    "读取 .docx  →  提取正文与小节标题\n" +
      "   →  按 250 字切分（重叠 50）\n" +
      "   →  每块加《文档名》· 小节标题 前缀\n" +
      "   →  本地 ONNX 向量化（passage: 前缀）\n" +
      "   →  写入 LanceDB 表 knowledge_chunks\n" +
      "   →  写 index-meta.json（文件指纹 + 索引配置）");

  bullet(slide, M, BODY_Y + 2.96, 6.5, 0.66, "切分器按中文标点断句",
    "：分隔符里把「。！？；…，、」排在空白之前，分块落在自然句边界上，不会把一句话拦腰截断。");
  bullet(slide, M, BODY_Y + 3.66, 6.5, 0.66, "元数据压平成标量",
    "：只留 source / chunk_index，因为 LanceDB 把元数据存成表列，嵌套对象映射不上去。");

  band(slide, M + 6.95, BODY_Y, W - 6.95, "过期检测：同时比较两件事", ACCENT2);
  card(slide, M + 6.95, BODY_Y + 0.42, W - 6.95, 1.12);
  slide.addText(
    [
      { text: "① 文件指纹\n", options: { bold: true, color: INK, fontSize: 12.5 } },
      { text: "文件名 + 大小 + 修改时间。增删改任一文件都算过期。", options: { color: INK2, fontSize: 11 } },
    ],
    { x: M + 7.2, y: BODY_Y + 0.52, w: W - 7.45, h: 0.92, fontFace: FONT, margin: 0, valign: "middle", lineSpacingMultiple: 1.28 },
  );
  card(slide, M + 6.95, BODY_Y + 1.68, W - 6.95, 1.12);
  slide.addText(
    [
      { text: "② 索引配置\n", options: { bold: true, color: INK, fontSize: 12.5 } },
      { text: "嵌入模型、query/passage 前缀、chunkSize、chunkOverlap、contextPrefix。", options: { color: INK2, fontSize: 11 } },
    ],
    { x: M + 7.2, y: BODY_Y + 1.78, w: W - 7.45, h: 0.92, fontFace: FONT, margin: 0, valign: "middle", lineSpacingMultiple: 1.28 },
  );

  card(slide, M + 6.95, BODY_Y + 3.0, W - 6.95, 1.7, PLANE);
  rect(slide, M + 6.95, BODY_Y + 3.0, 0.055, 1.7, ACCENT2);
  slide.addText(
    [
      { text: "第 ② 点比看上去重要得多。\n", options: { bold: true, color: INK, fontSize: 12 } },
      {
        text: "换嵌入模型后，旧向量依然可读、依然能查，但它们的含义已经变了。没有这道检查，重启后会静默地在一个「用旧模型建的索引」上做新查询 —— 结果全错，却毫无提示。",
        options: { color: INK2, fontSize: 11 } },
    ],
    { x: M + 7.2, y: BODY_Y + 3.12, w: W - 7.45, h: 1.46, fontFace: FONT, margin: 0, valign: "middle", lineSpacingMultiple: 1.3 },
  );

  footer(slide, 8);
}

// ===========================================================================
// 09 核心逻辑：问答链路
// ===========================================================================
{
  const slide = base("08 / 核心逻辑", "问答链路：检索 → 归纳 → 引用", null);

  band(slide, M, BODY_Y, 7.1, "一次请求发生了什么");
  code(slide, M, BODY_Y + 0.42, 7.1, 2.34,
    "用户问题\n" +
      "   →  embedQuery(\"query: …\") 向量化\n" +
      "   →  LanceDB cosine 相似度 top-K（K=8）\n" +
      "   →  组装提示词：\n" +
      "        SystemMessage  严格只依据参考资料、必须标注 [n]、不足就说不知道\n" +
      "      + 最近 6 轮对话历史（HumanMessage / AIMessage 真实消息）\n" +
      "      + 本轮「参考资料」+「用户问题」\n" +
      "   →  LLM 生成\n" +
      "   →  返回 { answer, sources[], mode, model, latencyMs }");

  card(slide, M, BODY_Y + 3.42, 7.1, 1.5, PLANE);
  rect(slide, M, BODY_Y + 3.42, 0.055, 1.5, ACCENT);
  slide.addText(
    [
      { text: "为什么历史用真实消息而不是模板变量？\n", options: { bold: true, color: INK, fontSize: 12 } },
      {
        text: "模型生成的答案里经常出现 { } —— ChatPromptTemplate 会尝试插值并报错；而只含 system 的模板会让请求的 messages 成为空数组，被 API 直接拒绝。",
        options: { color: INK2, fontSize: 11 } },
    ],
    { x: M + 0.3, y: BODY_Y + 3.54, w: 6.5, h: 1.26, fontFace: FONT, margin: 0, valign: "middle", lineSpacingMultiple: 1.3 },
  );

  band(slide, M + 7.55, BODY_Y, W - 7.55, "降级路径：系统不会因为缺 key 崩溃", ACCENT2);
  const rows = [
    [th("状态"), th("行为")],
    ["有 key", "mode: llm，模型归纳回答，带引用"],
    ["无 key", "mode: extractive，返回原始片段并标注未经归纳"],
    ["索引构建中", "503，提示稍候重试"],
    ["索引不存在", "503，提示先建索引"],
  ];
  table(slide, rows, M + 7.55, BODY_Y + 0.42, W - 7.55, [1.45, 2.833], { rowH: 0.56, size: 10.5 });

  card(slide, M + 7.55, BODY_Y + 3.42, W - 7.55, 1.5);
  slide.addText(
    [
      { text: "对话历史只取最近 6 轮\n", options: { bold: true, color: INK, fontSize: 12 } },
      {
        text: "更早的轮次直接丢弃 —— 控制提示词长度，避免历史淹没本轮检索到的参考资料。",
        options: { color: INK2, fontSize: 11 } },
    ],
    { x: M + 7.8, y: BODY_Y + 3.54, w: W - 8.05, h: 1.26, fontFace: FONT, margin: 0, valign: "middle", lineSpacingMultiple: 1.3 },
  );

  footer(slide, 9);
}

// ===========================================================================
// 10 实现：本地嵌入
// ===========================================================================
{
  const slide = base("09 / 实现", "本地嵌入：E5 的 query / passage 前缀", null);

  para(slide, M, BODY_Y, W, 0.66,
    "E5 这类「检索训练」模型要求输入带角色前缀 —— 问题加 query: ，文档块加 passage: 。两边必须对称，否则分数没有意义。" +
      "为避免索引侧和查询侧用错前缀，它们共用同一个单例：",
    { size: 12.5 });

  code(slide, M, BODY_Y + 0.78, 7.35, 2.62,
    "class PrefixedEmbeddings extends HuggingFaceTransformersEmbeddings {\n" +
      "  override embedDocuments(texts: string[]) {\n" +
      "    return super.embedDocuments(texts.map(t => `${this.passagePrefix}${t}`));\n" +
      "  }\n" +
      "  override embedQuery(text: string) {\n" +
      "    return super.embedQuery(`${this.queryPrefix}${text}`);\n" +
      "  }\n" +
      "}\n" +
      "let instance: PrefixedEmbeddings | null = null;\n" +
      "export function getEmbeddings() {\n" +
      "  instance ??= new PrefixedEmbeddings({ model: RAG.embeddingModel, pretrainedOptions });\n" +
      "  return instance;\n" +
      "}");

  const notes = [
    ["单例，不是每次新建", "索引和查询必须用同一个模型实例，否则两批向量不在同一空间。"],
    ["q8 量化加载", "约 1/4 体积，CPU 上足够快；显式写出来同时压掉「dtype not specified」告警。"],
    ["缓存到 server/data/hf-cache/", "首次下载后完全离线，不碰网络。"],
  ];
  notes.forEach((n, i) => {
    const y = BODY_Y + 0.78 + i * 0.92;
    card(slide, M + 7.8, y, W - 7.8, 0.8);
    slide.addText(
      [
        { text: n[0] + "\n", options: { bold: true, color: INK, fontSize: 11.5 } },
        { text: n[1], options: { color: INK2, fontSize: 10 } },
      ],
      { x: M + 8.02, y: y + 0.04, w: W - 8.24, h: 0.72, fontFace: FONT, margin: 0, valign: "middle", lineSpacingMultiple: 1.24 },
    );
  });

  card(slide, M, 5.42, W, 1.28, PLANE);
  rect(slide, M, 5.42, 0.055, 1.28, ACCENT2);
  slide.addText(
    [
      { text: "前缀是硬约束，不是可选项。\n", options: { bold: true, color: INK, fontSize: 13 } },
      {
        text: "查询侧写 query: 、索引侧写 passage: ，配对只在这一处定义，不可能漂移。换模型时（例如 paraphrase-multilingual-MiniLM-L12-v2）把两个前缀置空即可 —— " +
          "前缀写错的后果不是报错，而是所有相似度分数一起失真，检索结果莫名其妙却查不出原因。",
        options: { color: INK2, fontSize: 11.5 } },
    ],
    { x: M + 0.3, y: 5.52, w: W - 0.6, h: 1.08, fontFace: FONT, margin: 0, valign: "middle", lineSpacingMultiple: 1.3 },
  );

  footer(slide, 10);
}

// ===========================================================================
// 11 实现：向量检索
// ===========================================================================
{
  const slide = base("10 / 实现", "向量检索：绕过 LangChain 的分数读取", null);

  card(slide, M, BODY_Y, W, 1.02, PLANE);
  rect(slide, M, BODY_Y, 0.055, 1.02, ACCENT2);
  slide.addText(
    [
      { text: "问题：", options: { bold: true, color: INK } },
      {
        text: "LangChain 的 LanceDB 封装 similaritySearchWithScore 读的是 score 列，而 LanceDB 实际输出的是 _distance —— 于是永远返回 undefined。写入侧仍走封装，读取侧绕开。",
        options: { color: INK2 },
      },
    ],
    { x: M + 0.3, y: BODY_Y, w: W - 0.6, h: 1.02, fontFace: FONT, fontSize: 12.5, margin: 0, valign: "middle", lineSpacingMultiple: 1.3 },
  );

  code(slide, M, BODY_Y + 1.26, 7.35, 2.2,
    "const rows = await table.query()\n" +
      "  .nearestTo(vector)\n" +
      "  .distanceType(\"cosine\")\n" +
      "  .limit(k)\n" +
      "  .toArray();\n\n" +
      "// cosine 距离 = 1 - cos(a,b)，所以相似度就是 1 - distance\n" +
      "similarity: Math.max(-1, Math.min(1, 1 - distance))");

  const notes = [
    ["距离转相似度", "LanceDB 给的是 cosine distance，减一取负后夹到 [-1, 1]，前端才好展示成百分比条。"],
    ["RESERVED_COLUMNS", "把 vector / text / _distance / _rowid / score 等系统列过滤掉，剩下的都当作元数据返回。"],
    ["前后一致", "写入用 LangChain 封装，读取用原生 API —— 各用各的长处，不让一层封装决定能不能拿到分数。"],
  ];
  notes.forEach((n, i) => {
    const y = BODY_Y + 1.26 + i * 0.78;
    card(slide, M + 7.8, y, W - 7.8, 0.66);
    slide.addText(
      [
        { text: n[0] + "  ", options: { bold: true, color: INK, fontSize: 11 } },
        { text: n[1], options: { color: INK2, fontSize: 9.5 } },
      ],
      { x: M + 8.02, y: y + 0.04, w: W - 8.24, h: 0.58, fontFace: FONT, margin: 0, valign: "middle", lineSpacingMultiple: 1.2 },
    );
  });

  band(slide, M, 5.62, W, "一条原则");
  para(slide, M, 6.04, W, 0.72,
    "遇到第三方封装不工作时，先定位是哪一层的问题，再决定绕开哪一层 —— 而不是把整个库换掉。这里只在读取侧用原生 API，" +
      "加载、切分、写入、模型抽象仍然吃 LangChain 的红利。",
    { size: 12 });

  footer(slide, 11);
}

// ===========================================================================
// 12 检索质量调优：失败案例
// ===========================================================================
{
  const slide = base("11 / 检索质量", "一个真实的失败案例", "文件级指标报 100% 满分，功能却是坏的");

  card(slide, M, BODY_Y + 0.12, 5.55, 1.56, PLANE);
  slide.addText(
    [
      { text: "提问\n", options: { color: MUTED, fontSize: 10.5, bold: true } },
      { text: "「Java 开发规范对异常处理有什么要求？」", options: { color: INK, fontSize: 15, bold: true } },
    ],
    { x: M + 0.3, y: BODY_Y + 0.24, w: 4.95, h: 1.32, fontFace: FONT, margin: 0, valign: "middle", lineSpacingMultiple: 1.35 },
  );

  band(slide, M, BODY_Y + 1.9, 5.55, "召回的 5 条实际长这样");
  const rows = [
    [th("排名"), th("内容"), th("判断")],
    ["#1", "该手册的前言（错误码、表结构、鉴权都提一句）", "错"],
    ["#2", "另一份文档的段落", "错"],
    ["#3", "手册的其他章节", "不切题"],
    ["#4", "手册的其他章节", "不切题"],
    ["#5", "手册的其他章节", "不切题"],
    ["#6", "2.2 异常处理 ← 真正回答问题的那一节", "被 K 截掉"],
  ];
  table(slide, rows, M, BODY_Y + 2.22, 5.55, [0.72, 3.53, 1.3], { rowH: 0.4, size: 10 });

  card(slide, M + 5.95, BODY_Y + 0.12, W - 5.95, 2.42);
  rect(slide, M + 5.95, BODY_Y + 0.12, 0.055, 2.42, ACCENT2);
  slide.addText(
    [
      { text: "为什么这很危险\n", options: { bold: true, color: INK, fontSize: 14 } },
      {
        text: "模型基于这种上下文给出的答案，听起来仍然很对 —— 因为《Java 开发规范》是公开资料。\n\n" +
          "它很可能是靠预训练记忆答的，而不是靠检索到的片段。答案没错，但与知识库无关 —— " +
          "这意味着换一份私有文档，同样的链路会给出一本正经的错误答案。",
        options: { color: INK2, fontSize: 11.5 } },
    ],
    { x: M + 6.25, y: BODY_Y + 0.26, w: W - 6.55, h: 2.14, fontFace: FONT, margin: 0, valign: "middle", lineSpacingMultiple: 1.32 },
  );

  card(slide, M + 5.95, BODY_Y + 2.72, W - 5.95, 2.06, PLANE);
  slide.addText(
    [
      { text: "两个成因\n", options: { bold: true, color: INK, fontSize: 14 } },
      {
        text: "① 嵌入模型上限是 512 token（中文约 1 字 1 token），800 字的分块被静默截断，尾部根本没参与向量化。\n\n" +
          "② 长而杂的段落向量落在语料中心，跟谁都不远 —— 于是「前言」这种什么都提一句的段落反而拿了最高分。",
        options: { color: INK2, fontSize: 11.5 } },
    ],
    { x: M + 6.25, y: BODY_Y + 2.86, w: W - 6.55, h: 1.78, fontFace: FONT, margin: 0, valign: "middle", lineSpacingMultiple: 1.32 },
  );

  footer(slide, 12);
}

// ===========================================================================
// 13 检索质量调优：两个改动 + MRR 图
// ===========================================================================
{
  const slide = base("11 / 检索质量", "两个改动，MRR 0.704 → 0.944", null);

  // ---- 左：柱状图（单系列 · 蓝 · 0 基线 · 数值直标在柱顶）----
  const cx = M + 0.52;          // 绘图区左
  const cy = 2.28;              // 绘图区顶
  const cw = 5.5;               // 绘图区宽
  const ch = 3.24;              // 绘图区高
  const baseY = cy + ch;        // 基线

  slide.addText("检索质量演进 · MRR@8（9 条问答对）", {
    x: M, y: BODY_Y + 0.02, w: 6.4, h: 0.3,
    fontFace: FONT, fontSize: 12, color: INK, bold: true, margin: 0, valign: "middle",
  });

  // 网格线 + y 轴刻度（0 / 0.25 / 0.50 / 0.75 / 1.00）
  for (const t of [0, 0.25, 0.5, 0.75, 1]) {
    const y = baseY - ch * t;
    if (t > 0) hairline(slide, cx, y, cw, LINE, 0.75);
    slide.addText(t.toFixed(2), {
      x: cx - 0.52, y: y - 0.11, w: 0.44, h: 0.22,
      fontFace: FONT, fontSize: 8, color: MUTED, align: "right", margin: 0, valign: "middle",
    });
  }
  // 基线
  hairline(slide, cx, baseY, cw, BASE, 1);

  const data = [
    { label: "paraphrase 模型\n800 字", v: 0.704 },
    { label: "E5\n800 字", v: 0.889 },
    { label: "E5\n400 字", v: 0.911 },
    { label: "E5 250 字\n+ 小节前缀", v: 0.944 },
  ];
  const bandW = cw / data.length;
  const barW = 0.44;            // 窄柱：明显窄于 band，留足空气
  const radius = 0.055;         // 圆角数据端（约 4px）

  data.forEach((d, i) => {
    const x = cx + i * bandW + (bandW - barW) / 2;
    const h = ch * d.v;
    const y = baseY - h;

    // 圆角柱 + 底部补块把下端切平（圆角只在数据端）
    slide.addShape(pptx.ShapeType.roundRect, {
      x, y, w: barW, h,
      rectRadius: radius,
      fill: { color: ACCENT }, line: { type: "none" },
    });
    rect(slide, x, baseY - radius * 2, barW, radius * 2, ACCENT);

    // 柱顶直标数值（文本用文本色阶，不穿数据色）
    slide.addText(d.v.toFixed(3), {
      x: x - 0.24, y: y - 0.3, w: barW + 0.48, h: 0.26,
      fontFace: FONT, fontSize: 10.5, color: INK, bold: true, align: "center",
      margin: 0, valign: "middle",
    });

    // 分类标签
    slide.addText(d.label, {
      x: x - 0.32, y: baseY + 0.1, w: barW + 0.64, h: 0.52,
      fontFace: FONT, fontSize: 8.5, color: MUTED, align: "center",
      margin: 0, valign: "top", lineSpacingMultiple: 1.15,
    });
  });

  // 单系列 → 不放图例；标题已说明画的是什么
  slide.addText("纵轴 0 → 1.00，未截断。单一系列，柱色统一，数值直接标注在柱顶。", {
    x: M, y: 6.24, w: 6.4, h: 0.28,
    fontFace: FONT, fontSize: 8.5, color: MUTED, margin: 0, valign: "middle",
  });

  // ---- 右：两个改动 ----
  band(slide, M + 6.75, BODY_Y + 0.02, W - 6.75, "改动 ①  CHUNK_SIZE 800 → 250");
  para(slide, M + 6.75, BODY_Y + 0.4, W - 6.75, 0.98,
    "嵌入模型上限是 512 token，不是字符。800 字分块静默截断尾部，那部分内容从来没参与过向量化。" +
      "降到 250 后每块都在限内，也更同质。",
    { size: 11 });

  band(slide, M + 6.75, BODY_Y + 1.5, W - 6.75, "改动 ②  加上下文前缀", ACCENT2);
  code(slide, M + 6.75, BODY_Y + 1.88, W - 6.75, 0.52,
    "《${title}》${heading ? ` · ${heading.text}` : \"\"}\\n${text}");
  para(slide, M + 6.75, BODY_Y + 2.56, W - 6.75, 0.72,
    "让每一块自己知道出处。小节标题只认编号行和「第X章节」，普通句子哪怕以数字开头也不会误判。",
    { size: 11 });

  card(slide, M + 6.75, BODY_Y + 3.48, W - 6.75, 1.34, PLANE);
  rect(slide, M + 6.75, BODY_Y + 3.48, 0.055, 1.34, ACCENT);
  slide.addText(
    [
      { text: "实测（异常处理那条查询）\n", options: { bold: true, color: INK, fontSize: 12 } },
      { text: "异文档混入 2 → 0 · top-8 内切题块 2 → 5 · top-1 从「前言」变成「2.2 异常处理」", options: { color: INK2, fontSize: 11 } },
    ],
    { x: M + 7.05, y: BODY_Y + 3.58, w: W - 7.35, h: 1.14, fontFace: FONT, margin: 0, valign: "middle", lineSpacingMultiple: 1.3 },
  );

  footer(slide, 13);
}

// ===========================================================================
// 14 效果评估
// ===========================================================================
{
  const slide = base("12 / 效果评估", "两组指标，缺一不可", "npm run eval —— 只跑检索，不调用模型，免费");

  const tiles = [
    { label: "top-1 命中", value: "8 / 9", sub: "89% · 文件级" },
    { label: "top-8 召回", value: "9 / 9", sub: "100% · 文件级" },
    { label: "MRR@8", value: "0.944", sub: "平均倒数排名" },
    { label: "on-topic@1", value: "5 / 5", sub: "100% · 块级" },
  ];
  const tw = (W - 0.84) / 4;
  tiles.forEach((t, i) => {
    const x = M + i * (tw + 0.28);
    card(slide, x, BODY_Y + 0.14, tw, 1.62);
    rect(slide, x, BODY_Y + 0.14, tw, 0.05, ACCENT);
    slide.addText(t.label, {
      x: x + 0.26, y: BODY_Y + 0.36, w: tw - 0.52, h: 0.3,
      fontFace: FONT, fontSize: 11, color: MUTED, margin: 0, valign: "middle",
    });
    slide.addText(t.value, {
      x: x + 0.26, y: BODY_Y + 0.72, w: tw - 0.52, h: 0.58,
      fontFace: FONT, fontSize: 30, color: INK, bold: true, margin: 0, valign: "middle",
    });
    slide.addText(t.sub, {
      x: x + 0.26, y: BODY_Y + 1.32, w: tw - 0.52, h: 0.28,
      fontFace: FONT, fontSize: 9.5, color: INK2, margin: 0, valign: "middle",
    });
  });

  band(slide, M, 3.86, 5.9, "文件级 vs 块级");
  bullet(slide, M, 4.28, 5.9, 0.72, "文件级（9 条问答对）",
    "：看「正确文档」排第几。回答的是「有没有找对文档」。");
  bullet(slide, M, 5.0, 5.9, 0.72, "块级（5 条人工标注）",
    "：看 top-1 的内容是否切题。回答的是「找到的那一段能不能回答问题」。");

  card(slide, M + 6.3, 3.86, W - 6.3, 2.86, PLANE);
  rect(slide, M + 6.3, 3.86, 0.055, 2.86, ACCENT2);
  slide.addText(
    [
      { text: "教训：指标设计错了，比没有指标更危险\n\n", options: { bold: true, color: INK, fontSize: 13 } },
      {
        text: "那个「异常处理」的例子，文件级指标报 100% 满分 —— 但 top-1 是前言，真正的答案被 TOP_K 截掉了。指标是绿的，功能是坏的。\n\n" +
          "最初的块级指标同样失败：查「异常处理」时前言里恰好有「异常」两个字，于是 9/9 全对 —— 指标自己把失败判成了成功。\n\n" +
          "最终方案是人工标注的 ON_TOPIC 正则用例，明确写出「top-1 必须匹配 捕获异常|Throwable|预检查|RuntimeException|catch 时」，" +
          "既排除前言也排除混进来的其他文档；不通过时进程退出码 1，可以直接卡 CI。",
        options: { color: INK2, fontSize: 11 } },
    ],
    { x: M + 6.6, y: 3.98, w: W - 6.9, h: 2.62, fontFace: FONT, margin: 0, valign: "top", lineSpacingMultiple: 1.28 },
  );

  footer(slide, 14);
}

// ===========================================================================
// 15 实现：模型 / Key 热更新
// ===========================================================================
{
  const slide = base("13 / 实现", "模型 / Key 热更新：改完不用重启", null);

  para(slide, M, BODY_Y, W, 0.56,
    "settings.ts 把配置写到 server/data/settings.json，每次保存都推回 process.env。配合两个设计，改动才能立刻生效：",
    { size: 12.5 });

  const two = [
    {
      t: "① config.ts 用 getter，不是常量",
      d: "llmProvider / llmModel 每次访问都重新读环境变量，而不是启动时求值一次存下来。",
    },
    {
      t: "② llm.ts 每个请求新建模型实例",
      d: "不是启动时建一个单例一直用 —— 新配置在下一个问题上就生效。",
    },
  ];
  const cw2 = (W - 0.32) / 2;
  two.forEach((it, i) => {
    const x = M + i * (cw2 + 0.32);
    card(slide, x, BODY_Y + 0.7, cw2, 1.42);
    rect(slide, x, BODY_Y + 0.7, 0.05, 1.42, ACCENT);
    slide.addText(it.t, {
      x: x + 0.28, y: BODY_Y + 0.86, w: cw2 - 0.56, h: 0.36,
      fontFace: FONT, fontSize: 14, color: INK, bold: true, margin: 0, valign: "middle",
    });
    para(slide, x + 0.28, BODY_Y + 1.3, cw2 - 0.56, 0.68, it.d, { size: 11.5 });
  });

  code(slide, M, BODY_Y + 2.36, 6.5, 1.86,
    "export const RAG = {\n" +
      "  // getter，不是普通值：settings.ts 每次保存都会重写\n" +
      "  // LLM_PROVIDER / LLM_MODEL，读时才取值 = 无需重启\n" +
      "  get llmProvider(): string { return currentProvider(); },\n" +
      "  get llmModel(): string { return currentModel(); },\n" +
      "};");

  band(slide, M + 6.95, BODY_Y + 2.36, W - 6.95, "合并规则：字段缺失 = 保持不变", ACCENT2);
  para(slide, M + 6.95, BODY_Y + 2.8, W - 6.95, 1.42,
    "前端从没拿到过 key，就不用回传 key —— 也就不存在「不小心把空字符串覆盖掉真 key」的问题。\n\n" +
      "只设置当前 provider 的 LLM_MODEL，避免切换供应商时继承上一家的模型名。",
    { size: 11.5 });

  card(slide, M, 6.02, W, 0.86, PLANE);
  rect(slide, M, 6.02, 0.055, 0.86, ACCENT);
  slide.addText(
    [
      { text: "不变式：", options: { bold: true, color: INK } },
      {
        text: "任何一次保存之后，当前 provider 一定有一个显式的 model —— 这样切换供应商永远不会继承上一家的模型 id。",
        options: { color: INK2 },
      },
    ],
    { x: M + 0.3, y: 6.02, w: W - 0.6, h: 0.86, fontFace: FONT, fontSize: 12, margin: 0, valign: "middle" },
  );

  footer(slide, 15);
}

// ===========================================================================
// 16 实现：渲染与安全边界
// ===========================================================================
{
  const slide = base("14 / 实现", "渲染与安全边界", null);

  band(slide, M, BODY_Y, 6.85, "Markdown 渲染：消毒必须放在最后");
  code(slide, M, BODY_Y + 0.42, 6.85, 1.86,
    "export function renderMarkdown(source: string): string {\n" +
      "  const html = marked.parse(source, { async: false });\n" +
      "  const withCitations = html.replace(/\\[(\\d{1,2})\\]/g,\n" +
      "    '<sup class=\"cite\" data-cite=\"$1\">$1</sup>');\n" +
      "  return DOMPurify.sanitize(withCitations);  // ← 最后\n" +
      "}");

  bullet(slide, M, BODY_Y + 2.5, 6.85, 0.66, "顺序不能换",
    "：先 Markdown 解析 → 再把 [1] 转成角标 → 最后 DOMPurify 消毒。");
  bullet(slide, M, BODY_Y + 3.16, 6.85, 0.66, "为什么要消毒",
    "：模型输出的内容、以及文档里回显进提示词的内容（文档是用户上传的），都要过一遍才能进 v-html。");

  band(slide, M + 7.3, BODY_Y, W - 7.3, "上传的几道防线", ACCENT2);
  const rows = [
    [th("防线"), th("实现")],
    ["只收 .docx", "SUPPORTED_EXTENSIONS 单点定义，加载与上传共用"],
    ["路径穿越", "文件名分隔符清除而非拒绝，..\\..\\x.docx 落地就是 x.docx"],
    ["伪装扩展名", "校验 ZIP 魔数 PK\\x03\\x04，改名的假文件挡下"],
    ["同名覆盖", "不覆盖，返回失败 —— 原始资料没有撤销余地"],
    ["中文文件名", "multer defParamCharset: utf8（默认 latin1 会解成乱码）"],
    ["大小 / 数量", "单文件 15MB，单次 8 个"],
  ];
  table(slide, rows, M + 7.3, BODY_Y + 0.42, W - 7.3, [1.55, 2.983], { rowH: 0.48, size: 10 });

  card(slide, M, 5.74, W, 1.04, PLANE);
  rect(slide, M, 5.74, 0.055, 1.04, ACCENT);
  slide.addText(
    [
      { text: "Key 不外泄：", options: { bold: true, color: INK } },
      {
        text: "GET /api/settings 只返回 apiKeySet 和掩码（前 3 + 后 4）；明文只落在 gitignore 的 server/data/settings.json；" +
          "API 默认只监听 127.0.0.1 —— 设置接口在这个端口上接收 Key，所以除非显式设 HOST=0.0.0.0，不对局域网开放。",
        options: { color: INK2 },
      },
    ],
    { x: M + 0.3, y: 5.74, w: W - 0.6, h: 1.04, fontFace: FONT, fontSize: 11.5, margin: 0, valign: "middle", lineSpacingMultiple: 1.28 },
  );

  footer(slide, 16);
}

// ===========================================================================
// 17 踩坑记录
// ===========================================================================
{
  const slide = base("15 / 踩坑", "踩过的坑与解法", "完整清单见 docs/智能体汇报.md 第八节");

  const rows = [
    [th("坑"), th("现象"), th("解法")],
    ["嵌入上限是 token 不是字符", "800 字分块静默截断，尾部没参与向量化", "CHUNK_SIZE 降到 250"],
    ["泛化段落霸榜", "「前言」这种什么都提一句的段落拿最高分", "分块变小 + 上下文前缀"],
    ["ChatAnthropic 只认 API_KEY", "用 ANTHROPIC_AUTH_TOKEN 直接抛 key not found", "传 createClient，交还给 SDK 凭证链"],
    ["claude-opus-5-5 拒绝采样参数", "传 temperature / top_k / top_p 直接 400", "不设这些参数"],
    ["LangChain 读不到 LanceDB 分数", "similaritySearchWithScore 恒返回 undefined", "直接查 LanceDB，读 _distance"],
    ["空 messages 被拒", "只有 system 的模板让 messages 成为空数组", "用真实 Message 对象"],
    ["镜像下载被截断", "Protobuf parsing failed", "比对 Content-Length，不一致就重下"],
    ["dotenv 不覆盖已有变量", ".env 里写了 key 却不生效", "换新终端，或写到系统环境变量"],
  ];
  table(slide, rows, M, BODY_Y + 0.16, W, [3.05, 4.65, 4.133], { rowH: 0.55, size: 10.5 });

  footer(slide, 17);
}

// ===========================================================================
// 18 后续 + 结语
// ===========================================================================
{
  const slide = base("16 / 规划", "未做的事 / 后续规划", null);

  const rows = [
    [th("项"), th("现状"), th("收益")],
    ["多轮检索", "每轮只用新问题召回，历史仅作上下文", "做多轮深度追问时按轮次分别召回"],
    ["流式输出", "接口一次性返回 JSON", "接 SSE 显著改善长回答的等待体验"],
    ["重排序 rerank", "召回 top-K 后无二次精排", "语料变大后收益明显"],
    ["混合检索", "纯向量", "向量 + 关键词互补，编号类查询更稳"],
    ["文件管理", "只能上传，界面上不能删除", "资料过期后无法清理"],
  ];
  table(slide, rows, M, BODY_Y + 0.12, W, [2.15, 4.75, 4.933], { rowH: 0.5, size: 11 });

  card(slide, M, 5.28, W, 1.5, PLANE);
  rect(slide, M, 5.28, 0.055, 1.5, ACCENT);
  slide.addText(
    [
      { text: "一句话总结\n", options: { bold: true, color: INK, fontSize: 14 } },
      {
        text: "链路已经跑通：文档 → 切分 → 本地向量化 → 本地向量库 → 检索 → 模型归纳 → 可点击引用。" +
          "真正花时间的不是把流程接起来，而是让「检索到的确实是能回答问题的那一段」—— " +
          "这件事只能靠可量化的指标和人工标注的用例来守住。",
        options: { color: INK2, fontSize: 12 } },
    ],
    { x: M + 0.3, y: 5.42, w: W - 0.6, h: 1.22, fontFace: FONT, margin: 0, valign: "middle", lineSpacingMultiple: 1.34 },
  );

  footer(slide, 18);
}

await pptx.writeFile({ fileName: OUT });
console.log("written:", OUT);
