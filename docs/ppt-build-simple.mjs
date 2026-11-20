import pptxgen from "pptxgenjs";
import { fileURLToPath } from "node:url";
import path from "node:path";

// ---------------------------------------------------------------------------
// 项目简介 PPT（轻量版）—— 背景 / 意义 / 用法 / 存在的问题
// 视觉体系与「智能体汇报.pptx」一致，但字号更大、每页信息更少、不放代码。
// 输出到脚本所在目录；PPT_OUT 可覆盖。
// ---------------------------------------------------------------------------

const HERE = path.dirname(fileURLToPath(import.meta.url));
const OUT = process.env.PPT_OUT ?? path.join(HERE, "项目简介.pptx");

const INK = "0B0B0B";
const INK2 = "52514E";
const MUTED = "898781";
const LINE = "E1E0D9";
const BASE = "C3C2B7";
const PLANE = "F9F9F7";
const ACCENT = "2A78D6";
const ACCENT2 = "EB6834";

const FONT = "Microsoft YaHei";
const MONO = "Consolas";

const pptx = new pptxgen();
pptx.layout = "LAYOUT_WIDE";
pptx.title = "知识库问答智能体 · 项目简介";
pptx.subject = "背景 / 意义 / 用法 / 存在的问题";
pptx.author = "standard_check";
pptx.theme = { headFontFace: FONT, bodyFontFace: FONT };

const M = 0.75;
const W = 13.333 - M * 2;
const BODY_Y = 1.86;

// ------------------------------------------------------------------ helpers

function rect(slide, x, y, w, h, fill, lineColor) {
  slide.addShape(pptx.ShapeType.rect, {
    x, y, w, h,
    fill: { color: fill },
    line: lineColor ? { color: lineColor, width: 0.75 } : { type: "none" },
  });
}

function hairline(slide, x, y, w, color = LINE, width = 0.75) {
  slide.addShape(pptx.ShapeType.line, { x, y, w, h: 0, line: { color, width } });
}

function base(kicker, title, subtitle) {
  const slide = pptx.addSlide();
  slide.background = { color: "FFFFFF" };

  slide.addText(kicker.toUpperCase(), {
    x: M, y: 0.44, w: W, h: 0.28,
    fontFace: FONT, fontSize: 11, color: MUTED, charSpacing: 2, bold: true,
    margin: 0, valign: "middle",
  });
  slide.addText(title, {
    x: M, y: 0.76, w: W, h: 0.58,
    fontFace: FONT, fontSize: 30, color: INK, bold: true, margin: 0, valign: "middle",
  });
  rect(slide, M, 1.44, 0.8, 0.05, ACCENT);
  if (subtitle) {
    slide.addText(subtitle, {
      x: M, y: 1.52, w: W, h: 0.34,
      fontFace: FONT, fontSize: 13, color: INK2, margin: 0, valign: "middle",
    });
  }
  return slide;
}

function footer(slide, n) {
  hairline(slide, M, 7.06, W);
  slide.addText("知识库问答智能体 · 项目简介", {
    x: M, y: 7.12, w: 5, h: 0.26,
    fontFace: FONT, fontSize: 9, color: MUTED, margin: 0, valign: "middle",
  });
  slide.addText(String(n).padStart(2, "0"), {
    x: M + W - 1, y: 7.12, w: 1, h: 0.26,
    fontFace: FONT, fontSize: 9, color: MUTED, align: "right", margin: 0, valign: "middle",
  });
}

function card(slide, x, y, w, h, fill = "FFFFFF") {
  rect(slide, x, y, w, h, fill, LINE);
}

function para(slide, x, y, w, h, text, opts = {}) {
  slide.addText(text, {
    x, y, w, h,
    fontFace: FONT, fontSize: opts.size ?? 15, color: opts.color ?? INK2,
    bold: opts.bold ?? false, valign: opts.valign ?? "top",
    margin: 0, lineSpacingMultiple: opts.ls ?? 1.42, align: opts.align ?? "left",
    fit: "shrink",
  });
}

function band(slide, x, y, w, text, color = ACCENT) {
  rect(slide, x, y + 0.035, 0.06, 0.24, color);
  slide.addText(text, {
    x: x + 0.2, y, w: w - 0.2, h: 0.32,
    fontFace: FONT, fontSize: 15, color: INK, bold: true, margin: 0, valign: "middle",
  });
}

/** 大号要点：序号 + 标题 + 说明 */
function point(slide, x, y, w, n, title, body, color = ACCENT) {
  slide.addText(n, {
    x, y: y + 0.02, w: 0.62, h: 0.5,
    fontFace: FONT, fontSize: 26, color, bold: true, margin: 0, valign: "top",
  });
  slide.addText(title, {
    x: x + 0.72, y: y + 0.06, w: w - 0.72, h: 0.4,
    fontFace: FONT, fontSize: 17, color: INK, bold: true, margin: 0, valign: "middle",
  });
  para(slide, x + 0.72, y + 0.56, w - 0.72, 0.9, body, { size: 13.5 });
}

function table(slide, rows, x, y, w, colW, opts = {}) {
  slide.addTable(rows, {
    x, y, w, colW,
    fontFace: FONT, fontSize: opts.size ?? 13,
    color: INK2, valign: "middle", align: "left",
    border: { type: "solid", color: LINE, pt: 0.75 },
    margin: [0.1, 0.16, 0.1, 0.16],
    rowH: opts.rowH ?? 0.62,
    autoPage: false,
  });
}

const th = (text) => ({
  text,
  options: { bold: true, color: INK, fill: { color: PLANE }, fontSize: 13 },
});

// ==========================================================================
// 01 封面
// ==========================================================================
{
  const slide = pptx.addSlide();
  slide.background = { color: "FFFFFF" };
  rect(slide, 0, 0, 13.333, 0.16, ACCENT);
  rect(slide, 0, 6.4, 13.333, 1.1, PLANE);

  slide.addText("项目简介 · 2026-11-18", {
    x: 1.1, y: 1.86, w: 10, h: 0.32,
    fontFace: FONT, fontSize: 13, color: MUTED, charSpacing: 2, margin: 0,
  });
  slide.addText("知识库问答智能体", {
    x: 1.1, y: 2.32, w: 11, h: 1.0,
    fontFace: FONT, fontSize: 50, color: INK, bold: true, margin: 0, valign: "middle",
  });
  rect(slide, 1.1, 3.52, 1.7, 0.06, ACCENT);
  slide.addText("把规范文档变成能问、能查、能核对的答案", {
    x: 1.1, y: 3.82, w: 11, h: 0.46,
    fontFace: FONT, fontSize: 21, color: INK2, margin: 0, valign: "middle",
  });

  const chips = ["答案带出处，可点击回原文", "本地运行，数据不出本机", "换模型不改代码"];
  chips.forEach((t, i) => {
    const x = 1.1 + i * 3.72;
    rect(slide, x, 5.5, 3.5, 0.52, "FFFFFF", LINE);
    slide.addText(t, {
      x: x + 0.18, y: 5.5, w: 3.16, h: 0.52,
      fontFace: FONT, fontSize: 11.5, color: INK2, margin: 0, valign: "middle",
    });
  });

  slide.addText("standard_check", {
    x: 1.1, y: 6.74, w: 6, h: 0.42,
    fontFace: FONT, fontSize: 11, color: MUTED, margin: 0, valign: "middle",
  });
}

// ==========================================================================
// 02 背景
// ==========================================================================
{
  const slide = base("01 / 背景", "查规范这件事，现在是怎么做的", null);

  const items = [
    ["文档多，找不着", "规范动辄几十上百页。你知道有这条规定，但翻不到它在哪一节。"],
    ["口径不一致", "同一条规范，不同人理解不同，落地的做法也不同，出了问题说不清该听谁的。"],
    ["新人上手慢", "遇到问题只能问老员工。人不在就卡住，经验留在个人脑子里，传不下去。"],
    ["记的是旧版本", "文档更新了，脑子里还是老规定。按记忆做事，等于按过期的规范做事。"],
  ];
  const cw = (W - 0.36) / 2;
  items.forEach(([t, d], i) => {
    const x = M + (i % 2) * (cw + 0.36);
    const y = BODY_Y + 0.16 + Math.floor(i / 2) * 2.28;
    card(slide, x, y, cw, 2.0);
    rect(slide, x, y, 0.06, 2.0, i % 2 === 0 ? ACCENT : ACCENT2);
    slide.addText(t, {
      x: x + 0.36, y: y + 0.3, w: cw - 0.72, h: 0.44,
      fontFace: FONT, fontSize: 20, color: INK, bold: true, margin: 0, valign: "middle",
    });
    para(slide, x + 0.36, y + 0.92, cw - 0.72, 0.86, d, { size: 14 });
  });

  footer(slide, 2);
}

// ==========================================================================
// 03 是什么
// ==========================================================================
{
  const slide = base("02 / 是什么", "一句话说清", null);

  card(slide, M, BODY_Y + 0.1, W, 1.36, PLANE);
  rect(slide, M, BODY_Y + 0.1, 0.06, 1.36, ACCENT);
  slide.addText(
    [
      { text: "把规范文档交给它，用自然语言提问 —— 它从文档里找依据、给答案，", options: { color: INK2 } },
      { text: "并告诉你答案来自哪一段原文。", options: { color: INK, bold: true } },
    ],
    {
      x: M + 0.36, y: BODY_Y + 0.1, w: W - 0.72, h: 1.36,
      fontFace: FONT, fontSize: 19, margin: 0, valign: "middle", lineSpacingMultiple: 1.42,
    },
  );

  band(slide, M, 3.62, W, "它是怎么工作的（不需要懂技术）");

  const steps = [
    ["读文档", "把 .docx 读成文字"],
    ["切小块", "按段落切成小片段"],
    ["算向量", "把文字变成可比对的数"],
    ["存起来", "放进本地向量库"],
    ["找依据", "按问题找最相关的几段"],
    ["给答案", "归纳成回答并标注出处"],
  ];
  const sw = 1.83, gap = 0.16;
  steps.forEach(([k, v], i) => {
    const x = M + i * (sw + gap);
    card(slide, x, 4.14, sw, 1.16);
    rect(slide, x, 4.14, sw, 0.05, ACCENT);
    slide.addText(
      [
        { text: k + "\n", options: { bold: true, color: INK, fontSize: 15 } },
        { text: v, options: { color: INK2, fontSize: 10.5 } },
      ],
      { x: x + 0.1, y: 4.24, w: sw - 0.2, h: 0.96,
        fontFace: FONT, margin: 0, valign: "middle", align: "center", lineSpacingMultiple: 1.24, fit: "shrink" },
    );
    if (i < steps.length - 1) {
      slide.addShape(pptx.ShapeType.line, {
        x: x + sw + 0.02, y: 4.72, w: gap - 0.04, h: 0,
        line: { color: BASE, width: 1.5, endArrowType: "triangle" },
      });
    }
  });

  card(slide, M, 5.66, W, 1.1, PLANE);
  rect(slide, M, 5.66, 0.06, 1.1, ACCENT2);
  slide.addText(
    [
      { text: "关键在最后一步：", options: { bold: true, color: INK } },
      {
        text: "答案里的 [1] [2] 是可点击的编号，点一下就跳到对应的原文片段，并显示相似度和内容摘要 —— " +
          "你可以自己核对这个结论是不是真的来自文档。",
        options: { color: INK2 },
      },
    ],
    {
      x: M + 0.36, y: 5.66, w: W - 0.72, h: 1.1,
      fontFace: FONT, fontSize: 14.5, margin: 0, valign: "middle", lineSpacingMultiple: 1.38,
    },
  );

  footer(slide, 3);
}

// ==========================================================================
// 04 三个特点
// ==========================================================================
{
  const slide = base("03 / 特点", "三个决定它能不能用的特点", null);

  const cols = [
    ["可追溯", "答案必须带引用编号。点击就能跳回原文，来源文件、相似度、内容摘要都摆在旁边。" +
      "「它说得对不对」这个问题，从无从判断变成可以核对。"],
    ["本地可跑", "读文档、切分、向量化、向量检索全部在本机完成，不依赖任何在线向量服务。" +
      "文档资料在本机处理 —— 这一点对私有文档是硬要求。"],
    ["换模型不改代码", "DeepSeek、通义千问、智谱 GLM、Ollama、Claude 任选。在界面填模型名和 Key 就切换，" +
      "保存即生效，不用重启也不用改代码。"],
  ];
  const cw = (W - 0.6) / 3;
  cols.forEach(([t, d], i) => {
    const x = M + i * (cw + 0.3);
    card(slide, x, BODY_Y + 0.18, cw, 3.9);
    rect(slide, x, BODY_Y + 0.18, cw, 0.07, ACCENT);
    slide.addText(String(i + 1).padStart(2, "0"), {
      x: x + 0.34, y: BODY_Y + 0.52, w: 1.2, h: 0.5,
      fontFace: FONT, fontSize: 22, color: ACCENT, bold: true, margin: 0, valign: "middle",
    });
    slide.addText(t, {
      x: x + 0.34, y: BODY_Y + 1.12, w: cw - 0.68, h: 0.5,
      fontFace: FONT, fontSize: 23, color: INK, bold: true, margin: 0, valign: "middle",
    });
    para(slide, x + 0.34, BODY_Y + 1.82, cw - 0.68, 2.0, d, { size: 14 });
  });

  footer(slide, 4);
}

// ==========================================================================
// 05 意义
// ==========================================================================
{
  const slide = base("04 / 意义", "它把什么变简单了", null);

  const rows = [
    [th("问题"), th("以前"), th("现在")],
    ["查一条规范", "翻几十页文档，还不一定找得到", "问一句话，几秒钟给答案"],
    ["答案的依据", "靠个人记忆或经验，说不清出处", "明确来自文档，带引用编号"],
    ["能不能核对", "不知道对不对，只能信人", "点引用跳原文，自己判断"],
    ["新人上手", "遇到问题只能问老员工", "自己查，不用等人"],
    ["规范更新了", "没人知道，照着老规定做", "文档一更新，索引自动重建"],
    ["数据安全", "交给外部服务有外泄风险", "全程本机处理"],
  ];
  table(slide, rows, M, 1.86, W, [2.5, 4.7, 4.633], { rowH: 0.54, size: 13.5 });

  card(slide, M, 5.78, W, 0.98, PLANE);
  rect(slide, M, 5.78, 0.06, 0.98, ACCENT);
  slide.addText(
    [
      { text: "更根本的变化：", options: { bold: true, color: INK } },
      {
        text: "「按规范做事」不再依赖某个人记不记得住。知识从个人脑子里搬进了可检索、可核对的系统里。",
        options: { color: INK2 },
      },
    ],
    {
      x: M + 0.36, y: 5.78, w: W - 0.72, h: 0.98,
      fontFace: FONT, fontSize: 15, margin: 0, valign: "middle", lineSpacingMultiple: 1.36,
    },
  );

  footer(slide, 5);
}

// ==========================================================================
// 06 如何使用：跑起来
// ==========================================================================
{
  const slide = base("05 / 如何使用", "三步跑起来", null);

  band(slide, M, BODY_Y, 7.2, "安装与启动");

  const cmds = [
    ["npm install && npm run install:all", "安装依赖"],
    ["npm run ingest", "建立索引（首次约 15 秒）"],
    ["npm run dev", "同时启动后端与前端"],
  ];
  cmds.forEach(([cmd, note], i) => {
    const y = BODY_Y + 0.5 + i * 1.02;
    card(slide, M, y, 7.2, 0.86, PLANE);
    rect(slide, M, y, 0.06, 0.86, ACCENT);
    slide.addText(cmd, {
      x: M + 0.3, y: y + 0.12, w: 5.0, h: 0.34,
      fontFace: MONO, fontSize: 14, color: INK, bold: true, margin: 0, valign: "middle",
    });
    slide.addText("# " + note, {
      x: M + 0.3, y: y + 0.46, w: 6.5, h: 0.3,
      fontFace: MONO, fontSize: 11, color: MUTED, margin: 0, valign: "middle",
    });
    slide.addText(String(i + 1), {
      x: M + 6.5, y: y + 0.12, w: 0.5, h: 0.62,
      fontFace: FONT, fontSize: 24, color: ACCENT, bold: true, align: "right", margin: 0, valign: "middle",
    });
  });

  card(slide, M, BODY_Y + 3.62, 7.2, 1.06, PLANE);
  rect(slide, M, BODY_Y + 3.62, 0.06, 1.06, ACCENT2);
  slide.addText(
    [
      { text: "打开 ", options: { color: INK2 } },
      { text: "http://localhost:5173", options: { color: ACCENT, bold: true } },
      { text: " 就能提问。后端单独跑在 3001 端口。", options: { color: INK2 } },
    ],
    {
      x: M + 0.36, y: BODY_Y + 3.62, w: 6.5, h: 1.06,
      fontFace: FONT, fontSize: 15, margin: 0, valign: "middle",
    },
  );

  band(slide, M + 7.7, BODY_Y, W - 7.7, "放文档进去就行", ACCENT2);
  const notes = [
    ["放到 knowledgeFile/", "支持 .docx。放进去不用做任何操作。"],
    ["索引自动重建", "检测到文件增删改，或参数变化，会自动重新建立索引。"],
    ["上传也行", "在界面上传文件，同样自动重建，不用重启。"],
  ];
  notes.forEach(([t, d], i) => {
    const y = BODY_Y + 0.5 + i * 1.42;
    card(slide, M + 7.7, y, W - 7.7, 1.26);
    slide.addText(t, {
      x: M + 7.94, y: y + 0.16, w: W - 8.18, h: 0.36,
      fontFace: FONT, fontSize: 15, color: INK, bold: true, margin: 0, valign: "middle",
    });
    para(slide, M + 7.94, y + 0.58, W - 8.18, 0.58, d, { size: 12.5 });
  });

  footer(slide, 6);
}

// ==========================================================================
// 07 如何使用：日常
// ==========================================================================
{
  const slide = base("05 / 如何使用", "日常怎么用", "不需要懂技术，界面上四件事");

  const items = [
    ["提问", "直接打字问，支持多轮对话。", ACCENT],
    ["换模型", "侧边栏填模型名和 API Key，保存立刻生效。", ACCENT2],
    ["加文档", "拖文件上传，成功后自动重建索引。", ACCENT2],
    ["核对", "点答案里的 [1]，跳到原文那一段。", ACCENT],
  ];
  const cw = (W - 0.36) / 2;
  items.forEach(([t, d, c], i) => {
    const x = M + (i % 2) * (cw + 0.36);
    const y = BODY_Y + 0.2 + Math.floor(i / 2) * 2.3;
    card(slide, x, y, cw, 2.04);
    rect(slide, x, y, 0.06, 2.04, c);
    slide.addText(String(i + 1), {
      x: x + 0.36, y: y + 0.32, w: 0.62, h: 0.62,
      fontFace: FONT, fontSize: 30, color: c, bold: true, margin: 0, valign: "middle",
    });
    slide.addText(t, {
      x: x + 1.12, y: y + 0.32, w: cw - 1.5, h: 0.62,
      fontFace: FONT, fontSize: 22, color: INK, bold: true, margin: 0, valign: "middle",
    });
    para(slide, x + 0.36, y + 1.14, cw - 0.72, 0.7, d, { size: 15 });
  });

  footer(slide, 7);
}

// ==========================================================================
// 08 效果
// ==========================================================================
{
  const slide = base("06 / 效果", "实测数据", "基于 4 份文档、398 个文本块");

  const tiles = [
    ["top-1 命中", "8 / 9", "第一条就是正确文档"],
    ["top-8 召回", "9 / 9", "前 8 条里一定有正确文档"],
    ["MRR@8", "0.917", "排名越靠前分越高"],
    ["切题率", "5 / 5", "人工标注，看内容切不切题"],
  ];
  const tw = (W - 0.84) / 4;
  tiles.forEach(([label, value, sub], i) => {
    const x = M + i * (tw + 0.28);
    card(slide, x, BODY_Y + 0.2, tw, 1.92);
    rect(slide, x, BODY_Y + 0.2, tw, 0.06, ACCENT);
    slide.addText(label, {
      x: x + 0.28, y: BODY_Y + 0.44, w: tw - 0.56, h: 0.32,
      fontFace: FONT, fontSize: 12.5, color: MUTED, margin: 0, valign: "middle",
    });
    slide.addText(value, {
      x: x + 0.28, y: BODY_Y + 0.84, w: tw - 0.56, h: 0.7,
      fontFace: FONT, fontSize: 36, color: INK, bold: true, margin: 0, valign: "middle",
    });
    slide.addText(sub, {
      x: x + 0.28, y: BODY_Y + 1.58, w: tw - 0.56, h: 0.36,
      fontFace: FONT, fontSize: 10.5, color: INK2, margin: 0, valign: "middle",
    });
  });

  band(slide, M, 4.22, W, "但这些数字要怎么看");

  card(slide, M, 4.66, W, 2.16, PLANE);
  rect(slide, M, 4.66, 0.06, 2.16, ACCENT2);
  slide.addText(
    [
      { text: "评估集很小 —— 只有 9 条问答对和 5 条人工标注。\n\n", options: { bold: true, color: INK, fontSize: 15 } },
      {
        text: "所以正确读法是「在这些样本上没发现问题」，而不是「准确率 89%」—— 样本量不足以支撑后者。\n" +
          "目前唯一站得住的结论是链路跑通了、检索能找对文档。要说「检索准确」，评估集还需要扩到几十条以上。",
        options: { color: INK2, fontSize: 13.5 } },
    ],
    {
      x: M + 0.36, y: 4.8, w: W - 0.72, h: 1.88,
      fontFace: FONT, margin: 0, valign: "top", lineSpacingMultiple: 1.38,
    },
  );

  footer(slide, 8);
}

// ==========================================================================
// 09 存在的问题：可信度
// ==========================================================================
{
  const slide = base("07 / 存在的问题", "最要紧的一条：答案可能不是来自文档", null);

  card(slide, M, BODY_Y + 0.12, W, 1.34, PLANE);
  rect(slide, M, BODY_Y + 0.12, 0.06, 1.34, ACCENT2);
  slide.addText(
    [
      { text: "实测中发现过：", options: { bold: true, color: INK } },
      {
        text: "问「Java 开发规范对异常处理有什么要求」时，真正回答问题的那一节没有被检索到。" +
          "模型照样给出了很像样的答案 —— 因为这份手册是公开资料，它很可能是凭自己的记忆答的。",
        options: { color: INK2 },
      },
    ],
    {
      x: M + 0.36, y: BODY_Y + 0.12, w: W - 0.72, h: 1.34,
      fontFace: FONT, fontSize: 15, margin: 0, valign: "middle", lineSpacingMultiple: 1.38,
    },
  );

  const pts = [
    ["答案看起来对，但和知识库无关",
      "换一份私有文档，同样的链路就会给出一本正经的错误答案 —— 而且不会有任何提示。"],
    ["引用编号不能当担保",
      "模型标了 [1][2]，不代表那两条真能支撑结论。目前没有机制校验「这句话是不是出自那段原文」。"],
    ["现有防线只有人工核对",
      "界面上能看到来源卡片和原文，但这依赖人愿意去点、去读。人不看就没有防线。"],
  ];
  pts.forEach(([t, d], i) => {
    const y = 3.5 + i * 1.28;
    point(slide, M, y, W, String(i + 1), t, d, ACCENT2);
  });

  footer(slide, 9);
}

// ==========================================================================
// 10 存在的问题：功能与性能
// ==========================================================================
{
  const slide = base("07 / 存在的问题", "功能与性能上的短板", null);

  const rows = [
    [th("问题"), th("现状"), th("影响")],
    ["只支持 .docx", "PDF、PPT、Excel 都读不了", "很多规范是 PDF 或 PPT，进不来"],
    ["不是流式输出", "整段答案一次性返回", "长回答要干等十几秒才看到内容"],
    ["多轮追问弱", "每轮只用新问题检索，历史不重新检索", "深度追问会退化，答非所问"],
    ["没有重排序", "召回 top-8 直接用", "文档变多后准确率会下降"],
    ["重建索引全量重算", "398 块约 15 秒，文档多了线性变慢", "加一份文档也要全部重算"],
    ["界面上不能删文档", "只能加不能减", "规范过期后清理不掉，会被检索到"],
    ["单机单人", "没有权限、审计、并发设计", "不能直接当成团队服务上线"],
  ];
  table(slide, rows, M, 1.72, W, [2.85, 4.6, 4.383], { rowH: 0.53, size: 12.5 });

  card(slide, M, 6.08, W, 0.9, PLANE);
  rect(slide, M, 6.08, 0.06, 0.9, ACCENT2);
  slide.addText(
    [
      { text: "一句话概括：", options: { bold: true, color: INK } },
      {
        text: "这是一个把流程跑通、并做了检索质量调优的可用原型 —— 还不是一个可以直接投入生产环境的产品。",
        options: { color: INK2 },
      },
    ],
    {
      x: M + 0.36, y: 6.08, w: W - 0.72, h: 0.9,
      fontFace: FONT, fontSize: 15, margin: 0, valign: "middle",
    },
  );

  footer(slide, 10);
}

// ==========================================================================
// 11 后续 + 结语
// ==========================================================================
{
  const slide = base("08 / 后续", "接下来先做什么", null);

  const rows = [
    [th("优先级"), th("要做的事"), th("为什么先做它")],
    ["1", "忠实度校验", "解决「答案可能不是来自文档」—— 这是唯一会让人被误导的问题"],
    ["2", "流式输出", "体验改善最明显，改动小、收益大"],
    ["3", "混合检索 + 重排序", "在文档变多之前做好，否则准确率会掉"],
    ["4", "多格式支持 + 文件管理", "把 PDF、PPT 收进来，过期的能删掉"],
    ["5", "多用户与权限", "从个人工具变成团队服务"],
  ];
  table(slide, rows, M, BODY_Y + 0.12, W, [1.35, 3.6, 6.883], { rowH: 0.6, size: 13.5 });

  card(slide, M, 5.72, W, 1.14, PLANE);
  rect(slide, M, 5.72, 0.06, 1.14, ACCENT);
  slide.addText(
    [
      { text: "结语\n", options: { bold: true, color: INK, fontSize: 17 } },
      {
        text: "链路已经跑通：文档进去，问一句话，答案带着出处出来。" +
          "真正花时间的不是把流程接起来，而是让「检索到的确实是能回答问题的那一段」，以及「答案确实是那段原文撑起来的」—— " +
          "这两件事只能靠可量化的指标和人工标注的用例来守住。",
        options: { color: INK2, fontSize: 14 } },
    ],
    {
      x: M + 0.36, y: 5.82, w: W - 0.72, h: 0.94,
      fontFace: FONT, margin: 0, valign: "middle", lineSpacingMultiple: 1.34,
    },
  );

  footer(slide, 11);
}

await pptx.writeFile({ fileName: OUT });
console.log("written:", OUT);
