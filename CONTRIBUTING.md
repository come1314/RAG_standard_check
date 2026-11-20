# 贡献指南

欢迎提交 Issue 和 Pull Request。

## 开发环境

- Node.js 20+（推荐 LTS）
- 首次运行：`npm install && npm run install:all`
- 准备文档：把 `.docx` 放进 `knowledgeFile/`（该目录默认不入库）
- 构建索引：`npm run ingest`
- 启动：`npm run dev`，打开 http://localhost:5173

不配置任何 API Key 也能跑通检索流程；回答生成需要配置模型凭证，见 README 的「配置」一节。

## 提交前检查

```bash
npm run typecheck
npm run build
```

`npm run eval` 依赖你自己索引的文档，结果随语料变化，不强求与 README 中的数值一致。

## 注意事项

- 不要把真实文档、`.env`、`server/data/` 或任何密钥提交到仓库。
- 只提交通用样例文档；`knowledgeFile/` 与 `server/data/` 已被 `.gitignore` 忽略。
- 改动分块 / 检索参数后，请重新 `npm run ingest` 再评估效果。

## 许可证

提交即表示你同意贡献内容按仓库的 MIT License 授权。
