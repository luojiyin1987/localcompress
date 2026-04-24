# Paper Squeeze

浏览器内本地压缩 PDF、PPTX 和 DOCX 的静态站点。文件全程不上传，所有处理在 Web Worker 中完成。

[![License: AGPL v3](https://img.shields.io/badge/License-AGPL%20v3-blue.svg)](https://www.gnu.org/licenses/agpl-3.0)

## 功能

| 格式 | 压缩方式 |
|---|---|
| PDF | Ghostscript WASM 重压缩，可选 QPDF WASM 结构优化 |
| PPTX | JSZip 解包 + jSquash 图片重压缩（OxiPNG / MozJPEG）后回写 |
| DOCX | JSZip 解包 + jSquash 图片重压缩（OxiPNG / MozJPEG）后回写 |

- **批量队列**：一次拖入多个文件，按顺序处理
- **三档压缩**：均衡 / 强力 / 原图，不同档位对应不同的图片编码策略
- **Worker 隔离**：压缩任务在 Web Worker 中运行，不阻塞主线程
- **File System Access API**：支持直接保存到本地目录（浏览器支持时）

## 技术栈

- [Vite](https://vitejs.dev/) — 构建工具
- [Cloudflare Pages](https://pages.cloudflare.com/) — 静态托管
- [@okathira/ghostpdl-wasm](https://www.npmjs.com/package/@okathira/ghostpdl-wasm) — Ghostscript WASM
- [@neslinesli93/qpdf-wasm](https://www.npmjs.com/package/@neslinesli93/qpdf-wasm) — QPDF WASM
- [@jsquash/jpeg](https://www.npmjs.com/package/@jsquash/jpeg) — MozJPEG WASM 编码
- [@jsquash/oxipng](https://www.npmjs.com/package/@jsquash/oxipng) — OxiPNG WASM 无损优化
- [JSZip](https://stuk.github.io/jszip/) — Office 文档解包

## 快速开始

```bash
npm install
npm run dev
```

打开 `http://localhost:5173`。

## 部署

项目已配置为 Cloudflare Pages 静态站点。一键部署：

```bash
npm run deploy
```

或使用 `wrangler pages deploy` 手动部署 `dist` 目录。

## 项目结构

```
.
├── app.js                  # 主线程 UI 逻辑
├── workers/pdf-worker.js   # Web Worker：WASM 压缩引擎
├── index.html              # 入口页面
├── styles.css              # 样式
├── package.json
└── wrangler.toml           # Cloudflare Pages 配置
```

## 开发

| 命令 | 说明 |
|---|---|
| `npm run dev` | 本地开发服务器 |
| `npm run build` | 生产构建 |
| `npm run preview` | 用 Wrangler 本地预览 |
| `npm run check` | 语法检查 + 构建验证 |
| `npm run deploy` | 构建并部署到 Cloudflare Pages |

## 隐私

这是一个纯静态站点。文件在浏览器中通过 Web Worker 和 WASM 处理，**不会上传到任何服务器**。

## License

[AGPL-3.0-or-later](LICENSE)
