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

## 设计说明

### 串行压缩，优先保证正确性

本工具采用**串行**处理队列中的文件，而非并行压缩。原因如下：

1. **WASM 引擎为单例**：Ghostscript 和 QPDF 的 WASM 模块内部使用全局内存文件系统（Emscripten FS），无法同时运行多个实例。并行调用会导致文件路径冲突或状态污染。
2. **控制内存峰值**：PDF 压缩过程中，原始文件、Ghostscript 输出、QPDF 输出可能同时驻留内存。串行执行可将内存峰值限制在单文件级别，避免在浏览器中触发 OOM。
3. **状态可预测**：串行处理让进度显示、错误回滚和日志收集更简单可靠，不需要额外的锁或隔离机制。

因此，即使加入多个文件，它们也会**逐个完成**，每个文件处理结束后才启动下一个。已完成的文件可以立即下载，不需要等待队列全部结束。

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
