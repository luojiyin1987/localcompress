# Agent Guide

## Project Overview

Paper Squeeze — browser-based local compression for PDF, PPTX, and DOCX. Pure static site (Cloudflare Pages). The main thread owns UI state, queue orchestration, Worker lifecycle, and local saving; compression work is dispatched to a Web Worker. PDF compression uses Ghostscript/QPDF WASM, while Office compression uses JSZip plus browser image APIs and jSquash WASM encoders. Nothing is uploaded to a server.

## Tech Stack

- Vite (build)
- Vanilla JS (ES modules, no framework)
- Cloudflare Pages + Wrangler
- WASM: Ghostscript, QPDF, MozJPEG, OxiPNG
- JSZip for Office documents

## Commands

| Command | Purpose |
|---|---|
| `npm install` | Install deps |
| `npm run dev` | Vite dev server (`http://localhost:5173`) |
| `npm run build` | Production build to `dist/` |
| `npm run preview` | Preview `dist/` with Wrangler locally |
| `npm run check` | Syntax check (`node --check`) + build verification |
| `npm run deploy` | Build and deploy to Cloudflare Pages |

## Architecture

- `app.js` — main thread UI, queue orchestration, Worker lifecycle/restart handling, and local save flow
- `workers/pdf-worker.js` — **all** compression logic (PDF + Office), despite the filename
- `index.html` — single static entry page
- `wrangler.toml` — Cloudflare Pages config

## Critical Constraints

- **Serial processing only**: Ghostscript and QPDF WASM modules use a singleton in-memory file system (Emscripten FS). Parallel instances cause path conflicts and state corruption. The queue in `app.js` processes files one at a time by design.
- **Worker lifecycle**: `app.js` keeps one active Worker and one active compression task at a time, but it can terminate and recreate the Worker on timeout, `error`, or `messageerror`. Restart paths must preserve predictable UI state and queue behavior.
- **WASM loading**: WASM binaries are imported with Vite's `?url` suffix (e.g., `import ghostscriptWasmUrl from "@okathira/ghostpdl-wasm/gs.wasm?url"`). Do not remove `?url`.
- **No test framework**: The project has no unit tests. Verification is `npm run check` (syntax + build) plus manual browser testing.
- **Modern browser APIs**: Office image recompression feature-detects `OffscreenCanvas` and `createImageBitmap` in `recompressImage()`. If either API is unavailable, Office files are still rebuilt with JSZip, but image recompression is skipped. Downloads prefer the File System Access API (`showSaveFilePicker`) when available and fall back to an anchor download.

## File Type Detection

Detection is in `app.js` `getFileKind()`. It checks both `file.type` and filename extension. Do not rely on MIME type alone — some OS/browsers send generic types for Office files.

## Style / Conventions

- ES modules everywhere (`"type": "module"` in `package.json`)
- Existing JavaScript uses semicolons — follow that style
- Chinese UI strings — keep user-facing text in Chinese
- Worker uses camelCase; `app.js` uses camelCase

## Deployment

- Target: Cloudflare Pages
- Output dir: `dist/` (configured in `wrangler.toml` and `package.json` scripts)
- Project name: `localcompress`

## License

AGPL-3.0-or-later. If modifying, ensure compliance.
