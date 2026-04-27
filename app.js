const compressionTimeoutMs = 300000;

let worker;
let activeCompression = null;

const fileKinds = {
  pdf: {
    mime: "application/pdf",
    extension: ".pdf",
    pickerDescription: "PDF Document",
    compressType: "compress-pdf",
    requiresEngine: true,
    isOffice: false,
  },
  pptx: {
    mime: "application/vnd.openxmlformats-officedocument.presentationml.presentation",
    extension: ".pptx",
    pickerDescription: "PowerPoint Presentation",
    compressType: "compress-office",
    requiresEngine: false,
    isOffice: true,
  },
  docx: {
    mime: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
    extension: ".docx",
    pickerDescription: "Word Document",
    compressType: "compress-office",
    requiresEngine: false,
    isOffice: true,
  },
};

const state = {
  files: [],
  isProcessing: false,
  engineReady: false,
  engineMessage: "压缩引擎加载中",
};

const dropzone = document.querySelector("#dropzone");
const fileInput = document.querySelector("#file-input");
const fileList = document.querySelector("#file-list");
const queueSummary = document.querySelector("#queue-summary");
const profileSelect = document.querySelector("#profile-select");
const structureToggle = document.querySelector("#structure-toggle");
const compressButton = document.querySelector("#compress-button");
const engineStatus = document.querySelector("#engine-status");
const rowTemplate = document.querySelector("#file-row-template");

const formatBytes = (bytes) => {
  if (!Number.isFinite(bytes) || bytes <= 0) {
    return "0 B";
  }

  const units = ["B", "KB", "MB", "GB"];
  const exponent = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
  const value = bytes / 1024 ** exponent;
  return `${value.toFixed(value >= 10 || exponent === 0 ? 0 : 1)} ${units[exponent]}`;
};

const setEngineStatus = (label, tone = "") => {
  engineStatus.textContent = label;
  engineStatus.className = `status-chip${tone ? ` ${tone}` : ""}`;
};

const setStructureToggleEnabled = (enabled) => {
  structureToggle.disabled = !enabled;
  if (!enabled) {
    structureToggle.checked = false;
  }
  structureToggle.closest(".checkbox-field")?.classList.toggle("disabled-field", !enabled);
};

const getKindConfig = (kind) => fileKinds[kind] ?? null;

const getFileKind = (file) => {
  const normalizedName = file.name.toLowerCase();

  return Object.entries(fileKinds).find(([, config]) => {
    return file.type === config.mime || normalizedName.endsWith(config.extension);
  })?.[0] ?? "";
};

const isOfficeKind = (kind) => Boolean(getKindConfig(kind)?.isOffice);

const canProcessItem = (item) => {
  const config = getKindConfig(item.kind);
  if (!config) {
    return false;
  }

  return config.requiresEngine ? state.engineReady : true;
};

const createItem = (file) => ({
  id: crypto.randomUUID(),
  file,
  kind: getFileKind(file),
  statusLabel: "待处理",
  tone: "",
  message: "已加入队列，等待开始。",
  outputBlob: null,
  outputName: "",
  resultBytes: 0,
});

const getItemById = (id) => state.files.find((item) => item.id === id);

const clearOutput = (item) => {
  item.outputBlob = null;
  item.outputName = "";
  item.resultBytes = 0;
};

const createClearedResultPatch = () => ({
  resultBytes: 0,
  outputName: "",
  outputBlob: null,
});

const createFailedItemPatch = (message) => ({
  statusLabel: "未完成",
  tone: "error",
  message,
  ...createClearedResultPatch(),
});

const getPickerTypes = (item) => {
  const config = getKindConfig(item.kind) ?? fileKinds.pdf;
  return [
    {
      description: config.pickerDescription,
      accept: {
        [config.mime]: [config.extension],
      },
    },
  ];
};

const fallbackDownload = (blob, filename) => {
  const downloadUrl = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = downloadUrl;
  anchor.download = filename;
  anchor.rel = "noopener";
  anchor.style.display = "none";
  document.body.append(anchor);
  anchor.click();
  anchor.remove();
  setTimeout(() => URL.revokeObjectURL(downloadUrl), 1000);
};

const saveBlob = async (item) => {
  if (!("showSaveFilePicker" in window)) {
    fallbackDownload(item.outputBlob, item.outputName);
    return;
  }

  const handle = await window.showSaveFilePicker({
    suggestedName: item.outputName,
    types: getPickerTypes(item),
  });
  const writable = await handle.createWritable();
  await writable.write(item.outputBlob);
  await writable.close();
};

const updateSummary = () => {
  const totalBytes = state.files.reduce((sum, item) => sum + item.file.size, 0);
  queueSummary.innerHTML = `
    <span>${state.files.length} 个文件</span>
    <span>总计 ${formatBytes(totalBytes)}</span>
  `;
};

const renderEmptyState = () => {
  const emptyRow = document.createElement("li");
  emptyRow.className = "empty-row";
  emptyRow.textContent = "队列为空。当前接受 PDF、PPTX 和 DOCX 文件，支持一次加入多个文件。";
  fileList.appendChild(emptyRow);
};

const renderItem = (item) => {
  const fragment = rowTemplate.content.cloneNode(true);
  const row = fragment.querySelector(".file-row");
  const name = fragment.querySelector(".file-name");
  const size = fragment.querySelector(".file-size");
  const badge = fragment.querySelector(".file-badge");
  const message = fragment.querySelector(".file-message");
  const download = fragment.querySelector(".download-link");

  name.textContent = item.file.name;
  size.textContent = item.resultBytes
    ? `${formatBytes(item.file.size)} -> ${formatBytes(item.resultBytes)}`
    : formatBytes(item.file.size);
  badge.textContent = item.statusLabel;
  badge.className = `file-badge ${item.tone}`;
  message.textContent = item.message;

  if (item.tone === "success" && item.outputBlob) {
    download.hidden = false;
    download.dataset.fileId = item.id;
    download.textContent = "下载文件";
  }

  row.dataset.fileId = item.id;
  fileList.appendChild(fragment);
};

const renderFiles = () => {
  fileList.textContent = "";

  if (state.files.length === 0) {
    renderEmptyState();
  }

  state.files.forEach(renderItem);
  updateSummary();
  compressButton.disabled =
    state.isProcessing || state.files.length === 0 || !state.files.some((item) => canProcessItem(item));
};

const markItem = (id, patch) => {
  const item = getItemById(id);
  if (!item) {
    return;
  }

  Object.assign(item, patch);
  renderFiles();
};

const addFiles = (incomingFiles) => {
  incomingFiles
    .filter((file) => getFileKind(file))
    .forEach((file) => {
      state.files.push(createItem(file));
    });

  renderFiles();
};

const handleSelection = (fileListLike) => {
  addFiles(Array.from(fileListLike));
};

const createWorker = () => {
  const nextWorker = new Worker(new URL("./workers/pdf-worker.js", import.meta.url), { type: "module" });
  nextWorker.addEventListener("message", handleWorkerMessage);
  nextWorker.addEventListener("error", handleWorkerError);
  nextWorker.addEventListener("messageerror", handleWorkerMessageError);
  return nextWorker;
};

const probeWorker = () => {
  worker.postMessage({ type: "probe-engine" });
};

const restartWorker = (message) => {
  worker?.terminate();
  worker = createWorker();
  state.engineReady = false;
  state.engineMessage = message;
  setEngineStatus(message);
  setStructureToggleEnabled(false);
  renderFiles();
  probeWorker();
};

const finishActiveCompression = (result) => {
  activeCompression?.resolve(result);
};

const compressFile = (item, profile, optimizeStructure) =>
  new Promise((resolve) => {
    if (activeCompression) {
      resolve({
        ok: false,
        id: item.id,
        error: "已有压缩任务正在运行，请稍后重试。",
      });
      return;
    }

    const channel = new MessageChannel();
    let resolved = false;
    let timeoutId;

    const cleanup = () => {
      clearTimeout(timeoutId);
      channel.port1.close();
    };

    const doResolve = (value) => {
      if (resolved) {
        return;
      }

      resolved = true;
      if (activeCompression?.id === item.id) {
        activeCompression = null;
      }
      cleanup();
      resolve(value);
    };

    activeCompression = {
      id: item.id,
      resolve: doResolve,
    };

    timeoutId = setTimeout(() => {
      doResolve({
        ok: false,
        id: item.id,
        error: "压缩任务超时，请重试或降低档位。",
        stopQueue: true,
      });
      restartWorker("压缩任务超时，压缩引擎已重启。");
    }, compressionTimeoutMs);

    channel.port1.onmessage = (event) => {
      doResolve(event.data);
    };

    channel.port1.onmessageerror = () => {
      doResolve({
        ok: false,
        id: item.id,
        error: "Worker 消息通道错误，数据无法解析。",
      });
    };

    item.file
      .arrayBuffer()
      .then((buffer) => {
        if (resolved) {
          return;
        }

        try {
          worker.postMessage(
            {
              type: getKindConfig(item.kind)?.compressType ?? "compress-pdf",
              id: item.id,
              kind: item.kind,
              name: item.file.name,
              buffer,
              profile,
              optimizeStructure,
            },
            [buffer, channel.port2]
          );
        } catch (error) {
          doResolve({
            ok: false,
            id: item.id,
            error: error instanceof Error ? error.message : "Worker 通信失败。",
          });
        }
      })
      .catch((error) => {
        doResolve({
          ok: false,
          id: item.id,
          error: error instanceof Error ? error.message : "读取文件失败。",
        });
      });
  });

const buildProcessingMessage = (item, profile, optimizeStructure, currentIndex, totalCount) => {
  if (isOfficeKind(item.kind)) {
    return `JSZip 正在解包 ${item.kind.toUpperCase()}（${currentIndex}/${totalCount}），并按 ${profile} 档位重压缩图片。`;
  }

  if (optimizeStructure) {
    return `Ghostscript 压缩后将执行 QPDF 结构优化（${currentIndex}/${totalCount}），使用 ${profile} 档位。`;
  }

  return `Ghostscript WASM 正在处理（${currentIndex}/${totalCount}），使用 ${profile} 档位。`;
};

const markItemPreparing = (item, profile, optimizeStructure, currentIndex, totalCount) => {
  markItem(item.id, {
    statusLabel: "压缩中",
    tone: "processing",
    message: buildProcessingMessage(item, profile, optimizeStructure, currentIndex, totalCount),
    ...createClearedResultPatch(),
  });
};

const isBinaryOutput = (value) => {
  return value instanceof ArrayBuffer || ArrayBuffer.isView(value);
};

const markItemSuccess = (item, result) => {
  if (!isBinaryOutput(result.buffer)) {
    markItem(item.id, createFailedItemPatch("压缩结果无效，未收到可保存的输出文件。"));
    return;
  }

  const blob = new Blob([result.buffer], {
    type: getKindConfig(item.kind)?.mime ?? fileKinds.pdf.mime,
  });

  markItem(item.id, {
    statusLabel: "已完成",
    tone: "success",
    message: result.message,
    resultBytes: blob.size,
    outputName: result.outputName,
    outputBlob: blob,
  });
};

const processQueue = async ({ profile, optimizeStructure }) => {
  const queue = state.files.filter((item) => item.kind);
  const totalCount = queue.length;
  let currentIndex = 0;

  // Ghostscript/QPDF 的 WASM 文件系统是单例，只能串行处理。
  for (const item of queue) {
    currentIndex += 1;
    clearOutput(item);

    if (item.kind === "pdf" && !state.engineReady) {
      markItem(item.id, createFailedItemPatch(state.engineMessage || "PDF 压缩引擎未就绪。"));
      continue;
    }

    const shouldOptimizeStructure = item.kind === "pdf" && optimizeStructure;
    markItemPreparing(item, profile, shouldOptimizeStructure, currentIndex, totalCount);

    const result = await compressFile(item, profile, shouldOptimizeStructure);
    if (!result.ok) {
      markItem(item.id, createFailedItemPatch(result.error));
      if (result.stopQueue) {
        break;
      }
      continue;
    }

    markItemSuccess(item, result);
  }
};

const handleCompressClick = async () => {
  state.isProcessing = true;
  renderFiles();

  try {
    await processQueue({
      profile: profileSelect.value,
      optimizeStructure: structureToggle.checked,
    });
  } finally {
    state.isProcessing = false;
    renderFiles();
  }
};

const handleDownloadClick = async (event) => {
  const link = event.target.closest(".download-link");
  if (!link) {
    return;
  }

  const item = getItemById(link.dataset.fileId);
  if (!item?.outputBlob || item.tone !== "success") {
    return;
  }

  const previousMessage = item.message;
  markItem(item.id, {
    message: "正在写入下载文件...",
  });

  try {
    await saveBlob(item);
    markItem(item.id, {
      message: "文件已保存到本地。",
    });
  } catch (error) {
    if (error instanceof DOMException && error.name === "AbortError") {
      markItem(item.id, {
        message: previousMessage,
      });
      return;
    }

    markItem(item.id, {
      message: `保存失败：${error instanceof Error ? error.message : "浏览器拒绝写入文件。"}`,
    });
  }
};

const handleDropzoneKeyboard = (event) => {
  if (event.key !== "Enter" && event.key !== " ") {
    return;
  }

  event.preventDefault();
  fileInput.click();
};

const bindDropzoneEvents = () => {
  dropzone.addEventListener("click", () => fileInput.click());
  dropzone.addEventListener("keydown", handleDropzoneKeyboard);

  ["dragenter", "dragover"].forEach((eventName) => {
    dropzone.addEventListener(eventName, (event) => {
      event.preventDefault();
      dropzone.classList.add("dragover");
    });
  });

  ["dragleave", "drop"].forEach((eventName) => {
    dropzone.addEventListener(eventName, (event) => {
      event.preventDefault();
      dropzone.classList.remove("dragover");
    });
  });

  dropzone.addEventListener("drop", (event) => {
    handleSelection(event.dataTransfer.files);
  });
};

const bindFileInputEvents = () => {
  fileInput.addEventListener("change", (event) => {
    handleSelection(event.target.files);
    fileInput.value = "";
  });
};

function handleWorkerMessage(event) {
  if (event.data?.type !== "engine-status") {
    return;
  }

  state.engineReady = event.data.ready;
  state.engineMessage = event.data.message;

  if (event.data.ready) {
    setEngineStatus(event.data.message, "ready");
    setStructureToggleEnabled(event.data.qpdfReady !== false);
  } else {
    setEngineStatus(event.data.message, event.data.loading ? "" : "error");
    setStructureToggleEnabled(false);
  }

  renderFiles();
}

function handleWorkerError(event) {
  console.error("Worker error:", event.message);
  const message = event.message || "Worker 发生错误";
  finishActiveCompression({
    ok: false,
    id: activeCompression?.id,
    error: `${message}，压缩引擎已重启。`,
    stopQueue: true,
  });
  restartWorker(`${message}，压缩引擎已重启。`);
}

function handleWorkerMessageError() {
  finishActiveCompression({
    ok: false,
    id: activeCompression?.id,
    error: "Worker 消息通道错误，压缩引擎已重启。",
    stopQueue: true,
  });
  restartWorker("Worker 消息通道错误，压缩引擎已重启。");
}

const init = () => {
  bindDropzoneEvents();
  bindFileInputEvents();
  compressButton.addEventListener("click", handleCompressClick);
  fileList.addEventListener("click", handleDownloadClick);

  worker = createWorker();
  probeWorker();
  renderFiles();
};

init();
