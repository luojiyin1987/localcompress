const compressionTimeoutMs = 300000;

let worker;
let activeCompression = null;

const pptxMime = "application/vnd.openxmlformats-officedocument.presentationml.presentation";
const docxMime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
const officeMimeByKind = {
  pptx: pptxMime,
  docx: docxMime,
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
const profileOptions = Array.from(document.querySelectorAll(".profile-option"));
const structureToggle = document.querySelector("#structure-toggle");
const compressButton = document.querySelector("#compress-button");
const engineStatus = document.querySelector("#engine-status");
const rowTemplate = document.querySelector("#file-row-template");

const setStructureToggleEnabled = (enabled) => {
  structureToggle.disabled = !enabled;
  if (!enabled) {
    structureToggle.checked = false;
  }
  structureToggle.closest(".checkbox-field")?.classList.toggle("disabled-field", !enabled);
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

const syncProfileOptions = (value) => {
  profileOptions.forEach((option) => {
    const isSelected = option.dataset.profile === value;
    option.classList.toggle("is-selected", isSelected);
    option.setAttribute("aria-checked", String(isSelected));
    option.tabIndex = isSelected ? 0 : -1;
  });
};

const selectProfile = (value, shouldDispatch = false) => {
  profileSelect.value = value;
  syncProfileOptions(profileSelect.value);

  if (shouldDispatch) {
    profileSelect.dispatchEvent(new Event("change", { bubbles: true }));
  }
};

const getFileKind = (file) => {
  const name = file.name.toLowerCase();

  if (file.type === "application/pdf" || name.endsWith(".pdf")) {
    return "pdf";
  }

  if (file.type === pptxMime || name.endsWith(".pptx")) {
    return "pptx";
  }

  if (file.type === docxMime || name.endsWith(".docx")) {
    return "docx";
  }

  return "";
};

const isOfficeKind = (kind) => kind === "pptx" || kind === "docx";

const canProcessItem = (item) => item.kind === "pdf" ? state.engineReady : isOfficeKind(item.kind);

const updateSummary = () => {
  const totalBytes = state.files.reduce((sum, item) => sum + item.file.size, 0);
  queueSummary.innerHTML = `
    <span>${state.files.length} 个文件</span>
    <span>总计 ${formatBytes(totalBytes)}</span>
  `;
};

const revokeDownload = (item) => {
  if (item.downloadUrl) {
    URL.revokeObjectURL(item.downloadUrl);
    item.downloadUrl = "";
  }
};

const clearOutput = (item) => {
  revokeDownload(item);
  item.outputBlob = null;
  item.outputName = "";
  item.resultBytes = 0;
};

const getPickerTypes = (item) => {
  if (item.kind === "pptx") {
    return [
      {
        description: "PowerPoint Presentation",
        accept: {
          [pptxMime]: [".pptx"],
        },
      },
    ];
  }

  if (item.kind === "docx") {
    return [
      {
        description: "Word Document",
        accept: {
          [docxMime]: [".docx"],
        },
      },
    ];
  }

  return [
    {
      description: "PDF Document",
      accept: {
        "application/pdf": [".pdf"],
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

  return downloadUrl;
};

const saveBlob = async (item) => {
  revokeDownload(item);

  if (!("showSaveFilePicker" in window)) {
    item.downloadUrl = fallbackDownload(item.outputBlob, item.outputName);
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

const renderFiles = () => {
  fileList.textContent = "";

  if (state.files.length === 0) {
    const emptyRow = document.createElement("li");
    emptyRow.className = "empty-row";
    emptyRow.textContent = "队列为空。当前接受 PDF、PPTX 和 DOCX 文件，支持一次加入多个文件。";
    fileList.appendChild(emptyRow);
  }

  state.files.forEach((item) => {
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
  });

  updateSummary();
  compressButton.disabled =
    state.isProcessing || state.files.length === 0 || !state.files.some((item) => canProcessItem(item));
};

const createItem = (file) => ({
  id: crypto.randomUUID(),
  file,
  kind: getFileKind(file),
  statusLabel: "待处理",
  tone: "",
  message: "已加入队列，等待开始。",
  downloadUrl: "",
  outputBlob: null,
  outputName: "",
  resultBytes: 0,
});

const addFiles = (incomingFiles) => {
  const supportedFiles = incomingFiles.filter((file) => getFileKind(file));

  supportedFiles.forEach((file) => {
    state.files.push(createItem(file));
  });

  renderFiles();
};

const markItem = (id, patch) => {
  const item = state.files.find((entry) => entry.id === id);
  if (!item) {
    return;
  }

  Object.assign(item, patch);
  renderFiles();
};

const handleSelection = (fileListLike) => {
  addFiles(Array.from(fileListLike));
};

profileOptions.forEach((option) => {
  option.addEventListener("click", () => {
    selectProfile(option.dataset.profile, true);
  });

  option.addEventListener("keydown", (event) => {
    const currentIndex = profileOptions.findIndex((entry) => entry.dataset.profile === profileSelect.value);
    let nextIndex = currentIndex;

    if (event.key === " " || event.key === "Enter") {
      event.preventDefault();
      selectProfile(option.dataset.profile, true);
      return;
    }

    if (event.key === "ArrowRight" || event.key === "ArrowDown") {
      nextIndex = (currentIndex + 1) % profileOptions.length;
    } else if (event.key === "ArrowLeft" || event.key === "ArrowUp") {
      nextIndex = (currentIndex - 1 + profileOptions.length) % profileOptions.length;
    } else if (event.key === "Home") {
      nextIndex = 0;
    } else if (event.key === "End") {
      nextIndex = profileOptions.length - 1;
    } else {
      return;
    }

    event.preventDefault();
    selectProfile(profileOptions[nextIndex].dataset.profile, true);
    profileOptions[nextIndex].focus();
  });
});

profileSelect.addEventListener("change", () => {
  syncProfileOptions(profileSelect.value);
});

dropzone.addEventListener("click", () => fileInput.click());
dropzone.addEventListener("keydown", (event) => {
  if (event.key === "Enter" || event.key === " ") {
    event.preventDefault();
    fileInput.click();
  }
});

fileInput.addEventListener("change", (event) => {
  handleSelection(event.target.files);
  fileInput.value = "";
});

["dragenter", "dragover"].forEach((type) => {
  dropzone.addEventListener(type, (event) => {
    event.preventDefault();
    dropzone.classList.add("dragover");
  });
});

["dragleave", "drop"].forEach((type) => {
  dropzone.addEventListener(type, (event) => {
    event.preventDefault();
    dropzone.classList.remove("dragover");
  });
});

dropzone.addEventListener("drop", (event) => {
  handleSelection(event.dataTransfer.files);
});

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
      if (!resolved) {
        resolved = true;
        if (activeCompression?.id === item.id) {
          activeCompression = null;
        }
        cleanup();
        resolve(value);
      }
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
              type: item.kind === "pdf" ? "compress-pdf" : "compress-office",
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

compressButton.addEventListener("click", async () => {
  state.isProcessing = true;
  renderFiles();

  const profile = profileSelect.value;
  const optimizeStructure = structureToggle.checked;

  let currentIndex = 0;
  const totalValid = state.files.filter((item) => item.kind).length;

  for (const item of state.files) {
    if (!item.kind) {
      continue;
    }

    currentIndex += 1;
    revokeDownload(item);
    clearOutput(item);
    const isOffice = isOfficeKind(item.kind);

    if (item.kind === "pdf" && !state.engineReady) {
      markItem(item.id, {
        statusLabel: "未完成",
        tone: "error",
        message: state.engineMessage || "PDF 压缩引擎未就绪。",
        resultBytes: 0,
        outputName: "",
        downloadUrl: "",
        outputBlob: null,
      });
      continue;
    }

    markItem(item.id, {
      statusLabel: "压缩中",
      tone: "processing",
      message: isOffice
        ? `JSZip 正在解包 ${item.kind.toUpperCase()}（${currentIndex}/${totalValid}），并按 ${profile} 档位重压缩图片。`
        : optimizeStructure
          ? `Ghostscript 压缩后将执行 QPDF 结构优化（${currentIndex}/${totalValid}），使用 ${profile} 档位。`
          : `Ghostscript WASM 正在处理（${currentIndex}/${totalValid}），使用 ${profile} 档位。`,
      resultBytes: 0,
      outputName: "",
      downloadUrl: "",
      outputBlob: null,
    });

    const result = await compressFile(item, profile, !isOffice && optimizeStructure);

    if (!result.ok) {
      markItem(item.id, {
        statusLabel: "未完成",
        tone: "error",
        message: result.error,
        resultBytes: 0,
        outputName: "",
        downloadUrl: "",
        outputBlob: null,
      });
      if (result.stopQueue) {
        break;
      }
      continue;
    }

    const blob = new Blob([result.buffer], {
      type: officeMimeByKind[item.kind] ?? "application/pdf",
    });
    markItem(item.id, {
      statusLabel: "已完成",
      tone: "success",
      message: result.message,
      resultBytes: blob.size,
      outputName: result.outputName,
      outputBlob: blob,
    });
  }

  state.isProcessing = false;
  renderFiles();
});

fileList.addEventListener("click", async (event) => {
  const link = event.target.closest(".download-link");
  if (!link) {
    return;
  }

  const item = state.files.find((entry) => entry.id === link.dataset.fileId);
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
});

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

worker = createWorker();
probeWorker();
selectProfile(profileSelect.value);
renderFiles();

window.addEventListener("beforeunload", () => {
  state.files.forEach((item) => revokeDownload(item));
});
