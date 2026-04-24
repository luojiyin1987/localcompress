const worker = new Worker(new URL("./workers/pdf-worker.js", import.meta.url), { type: "module" });

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

  window.setTimeout(() => {
    URL.revokeObjectURL(downloadUrl);
  }, 10 * 60 * 1000);
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
    const channel = new MessageChannel();

    channel.port1.onmessage = (event) => {
      resolve(event.data);
      channel.port1.close();
    };

    item.file
      .arrayBuffer()
      .then((buffer) => {
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
      })
      .catch((error) => {
        resolve({
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

  for (const item of state.files) {
    if (!item.kind) {
      continue;
    }

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
        ? `JSZip 正在解包 ${item.kind.toUpperCase()}，并按 ${profile} 档位重压缩图片。`
        : optimizeStructure
          ? `Ghostscript 压缩后将执行 QPDF 结构优化，使用 ${profile} 档位。`
          : `Ghostscript WASM 正在处理，使用 ${profile} 档位。`,
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

worker.addEventListener("message", (event) => {
  if (event.data?.type !== "engine-status") {
    return;
  }

  state.engineReady = event.data.ready;
  state.engineMessage = event.data.message;

  if (event.data.ready) {
    setEngineStatus("引擎已就绪", "ready");
    structureToggle.disabled = false;
    structureToggle.closest(".checkbox-field")?.classList.remove("disabled-field");
  } else {
    setEngineStatus(event.data.message, event.data.loading ? "" : "error");
    structureToggle.disabled = true;
    structureToggle.closest(".checkbox-field")?.classList.add("disabled-field");
  }

  renderFiles();
});

worker.postMessage({ type: "probe-engine" });
renderFiles();

window.addEventListener("beforeunload", () => {
  state.files.forEach((item) => revokeDownload(item));
});
