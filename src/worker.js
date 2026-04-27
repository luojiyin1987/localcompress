import { state, activeCompression, finishActiveCompression } from "./state.js";
import { setEngineStatus, setStructureToggleEnabled, renderFiles } from "./render.js";

let worker;

export const createWorker = () => {
  const nextWorker = new Worker(new URL("../workers/pdf-worker.js", import.meta.url), { type: "module" });
  nextWorker.addEventListener("message", handleWorkerMessage);
  nextWorker.addEventListener("error", handleWorkerError);
  nextWorker.addEventListener("messageerror", handleWorkerMessageError);
  worker = nextWorker;
  return nextWorker;
};

export const getWorker = () => worker;

export const probeWorker = () => {
  worker.postMessage({ type: "probe-engine" });
};

export const restartWorker = (message) => {
  worker?.terminate();
  createWorker();
  state.engineReady = false;
  state.engineMessage = message;
  setEngineStatus(message);
  setStructureToggleEnabled(false);
  renderFiles();
  probeWorker();
};

export function handleWorkerMessage(event) {
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

export function handleWorkerError(event) {
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

export function handleWorkerMessageError() {
  finishActiveCompression({
    ok: false,
    id: activeCompression?.id,
    error: "Worker 消息通道错误，压缩引擎已重启。",
    stopQueue: true,
  });
  restartWorker("Worker 消息通道错误，压缩引擎已重启。");
}
