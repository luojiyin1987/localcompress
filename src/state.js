export const state = {
  files: [],
  isProcessing: false,
  engineReady: false,
  engineMessage: "压缩引擎加载中",
};

export const dropzone = document.querySelector("#dropzone");
export const fileInput = document.querySelector("#file-input");
export const fileList = document.querySelector("#file-list");
export const queueSummary = document.querySelector("#queue-summary");
export const profileSelect = document.querySelector("#profile-select");
export const structureToggle = document.querySelector("#structure-toggle");
export const compressButton = document.querySelector("#compress-button");
export const engineStatus = document.querySelector("#engine-status");
export const rowTemplate = document.querySelector("#file-row-template");

export let activeCompression = null;

export const finishActiveCompression = (result) => {
  activeCompression?.resolve(result);
};

export const setActiveCompression = (value) => {
  activeCompression = value;
};
