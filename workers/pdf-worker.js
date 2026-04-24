import loadGhostscript from "@okathira/ghostpdl-wasm";
import ghostscriptWasmUrl from "@okathira/ghostpdl-wasm/gs.wasm?url";
import loadQpdf from "@neslinesli93/qpdf-wasm";
import qpdfWasmUrl from "@neslinesli93/qpdf-wasm/dist/qpdf.wasm?url";
import JSZip from "jszip";

const profiles = {
  balanced: "/ebook",
  strong: "/screen",
  archive: "/printer",
};

const pptxProfiles = {
  balanced: { quality: 0.62, maxDimension: 1600 },
  strong: { quality: 0.42, maxDimension: 1280 },
  archive: { quality: 0.78, maxDimension: 2200 },
};

let ghostscriptModulePromise;
let qpdfModulePromise;
let activeLogs = null;

self.postMessage({
  type: "engine-status",
  ready: false,
  loading: true,
  message: "压缩引擎加载中",
});

const postEngineStatus = (ready, message, loading = false) => {
  self.postMessage({
    type: "engine-status",
    ready,
    loading,
    message,
  });
};

const getModule = async () => {
  if (!ghostscriptModulePromise) {
    ghostscriptModulePromise = loadGhostscript({
      locateFile: (path) => (path.endsWith("gs.wasm") ? ghostscriptWasmUrl : path),
      print: (line) => activeLogs?.push(line),
      printErr: (line) => activeLogs?.push(line),
    });
  }

  return ghostscriptModulePromise;
};

const getQpdfModule = async () => {
  if (!qpdfModulePromise) {
    qpdfModulePromise = loadQpdf({
      locateFile: (path) => (path.endsWith("qpdf.wasm") ? qpdfWasmUrl : path),
      print: (line) => activeLogs?.push(line),
      printErr: (line) => activeLogs?.push(line),
    });
  }

  return qpdfModulePromise;
};

const removeFile = (fs, path) => {
  try {
    fs.unlink(path);
  } catch {
    // Ignore cleanup misses in Emscripten's in-memory FS.
  }
};

const toTransferableBuffer = (bytes) => {
  if (bytes.byteOffset === 0 && bytes.byteLength === bytes.buffer.byteLength) {
    return bytes.buffer;
  }

  return bytes.slice().buffer;
};

const buildOutputName = (name, profile, optimizeStructure) => {
  const baseName = name.replace(/\.pdf$/i, "");
  return `${baseName}.${profile}${optimizeStructure ? ".qpdf" : ""}.pdf`;
};

const buildOfficeOutputName = (name, profile, kind) => {
  const extension = kind === "docx" ? "docx" : "pptx";
  const baseName = name.replace(new RegExp(`\\.${extension}$`, "i"), "");
  return `${baseName}.${profile}.${extension}`;
};

const optimizePdfStructure = async ({ id, buffer }) => {
  const qpdf = await getQpdfModule();
  const inputPath = `/qpdf-input-${id}.pdf`;
  const outputPath = `/qpdf-output-${id}.pdf`;

  removeFile(qpdf.FS, inputPath);
  removeFile(qpdf.FS, outputPath);
  qpdf.FS.writeFile(inputPath, new Uint8Array(buffer));

  const exitCode = qpdf.callMain([
    "--object-streams=generate",
    "--compress-streams=y",
    "--decode-level=generalized",
    "--recompress-flate",
    "--compression-level=9",
    inputPath,
    outputPath,
  ]);

  if (exitCode !== 0) {
    const lastLog = activeLogs.slice(-4).join(" ");
    throw new Error(lastLog || `QPDF 退出码 ${exitCode}`);
  }

  const output = qpdf.FS.readFile(outputPath, { encoding: "binary" });
  removeFile(qpdf.FS, inputPath);
  removeFile(qpdf.FS, outputPath);
  return toTransferableBuffer(output);
};

const compressPdf = async ({ id, name, buffer, profile, optimizeStructure }) => {
  const Module = await getModule();
  const inputPath = `/input-${id}.pdf`;
  const outputPath = `/output-${id}.pdf`;
  const setting = profiles[profile] ?? profiles.balanced;

  removeFile(Module.FS, inputPath);
  removeFile(Module.FS, outputPath);

  activeLogs = [];
  Module.FS.writeFile(inputPath, new Uint8Array(buffer));

  const exitCode = Module.callMain([
    "-sDEVICE=pdfwrite",
    `-dPDFSETTINGS=${setting}`,
    "-dCompatibilityLevel=1.5",
    "-dDetectDuplicateImages=true",
    "-dCompressFonts=true",
    "-dSubsetFonts=true",
    "-dNOPAUSE",
    "-dQUIET",
    "-dBATCH",
    `-sOutputFile=${outputPath}`,
    inputPath,
  ]);

  if (exitCode !== 0) {
    const lastLog = activeLogs.slice(-4).join(" ");
    throw new Error(lastLog || `Ghostscript 退出码 ${exitCode}`);
  }

  const ghostscriptOutput = Module.FS.readFile(outputPath, { encoding: "binary" });
  removeFile(Module.FS, inputPath);
  removeFile(Module.FS, outputPath);

  const ghostscriptOutputBuffer = toTransferableBuffer(ghostscriptOutput);
  const outputBuffer = optimizeStructure
    ? await optimizePdfStructure({ id, buffer: ghostscriptOutputBuffer })
    : ghostscriptOutputBuffer;

  activeLogs = null;

  return {
    ok: true,
    id,
    buffer: outputBuffer,
    outputName: buildOutputName(name, profile, optimizeStructure),
    message: optimizeStructure
      ? "压缩完成，并已通过 QPDF 重写 PDF 结构。"
      : "压缩完成，文件已在浏览器本地生成。",
  };
};

const isOfficeMediaImage = (path) =>
  /^(ppt|word)\/media\/.+\.(png|jpe?g)$/i.test(path);

const getJpegPath = (path) => path.replace(/\.(png|jpe?g)$/i, ".jpg");

const getAvailableJpegPath = (zip, path) => {
  const preferredPath = getJpegPath(path);

  if (preferredPath === path || !zip.file(preferredPath)) {
    return preferredPath;
  }

  const dotIndex = preferredPath.lastIndexOf(".");
  const baseName = preferredPath.slice(0, dotIndex);
  const extension = preferredPath.slice(dotIndex);
  let index = 1;
  let candidate = `${baseName}-${index}${extension}`;

  while (zip.file(candidate)) {
    index += 1;
    candidate = `${baseName}-${index}${extension}`;
  }

  return candidate;
};

const recompressImage = async (bytes, path, profile) => {
  if (!("createImageBitmap" in self) || !("OffscreenCanvas" in self)) {
    return null;
  }

  const settings = pptxProfiles[profile] ?? pptxProfiles.balanced;
  const sourceMime = path.toLowerCase().endsWith(".png") ? "image/png" : "image/jpeg";
  const imageBlob = new Blob([bytes], { type: sourceMime });
  const bitmap = await createImageBitmap(imageBlob);
  const scale = Math.min(1, settings.maxDimension / Math.max(bitmap.width, bitmap.height));
  const width = Math.max(1, Math.round(bitmap.width * scale));
  const height = Math.max(1, Math.round(bitmap.height * scale));
  const canvas = new OffscreenCanvas(width, height);
  const context = canvas.getContext("2d", { alpha: false });

  context.fillStyle = "#fff";
  context.fillRect(0, 0, width, height);
  context.drawImage(bitmap, 0, 0, width, height);
  bitmap.close?.();

  const outputBlob = await canvas.convertToBlob({
    type: "image/jpeg",
    quality: settings.quality,
  });
  const output = new Uint8Array(await outputBlob.arrayBuffer());

  return output.byteLength < bytes.byteLength
    ? {
        bytes: output,
        converted: sourceMime === "image/png",
      }
    : null;
};

const replaceZipTextReferences = async (zip, fromPath, toPath) => {
  const replacements = [
    [fromPath, toPath],
    [fromPath.split("/").pop(), toPath.split("/").pop()],
  ];
  const textEntries = Object.values(zip.files).filter(
    (entry) =>
      !entry.dir &&
      /\.(xml|rels)$/i.test(entry.name) &&
      !/^(ppt|word)\/media\//i.test(entry.name)
  );

  await Promise.all(
    textEntries.map(async (entry) => {
      const original = await entry.async("string");
      let updated = original;

      const escapeRegExp = (s) => s.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");

      replacements.forEach(([from, to]) => {
        const re = new RegExp(escapeRegExp(from) + "\\b", "g");
        updated = updated.replace(re, to);
      });

      if (updated !== original) {
        zip.file(entry.name, updated);
      }
    })
  );
};

const ensureJpegContentType = async (zip) => {
  const contentTypes = zip.file("[Content_Types].xml");
  if (!contentTypes) {
    return;
  }

  const original = await contentTypes.async("string");
  if (/Extension="jpe?g"/i.test(original)) {
    return;
  }

  const updated = original.replace(
    "</Types>",
    '<Default Extension="jpg" ContentType="image/jpeg"/></Types>'
  );
  zip.file("[Content_Types].xml", updated);
};

const compressOffice = async ({ id, name, buffer, profile, kind }) => {
  const zip = await JSZip.loadAsync(buffer);
  const imageEntries = Object.values(zip.files).filter(
    (entry) => !entry.dir && isOfficeMediaImage(entry.name)
  );
  let optimizedImages = 0;

  for (const entry of imageEntries) {
    try {
      const original = await entry.async("uint8array");
      const recompressed = await recompressImage(original, entry.name, profile);

      if (recompressed) {
        const newPath = recompressed.converted ? getAvailableJpegPath(zip, entry.name) : entry.name;
        zip.file(newPath, recompressed.bytes, {
          binary: true,
          compression: "DEFLATE",
          compressionOptions: { level: 9 },
        });
        if (newPath !== entry.name) {
          zip.remove(entry.name);
          await replaceZipTextReferences(zip, entry.name, newPath);
          await ensureJpegContentType(zip);
        }
        optimizedImages += 1;
      }
    } catch (error) {
      activeLogs?.push(
        `${entry.name}: ${error instanceof Error ? error.message : "图片重压缩失败"}`
      );
    }
  }

  const output = await zip.generateAsync({
    type: "uint8array",
    compression: "DEFLATE",
    compressionOptions: { level: 9 },
  });

  return {
    ok: true,
    id,
    buffer: toTransferableBuffer(output),
    outputName: buildOfficeOutputName(name, profile, kind),
    message:
      optimizedImages > 0
        ? `${kind.toUpperCase()} 压缩完成，已重压缩 ${optimizedImages}/${imageEntries.length} 张图片。`
        : `${kind.toUpperCase()} 已重新打包，${imageEntries.length} 张图片未产生更小版本。`,
  };
};

self.addEventListener("message", async (event) => {
  const message = event.data;

  if (message.type === "probe-engine") {
    try {
      await getModule();
      postEngineStatus(true, "Ghostscript WASM 已就绪");
    } catch (error) {
      postEngineStatus(
        false,
        error instanceof Error ? error.message : "Ghostscript WASM 加载失败"
      );
    }
    return;
  }

  if (message.type !== "compress-pdf" && message.type !== "compress-office") {
    return;
  }

  const port = event.ports[0];

  if (!port) {
    return;
  }

  try {
    activeLogs = [];
    const result =
      message.type === "compress-office" ? await compressOffice(message) : await compressPdf(message);
    port.postMessage(result, result.ok ? [result.buffer] : []);
  } catch (error) {
    port.postMessage({
      ok: false,
      id: message.id,
      error: error instanceof Error ? error.message : "压缩 Worker 执行失败。",
    });
  } finally {
    activeLogs = null;
  }
});
