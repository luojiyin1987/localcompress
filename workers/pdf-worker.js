import loadGhostscript from "@okathira/ghostpdl-wasm";
import ghostscriptWasmUrl from "@okathira/ghostpdl-wasm/gs.wasm?url";
import loadQpdf from "@neslinesli93/qpdf-wasm";
import qpdfWasmUrl from "@neslinesli93/qpdf-wasm/dist/qpdf.wasm?url";
import JSZip from "jszip";

const createLazyImport = (loader, exportName) => {
  let modulePromise;

  return async () => {
    if (!modulePromise) {
      modulePromise = loader();
    }

    try {
      const module = await modulePromise;
      return module[exportName];
    } catch (error) {
      modulePromise = null;
      throw error;
    }
  };
};

const getJpegEncoder = createLazyImport(() => import("@jsquash/jpeg"), "encode");
const getOxipngEncoder = createLazyImport(() => import("@jsquash/oxipng"), "optimise");

const pdfProfiles = {
  balanced: "/ebook",
  strong: "/screen",
  archive: "/printer",
};

const imageProfiles = {
  balanced: {
    quality: 62,
    maxDimension: 1600,
    allowScale: true,
    pngStrategy: "oxipng-then-jpeg",
  },
  strong: {
    quality: 42,
    maxDimension: 1280,
    allowScale: true,
    pngStrategy: "direct-jpeg",
  },
  archive: {
    quality: 78,
    maxDimension: Infinity,
    allowScale: false,
    pngStrategy: "oxipng-only",
  },
};

const jpegEncodeOptions = {
  progressive: true,
  optimize_coding: true,
  trellis_multipass: true,
};

let currentLogCollector = null;

const setLogCollector = (collector) => {
  currentLogCollector = collector;
};

const collectLogLine = (line) => {
  currentLogCollector?.push(line);
};

const getRecentLogs = (count = 4) => currentLogCollector?.slice(-count).join(" ") ?? "";

const createWasmModuleLoader = (loadModule, wasmFilename, wasmUrl) => {
  let modulePromise;

  return async () => {
    if (!modulePromise) {
      modulePromise = loadModule({
        locateFile: (path) => (path.endsWith(wasmFilename) ? wasmUrl : path),
        print: collectLogLine,
        printErr: collectLogLine,
      });
    }

    try {
      return await modulePromise;
    } catch (error) {
      modulePromise = null;
      throw error;
    }
  };
};

const getGhostscriptModule = createWasmModuleLoader(loadGhostscript, "gs.wasm", ghostscriptWasmUrl);
const getQpdfModule = createWasmModuleLoader(loadQpdf, "qpdf.wasm", qpdfWasmUrl);

const postEngineStatus = (ready, message, loading = false, qpdfReady = false) => {
  self.postMessage({
    type: "engine-status",
    ready,
    loading,
    message,
    qpdfReady,
  });
};

postEngineStatus(false, "压缩引擎加载中", true);

const removeFsFile = (fs, path) => {
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

  return bytes.buffer.slice(bytes.byteOffset, bytes.byteOffset + bytes.byteLength);
};

const writeBinaryFile = (fs, path, buffer) => {
  fs.writeFile(path, new Uint8Array(buffer));
};

const runModuleCommand = ({ runtime, inputPath, outputPath, inputBuffer, args, label }) => {
  removeFsFile(runtime.FS, inputPath);
  removeFsFile(runtime.FS, outputPath);

  try {
    writeBinaryFile(runtime.FS, inputPath, inputBuffer);

    const exitCode = runtime.callMain(args);
    if (exitCode !== 0) {
      throw new Error(getRecentLogs() || `${label} 退出码 ${exitCode}`);
    }

    const output = runtime.FS.readFile(outputPath, { encoding: "binary" });
    return toTransferableBuffer(output);
  } finally {
    removeFsFile(runtime.FS, inputPath);
    removeFsFile(runtime.FS, outputPath);
  }
};

const buildPdfOutputName = (name, profile, optimizeStructure) => {
  const baseName = name.replace(/\.pdf$/i, "");
  return `${baseName}.${profile}${optimizeStructure ? ".qpdf" : ""}.pdf`;
};

const buildOfficeOutputName = (name, profile, kind) => {
  const extension = kind === "docx" ? "docx" : "pptx";
  const baseName = name.replace(new RegExp(`\\.${extension}$`, "i"), "");
  return `${baseName}.${profile}.${extension}`;
};

const buildGhostscriptArgs = (setting, inputPath, outputPath) => [
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
];

const qpdfArgs = (inputPath, outputPath) => [
  "--object-streams=generate",
  "--compress-streams=y",
  "--decode-level=generalized",
  "--recompress-flate",
  "--compression-level=9",
  inputPath,
  outputPath,
];

const optimizePdfStructure = async ({ id, buffer }) => {
  const qpdf = await getQpdfModule();
  return runModuleCommand({
    runtime: qpdf,
    inputPath: `/qpdf-input-${id}.pdf`,
    outputPath: `/qpdf-output-${id}.pdf`,
    inputBuffer: buffer,
    args: qpdfArgs(`/qpdf-input-${id}.pdf`, `/qpdf-output-${id}.pdf`),
    label: "QPDF",
  });
};

const buildPdfResultMessage = ({ outputBuffer, sourceBuffer, qpdfBuffer, optimizeStructure }) => {
  if (outputBuffer === sourceBuffer) {
    return "压缩后文件未变小，已返回原文件。";
  }

  if (outputBuffer === qpdfBuffer) {
    return "压缩完成，并已通过 QPDF 重写 PDF 结构。";
  }

  if (optimizeStructure) {
    return "压缩完成，QPDF 输出未进一步变小，已保留较小版本。";
  }

  return "压缩完成，文件已在浏览器本地生成。";
};

const compressPdf = async ({ id, name, buffer, profile, optimizeStructure }) => {
  const ghostscript = await getGhostscriptModule();
  const inputPath = `/input-${id}.pdf`;
  const outputPath = `/output-${id}.pdf`;
  const setting = pdfProfiles[profile] ?? pdfProfiles.balanced;

  const ghostscriptOutputBuffer = runModuleCommand({
    runtime: ghostscript,
    inputPath,
    outputPath,
    inputBuffer: buffer,
    args: buildGhostscriptArgs(setting, inputPath, outputPath),
    label: "Ghostscript",
  });

  let qpdfOutputBuffer = null;
  let outputBuffer = ghostscriptOutputBuffer;

  if (optimizeStructure) {
    qpdfOutputBuffer = await optimizePdfStructure({ id, buffer: ghostscriptOutputBuffer });
    if (qpdfOutputBuffer.byteLength < outputBuffer.byteLength) {
      outputBuffer = qpdfOutputBuffer;
    }
  }

  if (outputBuffer.byteLength >= buffer.byteLength) {
    outputBuffer = buffer;
  }

  return {
    ok: true,
    id,
    buffer: outputBuffer,
    outputName: buildPdfOutputName(name, profile, optimizeStructure),
    message: buildPdfResultMessage({
      outputBuffer,
      sourceBuffer: buffer,
      qpdfBuffer: qpdfOutputBuffer,
      optimizeStructure,
    }),
  };
};

const officeMediaImagePattern = /^(ppt|word)\/media\/.+\.(png|jpe?g)$/i;

const isOfficeMediaImage = (path) => officeMediaImagePattern.test(path);

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

const getBitmapImageData = (bitmap, width, height, flattenAlpha = false) => {
  const canvas = new OffscreenCanvas(width, height);
  const context = canvas.getContext("2d", { alpha: !flattenAlpha });

  if (!context) {
    throw new Error("浏览器无法创建图片处理上下文。");
  }

  if (flattenAlpha) {
    context.fillStyle = "#fff";
    context.fillRect(0, 0, width, height);
  }

  context.drawImage(bitmap, 0, 0, width, height);
  return context.getImageData(0, 0, width, height);
};

const hasTransparentPixels = (imageData) => {
  const data = imageData.data;
  for (let index = 3; index < data.length; index += 4) {
    if (data[index] < 255) {
      return true;
    }
  }
  return false;
};

const getImageProfile = (profile) => imageProfiles[profile] ?? imageProfiles.balanced;

const getScaledBitmapSize = (bitmap, settings) => {
  const largestSide = Math.max(bitmap.width, bitmap.height);
  const shouldScale = settings.allowScale && settings.maxDimension < largestSide;
  const scale = shouldScale ? Math.min(1, settings.maxDimension / largestSide) : 1;

  return {
    width: Math.max(1, Math.round(bitmap.width * scale)),
    height: Math.max(1, Math.round(bitmap.height * scale)),
  };
};

const encodeJpeg = async (imageData, quality) => {
  const encodeImage = await getJpegEncoder();
  return encodeImage(imageData, {
    quality,
    ...jpegEncodeOptions,
  });
};

const encodeJpegFromBitmap = (bitmap, width, height, quality) => {
  return encodeJpeg(getBitmapImageData(bitmap, width, height, true), quality);
};

const maybeUseSmallerImage = (candidateBytes, originalBytes, converted) => {
  return candidateBytes.byteLength < originalBytes.byteLength
    ? { bytes: candidateBytes, converted }
    : null;
};

const recompressPngImage = async (bitmap, originalBytes, width, height, settings) => {
  if (settings.pngStrategy === "direct-jpeg") {
    const jpegBytes = await encodeJpegFromBitmap(bitmap, width, height, settings.quality);
    return maybeUseSmallerImage(jpegBytes, originalBytes, true);
  }

  const pngImageData = getBitmapImageData(bitmap, width, height);
  const optimisePng = await getOxipngEncoder();
  const pngBytes = await optimisePng(pngImageData, {
    level: 2,
    interlace: false,
    optimiseAlpha: true,
  });

  const optimizedPng = maybeUseSmallerImage(pngBytes, originalBytes, false);
  if (optimizedPng) {
    return optimizedPng;
  }

  if (settings.pngStrategy === "oxipng-only" || hasTransparentPixels(pngImageData)) {
    return null;
  }

  const jpegBytes = await encodeJpegFromBitmap(bitmap, width, height, settings.quality);
  return maybeUseSmallerImage(jpegBytes, originalBytes, true);
};

const recompressJpegImage = async (bitmap, originalBytes, width, height, settings) => {
  const jpegBytes = await encodeJpegFromBitmap(bitmap, width, height, settings.quality);
  return maybeUseSmallerImage(jpegBytes, originalBytes, false);
};

const recompressImage = async (bytes, path, profile) => {
  if (!("createImageBitmap" in self) || !("OffscreenCanvas" in self)) {
    return null;
  }

  const settings = getImageProfile(profile);
  const isPng = path.toLowerCase().endsWith(".png");
  const bitmap = await createImageBitmap(
    new Blob([bytes], { type: isPng ? "image/png" : "image/jpeg" })
  );

  try {
    const { width, height } = getScaledBitmapSize(bitmap, settings);
    return isPng
      ? await recompressPngImage(bitmap, bytes, width, height, settings)
      : await recompressJpegImage(bitmap, bytes, width, height, settings);
  } finally {
    bitmap.close?.();
  }
};

const decodeZipPath = (path) => {
  try {
    return decodeURIComponent(path);
  } catch {
    return path;
  }
};

const normalizeZipPath = (path) => {
  const parts = [];

  decodeZipPath(path).split("/").forEach((part) => {
    if (!part || part === ".") {
      return;
    }

    if (part === "..") {
      parts.pop();
      return;
    }

    parts.push(part);
  });

  return parts.join("/");
};

const getRelationshipSourceDir = (relsPath) => {
  const marker = "/_rels/";
  const markerIndex = relsPath.indexOf(marker);

  if (markerIndex === -1 || !relsPath.endsWith(".rels")) {
    return "";
  }

  const ownerDir = relsPath.slice(0, markerIndex);
  const sourceName = relsPath.slice(markerIndex + marker.length, -".rels".length);
  const sourcePath = ownerDir ? `${ownerDir}/${sourceName}` : sourceName;
  const slashIndex = sourcePath.lastIndexOf("/");
  return slashIndex === -1 ? "" : sourcePath.slice(0, slashIndex);
};

const resolveRelationshipTarget = (relsPath, target) => {
  if (/^[a-z][a-z0-9+.-]*:/i.test(target) || target.startsWith("/")) {
    return "";
  }

  const sourceDir = getRelationshipSourceDir(relsPath);
  return normalizeZipPath(sourceDir ? `${sourceDir}/${target}` : target);
};

const buildRelativeZipPath = (fromDir, toPath) => {
  const fromParts = normalizeZipPath(fromDir).split("/").filter(Boolean);
  const toParts = normalizeZipPath(toPath).split("/").filter(Boolean);

  while (fromParts.length > 0 && toParts.length > 0 && fromParts[0] === toParts[0]) {
    fromParts.shift();
    toParts.shift();
  }

  return [...fromParts.map(() => ".."), ...toParts].join("/");
};

const escapeXmlAttribute = (value) =>
  value
    .replace(/&/g, "&amp;")
    .replace(/"/g, "&quot;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;");

const unescapeXmlAttribute = (value) =>
  value
    .replace(/&quot;/g, "\"")
    .replace(/&apos;/g, "'")
    .replace(/&lt;/g, "<")
    .replace(/&gt;/g, ">")
    .replace(/&amp;/g, "&");

const listRelationshipEntries = (zip) =>
  Object.values(zip.files).filter((entry) => !entry.dir && /\.rels$/i.test(entry.name));

const replaceZipTextReferences = async (zip, fromPath, toPath) => {
  for (const entry of listRelationshipEntries(zip)) {
    const original = await entry.async("string");
    const sourceDir = getRelationshipSourceDir(entry.name);
    const updated = original.replace(
      /\bTarget=(["'])([^"']+)\1/g,
      (match, quote, target) => {
        const resolvedTarget = resolveRelationshipTarget(entry.name, unescapeXmlAttribute(target));
        if (resolvedTarget !== normalizeZipPath(fromPath)) {
          return match;
        }

        const relativeTarget = buildRelativeZipPath(sourceDir, toPath);
        return `Target=${quote}${escapeXmlAttribute(relativeTarget)}${quote}`;
      }
    );

    if (updated !== original) {
      zip.file(entry.name, updated);
    }
  }
};

const removeContentTypeOverrides = (xml, removedPaths) => {
  const removedPartPaths = new Set(removedPaths.map((path) => normalizeZipPath(path)));
  if (removedPartPaths.size === 0) {
    return xml;
  }

  return xml.replace(/<Override\b[^>]*>(?:\s*<\/Override>)?/gi, (tag) => {
    const partNameMatch = tag.match(/\bPartName=(["'])([^"']+)\1/i);
    if (!partNameMatch) {
      return tag;
    }

    const partPath = normalizeZipPath(unescapeXmlAttribute(partNameMatch[2]).replace(/^\/+/, ""));
    return removedPartPaths.has(partPath) ? "" : tag;
  });
};

const ensureJpegContentType = async (zip, removedPaths = []) => {
  const contentTypes = zip.file("[Content_Types].xml");
  if (!contentTypes) {
    return;
  }

  const original = await contentTypes.async("string");
  let updated = removeContentTypeOverrides(original, removedPaths);

  if (!/Extension="jpg"/i.test(updated)) {
    updated = updated.replace(
      "</Types>",
      '<Default Extension="jpg" ContentType="image/jpeg"/></Types>'
    );
  }

  if (updated !== original) {
    zip.file("[Content_Types].xml", updated);
  }
};

const writeZipBinaryFile = (zip, path, bytes) => {
  zip.file(path, bytes, {
    binary: true,
    compression: "DEFLATE",
    compressionOptions: { level: 9 },
  });
};

const listOfficeImageEntries = (zip) =>
  Object.values(zip.files).filter((entry) => !entry.dir && isOfficeMediaImage(entry.name));

const updateOfficeImageEntry = async (zip, entry, recompressed) => {
  const newPath = recompressed.converted ? getAvailableJpegPath(zip, entry.name) : entry.name;
  writeZipBinaryFile(zip, newPath, recompressed.bytes);

  if (newPath === entry.name) {
    return {
      optimized: true,
      converted: false,
      removedPath: "",
    };
  }

  zip.remove(entry.name);
  await replaceZipTextReferences(zip, entry.name, newPath);

  return {
    optimized: true,
    converted: true,
    removedPath: entry.name,
  };
};

const optimizeOfficeImages = async (zip, imageEntries, profile) => {
  let optimizedImages = 0;
  let hasPngToJpeg = false;
  const removedImagePaths = [];

  for (const entry of imageEntries) {
    try {
      const original = await entry.async("uint8array");
      const recompressed = await recompressImage(original, entry.name, profile);
      if (!recompressed) {
        continue;
      }

      const update = await updateOfficeImageEntry(zip, entry, recompressed);
      optimizedImages += 1;

      if (update.converted) {
        hasPngToJpeg = true;
      }

      if (update.removedPath) {
        removedImagePaths.push(update.removedPath);
      }
    } catch (error) {
      collectLogLine(
        `${entry.name}: ${error instanceof Error ? error.message : "图片重压缩失败"}`
      );
    }
  }

  return {
    optimizedImages,
    hasPngToJpeg,
    removedImagePaths,
  };
};

const buildOfficeResultMessage = ({ kind, optimizedImages, totalImages, outputBuffer, sourceBuffer }) => {
  if (outputBuffer === sourceBuffer) {
    return "压缩后文件未变小，已返回原文件。";
  }

  if (optimizedImages > 0) {
    return `${kind.toUpperCase()} 压缩完成，已重压缩 ${optimizedImages}/${totalImages} 张图片。`;
  }

  return `${kind.toUpperCase()} 已重新打包，${totalImages} 张图片未产生更小版本。`;
};

const compressOffice = async ({ id, name, buffer, profile, kind }) => {
  const zip = await JSZip.loadAsync(buffer);
  const imageEntries = listOfficeImageEntries(zip);
  const { optimizedImages, hasPngToJpeg, removedImagePaths } = await optimizeOfficeImages(
    zip,
    imageEntries,
    profile
  );

  if (hasPngToJpeg) {
    await ensureJpegContentType(zip, removedImagePaths);
  }

  const output = await zip.generateAsync({
    type: "uint8array",
    compression: "DEFLATE",
    compressionOptions: { level: 9 },
  });

  const outputBuffer = output.byteLength < buffer.byteLength
    ? toTransferableBuffer(output)
    : buffer;

  return {
    ok: true,
    id,
    buffer: outputBuffer,
    outputName: buildOfficeOutputName(name, profile, kind),
    message: buildOfficeResultMessage({
      kind,
      optimizedImages,
      totalImages: imageEntries.length,
      outputBuffer,
      sourceBuffer: buffer,
    }),
  };
};

const probeEngine = async () => {
  try {
    await getGhostscriptModule();

    let qpdfReady = false;
    let statusMessage = "Ghostscript WASM 已就绪";

    try {
      await getQpdfModule();
      qpdfReady = true;
      statusMessage = "Ghostscript 与 QPDF 均已就绪";
    } catch {
      statusMessage = "Ghostscript 已就绪，QPDF 结构优化不可用";
    }

    postEngineStatus(true, statusMessage, false, qpdfReady);
  } catch (error) {
    postEngineStatus(
      false,
      error instanceof Error ? error.message : "Ghostscript WASM 加载失败"
    );
  }
};

const isCompressionRequest = (message) => {
  return message.type === "compress-pdf" || message.type === "compress-office";
};

const runCompression = (message) => {
  return message.type === "compress-office" ? compressOffice(message) : compressPdf(message);
};

const handleCompressionMessage = async (message, port) => {
  try {
    const logs = [];
    setLogCollector(logs);
    const result = await runCompression(message);
    port.postMessage(result, result.ok ? [result.buffer] : []);
  } catch (error) {
    port.postMessage({
      ok: false,
      id: message.id,
      error: error instanceof Error ? error.message : "压缩 Worker 执行失败。",
    });
  } finally {
    setLogCollector(null);
  }
};

self.addEventListener("message", async (event) => {
  const message = event.data;

  if (message.type === "probe-engine") {
    await probeEngine();
    return;
  }

  if (!isCompressionRequest(message)) {
    return;
  }

  const port = event.ports[0];
  if (!port) {
    return;
  }

  await handleCompressionMessage(message, port);
});
