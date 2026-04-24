import loadGhostscript from "@okathira/ghostpdl-wasm";
import ghostscriptWasmUrl from "@okathira/ghostpdl-wasm/gs.wasm?url";
import loadQpdf from "@neslinesli93/qpdf-wasm";
import qpdfWasmUrl from "@neslinesli93/qpdf-wasm/dist/qpdf.wasm?url";
import JSZip from "jszip";

let jpegModulePromise;
const getJpegEncoder = async () => {
  if (!jpegModulePromise) {
    jpegModulePromise = import("@jsquash/jpeg");
  }

  const { encode } = await jpegModulePromise;
  return encode;
};

let oxipngModulePromise;
const getOxipngEncoder = async () => {
  if (!oxipngModulePromise) {
    oxipngModulePromise = import("@jsquash/oxipng");
  }

  const { optimise } = await oxipngModulePromise;
  return optimise;
};

const profiles = {
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

const recompressImage = async (bytes, path, profile) => {
  if (!("createImageBitmap" in self) || !("OffscreenCanvas" in self)) {
    return null;
  }

  const settings = imageProfiles[profile] ?? imageProfiles.balanced;
  const isPng = path.toLowerCase().endsWith(".png");

  // 1. Decode to bitmap
  const bitmap = await createImageBitmap(
    new Blob([bytes], { type: isPng ? "image/png" : "image/jpeg" })
  );

  // 2. Scale (Archive 档 allowScale=false，跳过缩放)
  const shouldScale =
    settings.allowScale && settings.maxDimension < Math.max(bitmap.width, bitmap.height);
  const scale = shouldScale
    ? Math.min(1, settings.maxDimension / Math.max(bitmap.width, bitmap.height))
    : 1;
  const width = Math.max(1, Math.round(bitmap.width * scale));
  const height = Math.max(1, Math.round(bitmap.height * scale));

  // 3. Encode by strategy
  try {
    if (isPng) {
      // Strong: 直接转 JPEG，最小体积
      if (settings.pngStrategy === "direct-jpeg") {
        const encodeJpeg = await getJpegEncoder();
        const jpegBytes = await encodeJpeg(getBitmapImageData(bitmap, width, height, true), {
          quality: settings.quality,
          progressive: true,
          optimize_coding: true,
          trellis_multipass: true,
        });
        return jpegBytes.byteLength < bytes.byteLength
          ? { bytes: jpegBytes, converted: true }
          : null;
      }

      // Balanced / Archive: 先尝试 OxiPNG 无损优化，保留透明通道
      const optimisePng = await getOxipngEncoder();
      const pngBytes = await optimisePng(getBitmapImageData(bitmap, width, height), {
        level: 2,
        interlace: false,
        optimiseAlpha: true,
      });

      if (pngBytes.byteLength < bytes.byteLength) {
        return { bytes: pngBytes, converted: false };
      }

      // Archive: 只保留无损优化，不转 JPEG
      if (settings.pngStrategy === "oxipng-only") {
        return null;
      }

      // Balanced: fallback 到 MozJPEG
      const encodeJpeg = await getJpegEncoder();
      const jpegBytes = await encodeJpeg(getBitmapImageData(bitmap, width, height, true), {
        quality: settings.quality,
        progressive: true,
        optimize_coding: true,
        trellis_multipass: true,
      });
      return jpegBytes.byteLength < bytes.byteLength
        ? { bytes: jpegBytes, converted: true }
        : null;
    }

    // JPEG re-encode with MozJPEG
    const encodeJpeg = await getJpegEncoder();
    const jpegBytes = await encodeJpeg(getBitmapImageData(bitmap, width, height, true), {
      quality: settings.quality,
      progressive: true,
      optimize_coding: true,
      trellis_multipass: true,
    });
    return jpegBytes.byteLength < bytes.byteLength
      ? { bytes: jpegBytes, converted: false }
      : null;
  } finally {
    bitmap.close?.();
  }
};

const replaceZipTextReferences = async (zip, fromPath, toPath) => {
  const textEntries = Object.values(zip.files).filter(
    (entry) => !entry.dir && /\.rels$/i.test(entry.name)
  );

  for (const entry of textEntries) {
    const original = await entry.async("string");
    const sourceDir = getRelationshipSourceDir(entry.name);
    const updated = original.replace(
      /\bTarget=(["'])([^"']+)\1/g,
      (match, quote, target) => {
        if (resolveRelationshipTarget(entry.name, unescapeXmlAttribute(target)) !== normalizeZipPath(fromPath)) {
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

const ensureJpegContentType = async (zip) => {
  const contentTypes = zip.file("[Content_Types].xml");
  if (!contentTypes) {
    return;
  }

  const original = await contentTypes.async("string");
  if (/Extension="jpg"/i.test(original)) {
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
  let hasPngToJpeg = false;

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
          hasPngToJpeg = true;
        }
        optimizedImages += 1;
      }
    } catch (error) {
      activeLogs?.push(
        `${entry.name}: ${error instanceof Error ? error.message : "图片重压缩失败"}`
      );
    }
  }

  if (hasPngToJpeg) {
    await ensureJpegContentType(zip);
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
