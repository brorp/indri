const express = require("express");
const fs = require("fs");
const fsp = require("fs/promises");
const os = require("os");
const path = require("path");
const { randomUUID } = require("crypto");
const { spawn } = require("child_process");

let createClient = null;
try {
  ({ createClient } = require("@supabase/supabase-js"));
} catch {
  createClient = null;
}

const app = express();
const PORT = Number(process.env.PORT || 3000);
const ACCESS_CODE = process.env.BOT_INDRI_CODE || "indricantik";
const MAX_FILE_MB = Number(process.env.MAX_FILE_MB || 70);
const MAX_FILE_BYTES = MAX_FILE_MB * 1024 * 1024;
const SUPABASE_URL = process.env.SUPABASE_URL || "";
const SUPABASE_SERVICE_ROLE_KEY = process.env.SUPABASE_SERVICE_ROLE_KEY || "";
const SUPABASE_BUCKET = process.env.SUPABASE_BUCKET || "";
const SUPABASE_UPLOAD_PREFIX = String(process.env.SUPABASE_UPLOAD_PREFIX || "indri-uploads").replace(
  /^\/+|\/+$/g,
  ""
);
const SUPABASE_DELETE_AFTER_PROCESS =
  String(process.env.SUPABASE_DELETE_AFTER_PROCESS || "true").toLowerCase() !== "false";
const SUPABASE_ENABLED = Boolean(SUPABASE_URL && SUPABASE_SERVICE_ROLE_KEY && SUPABASE_BUCKET);
const supabase = SUPABASE_ENABLED && createClient
  ? createClient(SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY, {
      auth: {
        persistSession: false,
        autoRefreshToken: false,
      },
    })
  : null;

const BASE_DIR = __dirname;
const SCRIPT_DIR = BASE_DIR;
const ASSET_DIR = path.join(BASE_DIR, "assets");
const ANALYZER_ABS_PATHS = {
  "ippd-hourly.js": path.join(BASE_DIR, "ippd-hourly.js"),
  "twamp-hourly.js": path.join(BASE_DIR, "twamp-hourly.js"),
  "twamp-daily.js": path.join(BASE_DIR, "twamp-daily.js"),
  "s1-hourly.js": path.join(BASE_DIR, "s1-hourly.js"),
  "s1-daily.js": path.join(BASE_DIR, "s1-daily.js"),
  "wpcsdm_step1.js": path.join(BASE_DIR, "wpcsdm_step1.js"),
  "wpcsdm_transform.js": path.join(BASE_DIR, "wpcsdm_transform.js"),
};

const TYPE_MAP = {
  "ippd-hourly": { script: "ippd-hourly.js" },
  "twamp-hourly": { script: "twamp-hourly.js" },
  "twamp-daily": { script: "twamp-daily.js" },
  "s1-hourly": { script: "s1-hourly.js" },
  "s1-daily": { script: "s1-daily.js" },
  "wpcsdm-step1": {
    script: "wpcsdm_step1.js",
    mode: "wpc-step1",
  },
  "wpcsdm-transform": {
    script: "wpcsdm_transform.js",
    mode: "wpc-transform",
  },
};

const WPC_DEFAULT_ASSETS = {
  sfxl: "NEW SFXL 14022026.xlsx",
  sitelist: "sitelist_mocn_20260212.csv",
  tagging: "TAGGING 13022026.xlsx",
};

app.use(express.json({ limit: "120mb" }));
app.use(express.static(path.join(BASE_DIR, "public")));

// Keep static references to analyzer files so deployment tracers include them.
for (const p of Object.values(ANALYZER_ABS_PATHS)) {
  try {
    fs.accessSync(p, fs.constants.F_OK);
  } catch {
    // optional: missing files are handled during request validation.
  }
}

function sanitizeFileName(name, fallback = "input.xlsx") {
  const picked = path.basename(name || fallback);
  const cleaned = picked.replace(/[^a-zA-Z0-9._-]/g, "_");
  if (!cleaned) return fallback;
  return cleaned;
}

function sanitizePathSegment(value, fallback = "misc") {
  const cleaned = String(value || "")
    .trim()
    .replace(/[^a-zA-Z0-9._-]/g, "-")
    .replace(/-+/g, "-")
    .replace(/^-+|-+$/g, "");
  return cleaned || fallback;
}

function normalizeStoragePath(input) {
  return String(input || "").replace(/^\/+/, "").trim();
}

function encodeStoragePathForUrl(objectPath) {
  return normalizeStoragePath(objectPath)
    .split("/")
    .filter(Boolean)
    .map((seg) => encodeURIComponent(seg))
    .join("/");
}

function buildStorageObjectPath(type, fileName) {
  const day = new Date().toISOString().slice(0, 10);
  const typePart = sanitizePathSegment(type, "misc");
  const safeFileName = sanitizeFileName(fileName, "upload.xlsx");
  const prefix = SUPABASE_UPLOAD_PREFIX ? `${SUPABASE_UPLOAD_PREFIX}/` : "";
  return `${prefix}${typePart}/${day}/${randomUUID()}-${safeFileName}`;
}

function resolveSignedUploadUrl(signedUrl, objectPath, token) {
  if (signedUrl && /^https?:\/\//i.test(signedUrl)) {
    return signedUrl;
  }

  const base = SUPABASE_URL.replace(/\/+$/g, "");
  if (signedUrl) {
    if (signedUrl.startsWith("/storage/v1/")) return `${base}${signedUrl}`;
    if (signedUrl.startsWith("/")) return `${base}/storage/v1${signedUrl}`;
    return `${base}/storage/v1/${signedUrl}`;
  }

  if (!token) return null;
  const encodedPath = encodeStoragePathForUrl(objectPath);
  return `${base}/storage/v1/object/upload/sign/${SUPABASE_BUCKET}/${encodedPath}?token=${encodeURIComponent(
    token
  )}`;
}

function ensureSupabaseReady() {
  if (supabase) return;
  if (!createClient) {
    throw new Error("Package @supabase/supabase-js belum ter-install di server.");
  }
  throw new Error(
    "Supabase storage belum dikonfigurasi. Set SUPABASE_URL, SUPABASE_SERVICE_ROLE_KEY, dan SUPABASE_BUCKET."
  );
}

function isManagedStoragePath(objectPath) {
  const normalized = normalizeStoragePath(objectPath);
  if (!normalized || !SUPABASE_UPLOAD_PREFIX) return false;
  return normalized.startsWith(`${SUPABASE_UPLOAD_PREFIX}/`);
}

function decodeBase64File(fileData) {
  if (typeof fileData !== "string" || !fileData.trim()) {
    throw new Error("File data is empty.");
  }

  const commaIdx = fileData.indexOf(",");
  const pureBase64 = commaIdx >= 0 ? fileData.slice(commaIdx + 1) : fileData;

  // Approximate decoded size from base64 length to reject oversized payload early.
  const padding = pureBase64.endsWith("==") ? 2 : pureBase64.endsWith("=") ? 1 : 0;
  const estimatedBytes = Math.floor((pureBase64.length * 3) / 4) - padding;
  if (estimatedBytes > MAX_FILE_BYTES) {
    throw new Error(`File exceeds ${MAX_FILE_MB}MB limit.`);
  }

  const buf = Buffer.from(pureBase64, "base64");

  if (!buf.length) {
    throw new Error("Failed to decode file upload.");
  }

  if (buf.length > MAX_FILE_BYTES) {
    throw new Error(`File exceeds ${MAX_FILE_MB}MB limit.`);
  }

  return buf;
}

async function writeUploadedFile(jobDir, uploaded, fallbackName) {
  if (!uploaded || typeof uploaded !== "object") return null;

  const { fileName, fileData } = uploaded;
  if (!fileName || !fileData) return null;

  const safeName = sanitizeFileName(fileName, fallbackName);
  const filePath = path.join(jobDir, safeName);
  const fileBuffer = decodeBase64File(fileData);
  await fsp.writeFile(filePath, fileBuffer);
  return filePath;
}

async function fileExists(targetPath) {
  try {
    await fsp.access(targetPath, fs.constants.F_OK);
    return true;
  } catch {
    return false;
  }
}

async function resolveScriptPath(scriptFileName) {
  const candidates = [
    ANALYZER_ABS_PATHS[scriptFileName],
    path.join(SCRIPT_DIR, scriptFileName),
    path.join(BASE_DIR, "scripts", scriptFileName),
    path.join(process.cwd(), scriptFileName),
    path.join(process.cwd(), "scripts", scriptFileName),
  ].filter(Boolean);

  for (const candidate of candidates) {
    if (await fileExists(candidate)) return candidate;
  }

  return { found: null, checked: candidates };
}

async function resolveCompanionAsset(fileName, envVarName) {
  const envPath = process.env[envVarName];
  const candidates = [
    envPath,
    path.join(ASSET_DIR, fileName),
    path.join(BASE_DIR, "..", fileName),
  ].filter(Boolean);

  for (const p of candidates) {
    if (await fileExists(p)) return p;
  }

  throw new Error(
    `Companion file not found: ${fileName}. Put it in indri/assets or set ${envVarName}.`
  );
}

async function downloadStorageObjectToPath(objectPath, outputPath) {
  ensureSupabaseReady();
  const normalizedPath = normalizeStoragePath(objectPath);
  if (!normalizedPath) {
    throw new Error("Storage path kosong.");
  }

  const { data, error } = await supabase.storage.from(SUPABASE_BUCKET).download(normalizedPath);
  if (error) {
    throw new Error(`Gagal download dari Supabase (${normalizedPath}): ${error.message}`);
  }

  if (!data) {
    throw new Error(`Object Supabase tidak ditemukan: ${normalizedPath}`);
  }

  const arrayBuffer = await data.arrayBuffer();
  const fileBuffer = Buffer.from(arrayBuffer);
  if (!fileBuffer.length) {
    throw new Error(`Object Supabase kosong: ${normalizedPath}`);
  }

  await fsp.writeFile(outputPath, fileBuffer);
  return normalizedPath;
}

async function writeCompanionFromAnySource({
  jobDir,
  uploaded,
  storagePath,
  fallbackName,
  assetEnvVarName,
  managedStorageCleanupPaths,
}) {
  const uploadedPath = await writeUploadedFile(jobDir, uploaded, fallbackName);
  if (uploadedPath) return uploadedPath;

  if (storagePath) {
    const targetPath = path.join(jobDir, sanitizeFileName(fallbackName, "companion.xlsx"));
    const normalizedStoragePath = await downloadStorageObjectToPath(storagePath, targetPath);
    if (isManagedStoragePath(normalizedStoragePath)) {
      managedStorageCleanupPaths.push(normalizedStoragePath);
    }
    return targetPath;
  }

  return resolveCompanionAsset(fallbackName, assetEnvVarName);
}

async function maybeDeleteStorageObjects(objectPaths) {
  if (!SUPABASE_DELETE_AFTER_PROCESS || !supabase) return;
  const uniqueManagedPaths = [...new Set(objectPaths.filter((p) => isManagedStoragePath(p)))];
  if (!uniqueManagedPaths.length) return;

  const { error } = await supabase.storage.from(SUPABASE_BUCKET).remove(uniqueManagedPaths);
  if (error) {
    console.warn("Failed to cleanup Supabase objects:", error.message);
  }
}

function spawnNode(commandArgs, cwd) {
  return new Promise((resolve, reject) => {
    const child = spawn(process.execPath, commandArgs, {
      cwd,
      env: process.env,
      stdio: ["ignore", "pipe", "pipe"],
    });

    let stdout = "";
    let stderr = "";

    child.stdout.on("data", (chunk) => {
      stdout += chunk.toString();
    });

    child.stderr.on("data", (chunk) => {
      stderr += chunk.toString();
    });

    child.on("error", reject);
    child.on("close", (code) => {
      if (code === 0) {
        resolve({ stdout, stderr });
        return;
      }
      const message =
        `Analyzer exited with code ${code}.\n` +
        `stdout:\n${stdout || "(empty)"}\n` +
        `stderr:\n${stderr || "(empty)"}`;
      reject(new Error(message));
    });
  });
}

async function getLatestOutputXlsx(jobDir, inputPath) {
  const entries = await fsp.readdir(jobDir, { withFileTypes: true });
  const candidates = [];

  for (const entry of entries) {
    if (!entry.isFile()) continue;
    if (!entry.name.toLowerCase().endsWith(".xlsx")) continue;

    const fullPath = path.join(jobDir, entry.name);
    if (fullPath === inputPath) continue;

    const st = await fsp.stat(fullPath);
    candidates.push({ fullPath, mtimeMs: st.mtimeMs });
  }

  if (!candidates.length) {
    throw new Error("Analyzer finished but output file was not found.");
  }

  candidates.sort((a, b) => b.mtimeMs - a.mtimeMs);
  return candidates[0].fullPath;
}

async function executeAnalyzer({ type, originalName, fileBuffer, companions, storage }) {
  const normalizedType = String(type || "").trim().replace(/\.js$/i, "");
  const cfg = TYPE_MAP[normalizedType];
  if (!cfg) throw new Error("Unsupported analyzer type.");

  const resolvedScript = await resolveScriptPath(cfg.script);
  const scriptPath = typeof resolvedScript === "string" ? resolvedScript : resolvedScript.found;
  if (!scriptPath) {
    const checked = (resolvedScript.checked || []).join(" | ");
    throw new Error(`Script not found: ${cfg.script}. Checked: ${checked}`);
  }

  const tempPrefix = path.join(os.tmpdir(), "indri-");
  const jobDir = await fsp.mkdtemp(tempPrefix);
  const managedStorageCleanupPaths = [];

  try {
    const inputName = sanitizeFileName(originalName, "input.xlsx");
    const inputPath = path.join(jobDir, inputName);

    if (storage?.inputPath) {
      const normalizedStoragePath = await downloadStorageObjectToPath(storage.inputPath, inputPath);
      if (isManagedStoragePath(normalizedStoragePath)) {
        managedStorageCleanupPaths.push(normalizedStoragePath);
      }
    } else {
      if (!fileBuffer) {
        throw new Error("File upload belum diisi.");
      }
      await fsp.writeFile(inputPath, fileBuffer);
    }

    const args = [scriptPath];
    let explicitOutputPath = null;

    if (!cfg.mode) {
      args.push(inputPath);
    } else if (cfg.mode === "wpc-step1") {
      const sfxlPath = await writeCompanionFromAnySource({
        jobDir,
        uploaded: companions?.sfxl,
        storagePath: storage?.companions?.sfxl,
        fallbackName: WPC_DEFAULT_ASSETS.sfxl,
        assetEnvVarName: "BOT_INDRI_SFXL_PATH",
        managedStorageCleanupPaths,
      });
      explicitOutputPath = path.join(jobDir, "output-wpcsdm-step1.xlsx");

      args.push("--wpc", inputPath, "--sfxl", sfxlPath, "--out", explicitOutputPath);
    } else if (cfg.mode === "wpc-transform") {
      const sfxlPath = await writeCompanionFromAnySource({
        jobDir,
        uploaded: companions?.sfxl,
        storagePath: storage?.companions?.sfxl,
        fallbackName: WPC_DEFAULT_ASSETS.sfxl,
        assetEnvVarName: "BOT_INDRI_SFXL_PATH",
        managedStorageCleanupPaths,
      });
      const sitelistPath = await writeCompanionFromAnySource({
        jobDir,
        uploaded: companions?.sitelist,
        storagePath: storage?.companions?.sitelist,
        fallbackName: WPC_DEFAULT_ASSETS.sitelist,
        assetEnvVarName: "BOT_INDRI_SITELIST_PATH",
        managedStorageCleanupPaths,
      });
      const taggingPath = await writeCompanionFromAnySource({
        jobDir,
        uploaded: companions?.tagging,
        storagePath: storage?.companions?.tagging,
        fallbackName: WPC_DEFAULT_ASSETS.tagging,
        assetEnvVarName: "BOT_INDRI_TAGGING_PATH",
        managedStorageCleanupPaths,
      });
      explicitOutputPath = path.join(jobDir, "output-wpcsdm-transform.xlsx");

      args.push(
        "--wpc", inputPath,
        "--sfxl", sfxlPath,
        "--sitelist", sitelistPath,
        "--tagging", taggingPath,
        "--out", explicitOutputPath
      );
    }

    await spawnNode(args, jobDir);

    const outputPath = explicitOutputPath && (await fileExists(explicitOutputPath))
      ? explicitOutputPath
      : await getLatestOutputXlsx(jobDir, inputPath);

    const outputBuffer = await fsp.readFile(outputPath);
    const outputName = sanitizeFileName(path.basename(outputPath), "output.xlsx");

    return { outputBuffer, outputName };
  } finally {
    await fsp.rm(jobDir, { recursive: true, force: true });
    await maybeDeleteStorageObjects(managedStorageCleanupPaths);
  }
}

app.get("/api/storage/config", (_req, res) => {
  res.json({
    enabled: Boolean(supabase),
    maxFileMb: MAX_FILE_MB,
  });
});

app.post("/api/storage/upload-url", async (req, res) => {
  try {
    const { accessCode, fileName, type } = req.body || {};
    if (accessCode !== ACCESS_CODE) {
      return res.status(401).json({ error: "Access code salah." });
    }

    ensureSupabaseReady();
    if (!fileName) {
      return res.status(400).json({ error: "Nama file upload tidak valid." });
    }

    const normalizedType = String(type || "").trim().replace(/\.js$/i, "");
    const objectPath = buildStorageObjectPath(normalizedType, fileName);
    const { data, error } = await supabase
      .storage
      .from(SUPABASE_BUCKET)
      .createSignedUploadUrl(objectPath);

    if (error) {
      throw new Error(error.message);
    }

    const resolvedPath = normalizeStoragePath(data?.path || objectPath);
    const uploadUrl = resolveSignedUploadUrl(data?.signedUrl, resolvedPath, data?.token);
    if (!uploadUrl) {
      throw new Error("Signed upload URL tidak tersedia.");
    }

    return res.json({
      uploadUrl,
      path: resolvedPath,
    });
  } catch (err) {
    return res.status(500).json({
      error: err.message || "Gagal membuat upload URL.",
    });
  }
});

app.post("/api/analyze", async (req, res) => {
  try {
    const { accessCode, type, fileName, fileData, companions, storage } = req.body || {};
    const normalizedType = String(type || "").trim().replace(/\.js$/i, "");
    const hasStorageInput = Boolean(storage && storage.inputPath);

    if (accessCode !== ACCESS_CODE) {
      return res.status(401).json({ error: "Access code salah." });
    }

    if (!normalizedType || !TYPE_MAP[normalizedType]) {
      return res.status(400).json({ error: "Tipe file tidak valid." });
    }

    if (!fileName || (!fileData && !hasStorageInput)) {
      return res.status(400).json({ error: "File upload belum diisi." });
    }

    const buffer = fileData ? decodeBase64File(fileData) : null;
    const { outputBuffer, outputName } = await executeAnalyzer({
      type: normalizedType,
      originalName: fileName,
      fileBuffer: buffer,
      companions,
      storage,
    });

    res.setHeader(
      "Content-Type",
      "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    );
    res.setHeader("Content-Disposition", `attachment; filename=\"${outputName}\"`);
    res.setHeader("X-Output-Filename", outputName);
    return res.status(200).send(outputBuffer);
  } catch (err) {
    return res.status(500).json({
      error: err.message || "Unknown analyzer error",
    });
  }
});

app.get("/api/health", (_req, res) => {
  res.json({ ok: true });
});

if (require.main === module) {
  app.listen(PORT, () => {
    console.log(`indri running on http://localhost:${PORT}`);
  });
}

module.exports = app;
