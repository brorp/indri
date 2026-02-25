const express = require("express");
const fs = require("fs");
const fsp = require("fs/promises");
const os = require("os");
const path = require("path");
const { spawn } = require("child_process");

const app = express();
const PORT = Number(process.env.PORT || 3000);
const ACCESS_CODE = process.env.BOT_INDRI_CODE || "indricantik";
const MAX_FILE_MB = Number(process.env.MAX_FILE_MB || 70);
const MAX_FILE_BYTES = MAX_FILE_MB * 1024 * 1024;

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

async function executeAnalyzer({ type, originalName, fileBuffer, companions }) {
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

  try {
    const inputName = sanitizeFileName(originalName, "input.xlsx");
    const inputPath = path.join(jobDir, inputName);
    await fsp.writeFile(inputPath, fileBuffer);

    const args = [scriptPath];
    let explicitOutputPath = null;

    if (!cfg.mode) {
      args.push(inputPath);
    } else if (cfg.mode === "wpc-step1") {
      const uploadedSfxlPath = await writeUploadedFile(jobDir, companions?.sfxl, WPC_DEFAULT_ASSETS.sfxl);
      const sfxlPath = uploadedSfxlPath || await resolveCompanionAsset(WPC_DEFAULT_ASSETS.sfxl, "BOT_INDRI_SFXL_PATH");
      explicitOutputPath = path.join(jobDir, "output-wpcsdm-step1.xlsx");

      args.push("--wpc", inputPath, "--sfxl", sfxlPath, "--out", explicitOutputPath);
    } else if (cfg.mode === "wpc-transform") {
      const uploadedSfxlPath = await writeUploadedFile(jobDir, companions?.sfxl, WPC_DEFAULT_ASSETS.sfxl);
      const uploadedSitelistPath = await writeUploadedFile(jobDir, companions?.sitelist, WPC_DEFAULT_ASSETS.sitelist);
      const uploadedTaggingPath = await writeUploadedFile(jobDir, companions?.tagging, WPC_DEFAULT_ASSETS.tagging);

      const sfxlPath = uploadedSfxlPath || await resolveCompanionAsset(WPC_DEFAULT_ASSETS.sfxl, "BOT_INDRI_SFXL_PATH");
      const sitelistPath = uploadedSitelistPath || await resolveCompanionAsset(WPC_DEFAULT_ASSETS.sitelist, "BOT_INDRI_SITELIST_PATH");
      const taggingPath = uploadedTaggingPath || await resolveCompanionAsset(WPC_DEFAULT_ASSETS.tagging, "BOT_INDRI_TAGGING_PATH");
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
  }
}

app.post("/api/analyze", async (req, res) => {
  try {
    const { accessCode, type, fileName, fileData, companions } = req.body || {};
    const normalizedType = String(type || "").trim().replace(/\.js$/i, "");

    if (accessCode !== ACCESS_CODE) {
      return res.status(401).json({ error: "Access code salah." });
    }

    if (!normalizedType || !TYPE_MAP[normalizedType]) {
      return res.status(400).json({ error: "Tipe file tidak valid." });
    }

    if (!fileName || !fileData) {
      return res.status(400).json({ error: "File upload belum diisi." });
    }

    const buffer = decodeBase64File(fileData);
    const { outputBuffer, outputName } = await executeAnalyzer({
      type: normalizedType,
      originalName: fileName,
      fileBuffer: buffer,
      companions,
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
