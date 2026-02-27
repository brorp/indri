const STORAGE_KEY = "bot_indri_access_code";
const FALLBACK_MAX_FILE_MB = 70;

let runtimeMaxFileMb = FALLBACK_MAX_FILE_MB;
let storageEnabled = false;
let storageConfigLoaded = false;
let storageConfigPromise = null;

const loginCard = document.getElementById("login-card");
const analyzerCard = document.getElementById("analyzer-card");

const loginForm = document.getElementById("login-form");
const loginInput = document.getElementById("access-code");
const loginError = document.getElementById("login-error");
const logoutBtn = document.getElementById("logout-btn");

const analyzeForm = document.getElementById("analyze-form");
const uploadFileInput = document.getElementById("upload-file");
const analyzeTypeInput = document.getElementById("analyze-type");
const companionFields = document.getElementById("companion-fields");
const companionSfxlInput = document.getElementById("companion-sfxl");
const companionSitelistInput = document.getElementById("companion-sitelist");
const companionTaggingInput = document.getElementById("companion-tagging");
const submitBtn = document.getElementById("submit-btn");
const analyzeError = document.getElementById("analyze-error");

const loadingBox = document.getElementById("loading-box");
const loadingText = document.getElementById("loading-text");

const resultBox = document.getElementById("result-box");
const resultName = document.getElementById("result-name");
const downloadLink = document.getElementById("download-link");

let loadingInterval = null;
let currentDownloadUrl = null;

function setLoggedIn(loggedIn) {
  if (loggedIn) {
    loginCard.classList.add("hidden");
    analyzerCard.classList.remove("hidden");
    return;
  }

  loginCard.classList.remove("hidden");
  analyzerCard.classList.add("hidden");
}

function showError(target, message) {
  target.textContent = message;
  target.classList.remove("hidden");
}

function hideError(target) {
  target.textContent = "";
  target.classList.add("hidden");
}

function startLoading() {
  submitBtn.disabled = true;
  loadingBox.classList.remove("hidden");

  let dots = 0;
  loadingInterval = setInterval(() => {
    dots = (dots + 1) % 4;
    loadingText.textContent = `Memproses data${".".repeat(dots)}`;
  }, 300);
}

function stopLoading() {
  submitBtn.disabled = false;
  loadingBox.classList.add("hidden");

  if (loadingInterval) {
    clearInterval(loadingInterval);
    loadingInterval = null;
  }
}

function fileToBase64(file) {
  return new Promise((resolve, reject) => {
    const reader = new FileReader();
    reader.onload = () => resolve(reader.result);
    reader.onerror = () => reject(new Error("Gagal membaca file upload."));
    reader.readAsDataURL(file);
  });
}

function assertFileWithinLimit(file, label) {
  if (!file) return;
  const maxFileBytes = runtimeMaxFileMb * 1024 * 1024;
  if (file.size > maxFileBytes) {
    throw new Error(`${label} melebihi batas ${runtimeMaxFileMb}MB.`);
  }
}

function isWpcMode(type) {
  return type === "wpcsdm-step1" || type === "wpcsdm-transform";
}

function setCompanionRequired(required) {
  companionSfxlInput.required = required;
  companionSitelistInput.required = required;
  companionTaggingInput.required = required;
}

function toggleCompanionFields() {
  const shouldShow = isWpcMode(analyzeTypeInput.value);
  companionFields.classList.toggle("hidden", !shouldShow);
  setCompanionRequired(shouldShow);

  if (!shouldShow) {
    companionSfxlInput.value = "";
    companionSitelistInput.value = "";
    companionTaggingInput.value = "";
  }
}

function applyLogin(code) {
  localStorage.setItem(STORAGE_KEY, code);
  setLoggedIn(true);
  hideError(loginError);
}

function logout() {
  localStorage.removeItem(STORAGE_KEY);
  stopLoading();
  setLoggedIn(false);
  loginInput.value = "";
  uploadFileInput.value = "";
  companionSfxlInput.value = "";
  companionSitelistInput.value = "";
  companionTaggingInput.value = "";
  hideError(analyzeError);
  resultBox.classList.add("hidden");

  if (currentDownloadUrl) {
    URL.revokeObjectURL(currentDownloadUrl);
    currentDownloadUrl = null;
  }
}

async function parseErrorResponse(response, fallbackMessage) {
  const rawText = await response.text();
  if (!rawText || !rawText.trim()) return fallbackMessage;

  try {
    const errJson = JSON.parse(rawText);
    if (errJson && errJson.error) return errJson.error;
  } catch {
    // keep raw text fallback
  }

  return rawText.trim();
}

async function loadStorageConfig() {
  try {
    const response = await fetch("/api/storage/config", { cache: "no-store" });
    if (!response.ok) {
      throw new Error(`HTTP ${response.status}`);
    }

    const config = await response.json();
    storageEnabled = Boolean(config && config.enabled);

    const parsedMaxMb = Number(config?.maxFileMb);
    if (Number.isFinite(parsedMaxMb) && parsedMaxMb > 0) {
      runtimeMaxFileMb = parsedMaxMb;
    }
  } catch {
    storageEnabled = false;
  } finally {
    storageConfigLoaded = true;
  }
}

function ensureStorageConfig() {
  if (storageConfigLoaded) {
    return Promise.resolve();
  }

  if (!storageConfigPromise) {
    storageConfigPromise = loadStorageConfig().finally(() => {
      storageConfigPromise = null;
    });
  }

  return storageConfigPromise;
}

async function getSignedUploadUrl({ accessCode, fileName, type }) {
  const response = await fetch("/api/storage/upload-url", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
    },
    body: JSON.stringify({
      accessCode,
      fileName,
      type,
    }),
  });

  if (!response.ok) {
    const message = await parseErrorResponse(
      response,
      `Gagal membuat upload URL (HTTP ${response.status}).`
    );
    throw new Error(message);
  }

  const payload = await response.json();
  if (!payload.uploadUrl || !payload.path) {
    throw new Error("Upload URL Supabase tidak valid.");
  }

  return payload;
}

async function uploadFileToStorage({ accessCode, file, type }) {
  const { uploadUrl, path } = await getSignedUploadUrl({
    accessCode,
    fileName: file.name,
    type,
  });

  const uploadResponse = await fetch(uploadUrl, {
    method: "PUT",
    headers: {
      "Content-Type": file.type || "application/octet-stream",
    },
    body: file,
  });

  if (!uploadResponse.ok) {
    const rawText = await uploadResponse.text();
    const detail = rawText && rawText.trim() ? `: ${rawText.trim()}` : "";
    throw new Error(`Gagal upload ke Supabase (HTTP ${uploadResponse.status})${detail}`);
  }

  return path;
}

loginForm.addEventListener("submit", (event) => {
  event.preventDefault();
  const code = (loginInput.value || "").trim();

  if (!code) {
    showError(loginError, "Access code wajib diisi.");
    return;
  }

  applyLogin(code);
});

logoutBtn.addEventListener("click", logout);
analyzeTypeInput.addEventListener("change", toggleCompanionFields);

analyzeForm.addEventListener("submit", async (event) => {
  event.preventDefault();
  hideError(analyzeError);
  resultBox.classList.add("hidden");

  const file = uploadFileInput.files && uploadFileInput.files[0];
  const type = analyzeTypeInput.value;
  const accessCode = localStorage.getItem(STORAGE_KEY);
  const needCompanion = isWpcMode(type);

  if (!file) {
    showError(analyzeError, "Pilih file dulu.");
    return;
  }

  if (
    needCompanion &&
    (!companionSfxlInput.files[0] || !companionSitelistInput.files[0] || !companionTaggingInput.files[0])
  ) {
    showError(analyzeError, "Untuk mode wpcsdm, upload SFXL, sitelist, dan tagging.");
    return;
  }

  if (!accessCode) {
    showError(analyzeError, "Session login tidak ditemukan. Login ulang.");
    setLoggedIn(false);
    return;
  }

  try {
    await ensureStorageConfig();
    startLoading();

    const payload = {
      accessCode,
      type,
      fileName: file.name,
    };

    if (storageEnabled) {
      const storage = {
        inputPath: await uploadFileToStorage({ accessCode, file, type }),
      };

      if (needCompanion) {
        const [sfxlPath, sitelistPath, taggingPath] = await Promise.all([
          uploadFileToStorage({
            accessCode,
            file: companionSfxlInput.files[0],
            type: `${type}-sfxl`,
          }),
          uploadFileToStorage({
            accessCode,
            file: companionSitelistInput.files[0],
            type: `${type}-sitelist`,
          }),
          uploadFileToStorage({
            accessCode,
            file: companionTaggingInput.files[0],
            type: `${type}-tagging`,
          }),
        ]);

        storage.companions = {
          sfxl: sfxlPath,
          sitelist: sitelistPath,
          tagging: taggingPath,
        };
      }

      payload.storage = storage;
    } else {
      assertFileWithinLimit(file, "File utama");
      payload.fileData = await fileToBase64(file);

      if (needCompanion) {
        const sfxlFile = companionSfxlInput.files[0];
        const sitelistFile = companionSitelistInput.files[0];
        const taggingFile = companionTaggingInput.files[0];

        assertFileWithinLimit(sfxlFile, "File SFXL");
        assertFileWithinLimit(sitelistFile, "File sitelist");
        assertFileWithinLimit(taggingFile, "File tagging");

        payload.companions = {
          sfxl: {
            fileName: sfxlFile.name,
            fileData: await fileToBase64(sfxlFile),
          },
          sitelist: {
            fileName: sitelistFile.name,
            fileData: await fileToBase64(sitelistFile),
          },
          tagging: {
            fileName: taggingFile.name,
            fileData: await fileToBase64(taggingFile),
          },
        };
      }
    }

    const response = await fetch("/api/analyze", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payload),
    });

    if (!response.ok) {
      if (response.status === 413) {
        throw new Error(
          "Upload terlalu besar untuk request body Vercel. Aktifkan Supabase Storage agar upload langsung ke bucket."
        );
      }

      const message = await parseErrorResponse(
        response,
        `Gagal memproses file (HTTP ${response.status}).`
      );
      throw new Error(message);
    }

    const blob = await response.blob();
    const outputName = response.headers.get("x-output-filename") || "output.xlsx";
    const blobUrl = URL.createObjectURL(blob);

    if (currentDownloadUrl) {
      URL.revokeObjectURL(currentDownloadUrl);
    }
    currentDownloadUrl = blobUrl;

    downloadLink.href = blobUrl;
    downloadLink.download = outputName;
    resultName.textContent = outputName;
    resultBox.classList.remove("hidden");
  } catch (err) {
    showError(analyzeError, err.message || "Unknown error.");
  } finally {
    stopLoading();
  }
});

const savedCode = localStorage.getItem(STORAGE_KEY);
stopLoading();
toggleCompanionFields();
ensureStorageConfig();
if (savedCode) {
  setLoggedIn(true);
} else {
  setLoggedIn(false);
}
