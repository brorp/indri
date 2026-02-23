const ACCESS_CODE = "indricantik";
const STORAGE_KEY = "bot_indri_access_code";
const MAX_FILE_MB = 70;
const MAX_FILE_BYTES = MAX_FILE_MB * 1024 * 1024;

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
  if (file.size > MAX_FILE_BYTES) {
    throw new Error(`${label} melebihi batas ${MAX_FILE_MB}MB.`);
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

loginForm.addEventListener("submit", (event) => {
  event.preventDefault();
  const code = (loginInput.value || "").trim();

  if (code !== ACCESS_CODE) {
    showError(loginError, "Access code salah.");
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

  if (needCompanion && (!companionSfxlInput.files[0] || !companionSitelistInput.files[0] || !companionTaggingInput.files[0])) {
    showError(analyzeError, "Untuk mode wpcsdm, upload SFXL, sitelist, dan tagging.");
    return;
  }

  if (!accessCode) {
    showError(analyzeError, "Session login tidak ditemukan. Login ulang.");
    setLoggedIn(false);
    return;
  }

  try {
    assertFileWithinLimit(file, "File utama");

    startLoading();

    const fileData = await fileToBase64(file);
    let companions = null;
    if (needCompanion) {
      const sfxlFile = companionSfxlInput.files[0];
      const sitelistFile = companionSitelistInput.files[0];
      const taggingFile = companionTaggingInput.files[0];

      assertFileWithinLimit(sfxlFile, "File SFXL");
      assertFileWithinLimit(sitelistFile, "File sitelist");
      assertFileWithinLimit(taggingFile, "File tagging");

      companions = {
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

    const response = await fetch("/api/analyze", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        accessCode,
        type,
        fileName: file.name,
        fileData,
        companions,
      }),
    });

    if (!response.ok) {
      let message = "Gagal memproses file.";
      try {
        const errJson = await response.json();
        if (errJson && errJson.error) message = errJson.error;
      } catch {
        // ignore json parse error
      }
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
if (savedCode === ACCESS_CODE) {
  setLoggedIn(true);
} else {
  setLoggedIn(false);
}
