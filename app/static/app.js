const form = document.getElementById("report-form");
const filesInput = document.getElementById("files");
const filePicker = document.getElementById("file-picker");
const selectedFiles = document.getElementById("selected-files");
const submitButton = document.getElementById("submit-button");
const validateButton = document.getElementById("validate-button");
const resultStatus = document.getElementById("result-status");
const resultMeta = document.getElementById("result-meta");
const outputFile = document.getElementById("output-file");
const downloadLink = document.getElementById("download-link");
const normalizedPreview = document.getElementById("normalized-preview");
const understandingJson = document.getElementById("understanding-json");
const reportJson = document.getElementById("report-json");
const statusText = document.getElementById("status-text");
const statusMeta = document.getElementById("status-meta");

filesInput.addEventListener("change", renderSelectedFiles);
form.addEventListener("submit", handleGenerateSubmit);
validateButton.addEventListener("click", handleValidateSubmit);

if (filePicker) {
  ["dragenter", "dragover"].forEach((eventName) => {
    filePicker.addEventListener(eventName, (event) => {
      event.preventDefault();
      filePicker.classList.add("is-dragover");
    });
  });

  ["dragleave", "drop"].forEach((eventName) => {
    filePicker.addEventListener(eventName, () => {
      filePicker.classList.remove("is-dragover");
    });
  });

  filePicker.addEventListener("drop", (event) => {
    event.preventDefault();
    const droppedFiles = event.dataTransfer?.files;
    if (!droppedFiles?.length) {
      return;
    }

    try {
      filesInput.files = droppedFiles;
    } catch (error) {
      console.error(error);
      setStatus("此瀏覽器不支援拖曳加入檔案，請改用點擊選擇。", "error");
      return;
    }
    renderSelectedFiles();
  });
}

document.querySelectorAll("[data-copy-target]").forEach((button) => {
  button.addEventListener("click", async () => {
    const targetId = button.getAttribute("data-copy-target");
    const target = document.getElementById(targetId);
    if (!target) {
      return;
    }

    try {
      await navigator.clipboard.writeText(target.textContent ?? "");
      const originalText = button.textContent;
      button.textContent = "已複製";
      window.setTimeout(() => {
        button.textContent = originalText;
      }, 1200);
    } catch (error) {
      console.error(error);
    }
  });
});

if (statusText && statusMeta) {
  loadStatus();
}
renderSelectedFiles();

async function loadStatus() {
  try {
    const response = await fetch("/api/status");
    const payload = await response.json();
    statusText.textContent = payload.status === "ok" ? "服務正常" : payload.status;
    statusMeta.textContent = `${payload.service} | v${payload.version}`;
  } catch (error) {
    statusText.textContent = "讀取失敗";
    statusMeta.textContent = "無法取得 API 狀態";
  }
}

function renderSelectedFiles() {
  const files = Array.from(filesInput.files ?? []);
  selectedFiles.innerHTML = "";
  filePicker?.classList.toggle("has-files", files.length > 0);

  if (!files.length) {
    const emptyMessage = document.createElement("p");
    emptyMessage.className = "empty-files";
    emptyMessage.textContent = "尚未選擇檔案";
    selectedFiles.appendChild(emptyMessage);
    return;
  }

  files.forEach((file) => {
    const chip = document.createElement("span");
    chip.className = "file-chip";
    chip.textContent = `${file.name} (${formatBytes(file.size)})`;
    selectedFiles.appendChild(chip);
  });
}

async function handleGenerateSubmit(event) {
  event.preventDefault();
  await runAction({
    endpoint: "/generate-report",
    mode: "generate",
    workingText: "正在整理資料並生成問卷報告...",
    successText: "報告生成完成。",
  });
}

async function handleValidateSubmit() {
  await runAction({
    endpoint: "/validate-questionnaire",
    mode: "validate",
    workingText: "正在驗證 agent 對問卷的理解...",
    successText: "問卷理解驗證完成。",
  });
}

async function runAction({ endpoint, mode, workingText, successText }) {
  const files = Array.from(filesInput.files ?? []);
  if (!files.length) {
    setStatus("請先上傳至少一個檔案。", "error");
    return;
  }

  const formData = new FormData(form);
  formData.delete("files");
  files.forEach((file) => formData.append("files", file));

  setBusyState(true, mode);
  setStatus(workingText, "info");

  if (mode === "validate") {
    resultMeta.hidden = true;
    outputFile.textContent = "-";
    downloadLink.textContent = "尚未產生";
    downloadLink.href = "#";
    reportJson.textContent = "尚未產生";
  }

  try {
    const response = await fetch(endpoint, {
      method: "POST",
      body: formData,
    });

    const payload = await response.json();
    if (!response.ok) {
      throw new Error(payload.detail ?? "處理失敗");
    }

    normalizedPreview.textContent = JSON.stringify(payload.normalized_preview, null, 2);

    if (payload.understanding_json) {
      understandingJson.textContent = JSON.stringify(payload.understanding_json, null, 2);
    }

    if (payload.report_json) {
      reportJson.textContent = JSON.stringify(payload.report_json, null, 2);
      outputFile.textContent = payload.output_file;
      downloadLink.textContent = payload.download_url;
      downloadLink.href = payload.download_url;
      resultMeta.hidden = false;
    }

    setStatus(successText, "success");
  } catch (error) {
    const message = error instanceof Error ? error.message : "發生未預期錯誤";
    setStatus(message, "error");
  } finally {
    setBusyState(false, mode);
  }
}

function setBusyState(isBusy, mode) {
  submitButton.disabled = isBusy;
  validateButton.disabled = isBusy;

  if (!isBusy) {
    submitButton.textContent = "生成問卷報告";
    validateButton.textContent = "先驗證問卷理解";
    return;
  }

  if (mode === "generate") {
    submitButton.textContent = "生成中...";
  } else {
    validateButton.textContent = "驗證中...";
  }
}

function setStatus(message, type) {
  resultStatus.textContent = message;
  resultStatus.className = `alert ${type}`;
}

function formatBytes(bytes) {
  if (!bytes) {
    return "0 B";
  }

  const units = ["B", "KB", "MB", "GB"];
  const exponent = Math.min(Math.floor(Math.log(bytes) / Math.log(1024)), units.length - 1);
  const value = bytes / 1024 ** exponent;
  return `${value.toFixed(value >= 10 || exponent === 0 ? 0 : 1)} ${units[exponent]}`;
}
