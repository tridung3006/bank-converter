const form = document.querySelector("#upload-form");
const fileInput = document.querySelector("#file-input");
const dropZone = document.querySelector("#drop-zone");
const selectedFiles = document.querySelector("#selected-files");
const clearButton = document.querySelector("#clear-button");
const convertButton = document.querySelector("#convert-button");
const statusBox = document.querySelector("#status");

let files = [];

fileInput.addEventListener("change", () => {
  setFiles(Array.from(fileInput.files || []));
});

clearButton.addEventListener("click", () => {
  setFiles([]);
  fileInput.value = "";
});

dropZone.addEventListener("dragover", (event) => {
  event.preventDefault();
  dropZone.classList.add("is-dragging");
});

dropZone.addEventListener("dragleave", () => {
  dropZone.classList.remove("is-dragging");
});

dropZone.addEventListener("drop", (event) => {
  event.preventDefault();
  dropZone.classList.remove("is-dragging");
  setFiles(Array.from(event.dataTransfer.files || []).filter(isExcelFile));
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();

  if (files.length === 0) {
    setStatus("Vui lòng chọn ít nhất một file Excel.", "error");
    return;
  }

  const formData = new FormData();
  for (const file of files) {
    formData.append("files", file);
  }

  convertButton.disabled = true;
  setStatus("Đang convert file, vui lòng chờ...", "");

  try {
    const response = await fetch("/convert", {
      method: "POST",
      body: formData,
    });

    if (!response.ok) {
      const detail = await response.json().catch(() => ({}));
      throw new Error(detail.error || "Convert file thất bại.");
    }

    const blob = await response.blob();
    const downloadName = getDownloadName(response) || "ERP_Import.xlsx";
    const savedPath = getSavedPath(response);
    const downloadUrl = URL.createObjectURL(blob);
    const link = document.createElement("a");
    link.href = downloadUrl;
    link.download = downloadName;
    document.body.appendChild(link);
    link.click();
    link.remove();
    URL.revokeObjectURL(downloadUrl);

    setStatus(`Đã convert xong ${files.length} file. File đã lưu tại: ${savedPath || "Downloads"}`, "success");
  } catch (error) {
    setStatus(error.message, "error");
  } finally {
    convertButton.disabled = files.length === 0;
  }
});

function setFiles(nextFiles) {
  files = nextFiles;
  renderFiles();
  convertButton.disabled = files.length === 0;
  statusBox.textContent = "";
  statusBox.className = "status";
}

function renderFiles() {
  selectedFiles.innerHTML = "";

  if (files.length === 0) {
    const empty = document.createElement("li");
    empty.className = "empty";
    empty.textContent = "Chưa có file nào.";
    selectedFiles.appendChild(empty);
    return;
  }

  for (const file of files) {
    const item = document.createElement("li");
    const name = document.createElement("span");
    const size = document.createElement("span");
    name.textContent = file.name;
    size.className = "file-size";
    size.textContent = formatSize(file.size);
    item.append(name, size);
    selectedFiles.appendChild(item);
  }
}

function setStatus(message, type) {
  statusBox.textContent = message;
  statusBox.className = `status ${type || ""}`.trim();
}

function isExcelFile(file) {
  return /\.(xlsx|xls|png|jpe?g|webp|bmp)$/i.test(file.name) || file.type.startsWith("image/");
}

function formatSize(bytes) {
  if (bytes < 1024) return `${bytes} B`;
  if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
  return `${(bytes / 1024 / 1024).toFixed(1)} MB`;
}

function getDownloadName(response) {
  const disposition = response.headers.get("Content-Disposition") || "";
  const match = disposition.match(/filename="([^"]+)"/i);
  return match ? match[1] : "";
}

function getSavedPath(response) {
  const savedPath = response.headers.get("X-Saved-Path") || "";
  return savedPath ? decodeURIComponent(savedPath) : "";
}
