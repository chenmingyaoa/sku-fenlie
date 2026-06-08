const skuForm = document.getElementById("sku-splitter-form");
const skuSubmitButton = document.getElementById("sku-submit-btn");
const sourceFileInput = document.getElementById("source-file");
const selectedFileName = document.getElementById("selected-file-name");
const downloadButton = document.getElementById("download-result-btn");
const saveStatus = document.getElementById("save-status");

if (sourceFileInput && selectedFileName) {
  sourceFileInput.addEventListener("change", (event) => {
    const file = event.target.files && event.target.files[0];
    selectedFileName.textContent = file
      ? `已选择：${file.name}`
      : "支持 `.xlsx` 和 `.csv`，推荐直接上传原始 Excel。";
  });
}

if (skuForm && skuSubmitButton) {
  skuForm.addEventListener("submit", () => {
    skuSubmitButton.disabled = true;
    skuSubmitButton.textContent = "处理中，请稍候...";
  });
}

function parseDownloadFilename(contentDisposition) {
  if (!contentDisposition) return "processed.xlsx";

  const utf8Match = contentDisposition.match(/filename\*=UTF-8''([^;]+)/i);
  if (utf8Match && utf8Match[1]) {
    return decodeURIComponent(utf8Match[1]);
  }

  const asciiMatch = contentDisposition.match(/filename="([^"]+)"/i);
  if (asciiMatch && asciiMatch[1]) {
    return asciiMatch[1];
  }

  return "processed.xlsx";
}

if (downloadButton && saveStatus) {
  downloadButton.addEventListener("click", async () => {
    const exportUrl = downloadButton.dataset.exportUrl;
    if (!exportUrl) return;

    downloadButton.disabled = true;
    downloadButton.textContent = "准备下载...";
    saveStatus.textContent = "正在生成文件，请稍候...";

    try {
      const response = await fetch(exportUrl, { method: "GET" });
      if (!response.ok) {
        throw new Error("下载失败，请稍后重试。");
      }

      const blob = await response.blob();
      const filename = parseDownloadFilename(response.headers.get("Content-Disposition"));
      const blobUrl = window.URL.createObjectURL(blob);
      const link = document.createElement("a");
      link.href = blobUrl;
      link.download = filename;
      document.body.appendChild(link);
      link.click();
      link.remove();
      window.URL.revokeObjectURL(blobUrl);

      saveStatus.textContent = "文件已发送到当前设备浏览器下载，请查看浏览器下载列表。";
      downloadButton.textContent = "重新下载";
      downloadButton.disabled = false;
    } catch (error) {
      saveStatus.textContent =
        "当前浏览器可能拦截了自动下载，请点击旁边的“直接打开下载链接”，或检查浏览器下载权限。";
      downloadButton.textContent = "下载到当前设备";
      downloadButton.disabled = false;
    }
  });
}
