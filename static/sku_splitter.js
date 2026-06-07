const skuForm = document.getElementById("sku-splitter-form");
const skuSubmitButton = document.getElementById("sku-submit-btn");
const sourceFileInput = document.getElementById("source-file");
const selectedFileName = document.getElementById("selected-file-name");
const saveToDownloadsButton = document.getElementById("save-to-downloads-btn");
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

if (saveToDownloadsButton && saveStatus) {
  saveToDownloadsButton.addEventListener("click", async () => {
    const saveUrl = saveToDownloadsButton.dataset.saveUrl;
    if (!saveUrl) return;

    saveToDownloadsButton.disabled = true;
    saveToDownloadsButton.textContent = "保存中...";
    saveStatus.textContent = "正在保存到本地下载目录...";

    try {
      const response = await fetch(saveUrl, { method: "POST" });
      const payload = await response.json();
      if (!response.ok) {
        throw new Error(payload.error || "保存失败，请稍后重试。");
      }

      saveStatus.innerHTML = `已保存到：<code>${payload.saved_path}</code>`;
      saveToDownloadsButton.textContent = "已保存";
    } catch (error) {
      saveStatus.textContent = error.message;
      saveToDownloadsButton.disabled = false;
      saveToDownloadsButton.textContent = "保存到下载目录";
    }
  });
}
