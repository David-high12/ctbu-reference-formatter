(() => {
  "use strict";

  const get = (id) => document.getElementById(id);
  const chooser = get("file-input");
  const chooseButton = get("choose-file");
  const processButton = get("process-file");
  const dropZone = get("drop-zone");
  const resultSection = get("result");
  const downloadButton = get("download");
  const errorMessage = get("error");
  let file;
  let output;
  let busy = false;

  function showError(message) {
    errorMessage.textContent = message;
    errorMessage.hidden = !message;
  }

  function select(files) {
    if (busy || !files.length) return;
    resultSection.hidden = true;
    output = null;
    file = null;
    const candidate = files[0];
    let error = "";
    if (files.length !== 1) error = "请一次选择一份论文。";
    else if (!/\.docx$/iu.test(candidate.name)) error = "请选择 .docx 文件。旧版 .doc 可以在 Word 或 WPS 中另存为 .docx。";
    else if (candidate.size > 30 * 1024 * 1024) error = "Word 文件超过 30 MB，请选择更小的 .docx。";
    else file = candidate;

    showError(error);
    get("file-title").textContent = file ? file.name : "拖入 Word 论文";
    get("file-details").textContent = file
      ? (file.size < 1024 * 1024 ? Math.max(1, Math.round(file.size / 1024)) + " KB" : (file.size / 1024 / 1024).toFixed(1) + " MB")
      : "支持 .docx · 最大 30 MB";
    chooseButton.textContent = file ? "更换文件" : "选择文件";
    processButton.disabled = !file;
    if (file) get("status").textContent = "已选择 " + file.name;
  }

  function renderResult(result) {
    get("total").textContent = String(result.references.length);
    get("changed").textContent = String(result.changed);
    get("review-count").textContent = String(result.reviews.length);
    get("result-summary").textContent = result.reviews.length
      ? "待复核条目已保留原文。"
      : result.changed ? "可以下载整理后的论文。" : "已识别的条目无需调整。";
    get("boundary-notice").hidden = !result.stoppedAtUnnumbered;

    const reviewBox = get("review-box");
    reviewBox.hidden = !result.reviews.length;
    reviewBox.open = !!result.reviews.length;
    get("review-list").replaceChildren(...result.reviews.map(({ number, reason }) => {
      const line = document.createElement("li");
      const label = document.createElement("strong");
      label.textContent = number;
      line.append(label, document.createTextNode(reason));
      return line;
    }));
    output = result.bytes;
    resultSection.hidden = false;
    get("result-heading").focus();
  }

  chooseButton.addEventListener("click", () => {
    chooser.value = "";
    chooser.click();
  });
  chooser.addEventListener("change", () => select(chooser.files));
  for (const event of ["dragover", "drop"]) {
    document.addEventListener(event, (e) => {
      if (Array.from(e.dataTransfer?.types ?? []).includes("Files")) e.preventDefault();
    });
  }
  dropZone.addEventListener("dragover", (event) => {
    if (!Array.from(event.dataTransfer?.types ?? []).includes("Files")) return;
    event.preventDefault();
    event.dataTransfer.dropEffect = busy ? "none" : "copy";
    dropZone.classList.toggle("dragging", !busy);
  });
  dropZone.addEventListener("dragleave", (event) => {
    if (!dropZone.contains(event.relatedTarget)) dropZone.classList.remove("dragging");
  });
  dropZone.addEventListener("drop", (event) => {
    event.preventDefault();
    dropZone.classList.remove("dragging");
    select(event.dataTransfer.files);
  });

  processButton.addEventListener("click", async () => {
    if (!file || busy) return;
    busy = true;
    showError("");
    resultSection.hidden = true;
    output = null;
    chooseButton.disabled = true;
    processButton.disabled = true;
    const idleLabel = processButton.innerHTML;
    processButton.textContent = "正在整理…";
    get("status").textContent = "正在整理参考文献";
    try {
      renderResult(await window.formatWord(await file.arrayBuffer()));
      get("status").textContent = "参考文献整理完成";
    } catch (error) {
      showError(error instanceof Error ? error.message : "整理失败，请检查 Word 文件后重试。");
      get("status").textContent = "整理失败";
    } finally {
      busy = false;
      processButton.innerHTML = idleLabel;
      chooseButton.disabled = false;
      processButton.disabled = false;
    }
  });

  downloadButton.addEventListener("click", () => {
    if (!output || !file) return;
    const link = document.createElement("a");
    const url = URL.createObjectURL(new Blob([output], {
      type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    }));
    link.href = url;
    link.download = file.name.replace(/\.docx$/iu, "") + "_参考文献整理.docx";
    document.body.append(link);
    link.click();
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  });
})();
