(() => {
  "use strict";

  const $ = (sel, root = document) => root.querySelector(sel);
  const $all = (sel, root = document) => Array.from(root.querySelectorAll(sel));

  function formatBytes(bytes) {
    if (bytes < 1024) return `${bytes} B`;
    if (bytes < 1024 * 1024) return `${(bytes / 1024).toFixed(1)} KB`;
    return `${(bytes / (1024 * 1024)).toFixed(2)} MB`;
  }

  function showStep(paneEl, stepName) {
    $all(".stepbody", paneEl).forEach((el) => {
      el.hidden = el.dataset.stepbody !== stepName;
    });
    $all(".steps__item", paneEl).forEach((el) => {
      el.classList.toggle("is-active", el.dataset.step === stepName);
    });
  }

  function cycleStatus(el, messages, intervalMs = 1400) {
    let i = 0;
    el.textContent = messages[0];
    const id = setInterval(() => {
      i = (i + 1) % messages.length;
      el.textContent = messages[i];
    }, intervalMs);
    return () => clearInterval(id);
  }

  async function postJson(url, body) {
    const res = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(body),
    });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(data.error || "Something went wrong.");
    return data;
  }

  async function postForm(url, formData) {
    const res = await fetch(url, { method: "POST", body: formData });
    const data = await res.json().catch(() => ({}));
    if (!res.ok) throw new Error(data.error || "Something went wrong.");
    return data;
  }

  function renderChangeList(ul, summaryLines) {
    ul.innerHTML = "";
    summaryLines.forEach((line) => {
      const li = document.createElement("li");
      li.textContent = line;
      ul.appendChild(li);
    });
  }

  function renderResult(container, data) {
    const failed = data.failedCount || 0;
    const unverified = data.unverifiedCount || 0;
    const clean = data.status === "complete";
    const title = failed ? "Your file is ready, but some changes failed" : clean ? "Your Excel file is ready" : "Your Excel file is ready (some changes unverifiable)";
    const esc = (t) => String(t).replace(/[&<>"]/g, (c) => ({ "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;" }[c]));
    const list = (items, key) => items.map((f) => `<li><strong>${esc(f.type)}</strong>: ${esc(f[key])}</li>`).join("");
    container.className = failed ? "result result--warn" : "result";
    container.innerHTML = `
      <p class="result__title">${title}</p>
      <p class="result__meta">${data.verifiedCount} verified successfully · ${failed} failed · ${unverified} unable to verify</p>
      <a class="result__download" href="${data.downloadUrl}">Download Excel</a>
      ${failed ? `<details class="result__failed" open><summary>Failed changes (not applied or not confirmed)</summary><ul>${list(data.failed, "error")}</ul></details>` : ""}
      ${unverified ? `<details class="result__failed"><summary>Applied but could not be verified</summary><ul>${list(data.unverified, "note")}</ul></details>` : ""}
    `;
  }

  function renderErrorResult(container, message, details) {
    container.className = "result result--error";
    container.innerHTML = `
      <p class="result__title">We couldn't finish that</p>
      <p class="result__meta">${message}</p>
      ${details && details.length ? `<details class="result__failed"><summary>Details</summary><ul>${details.map((d) => `<li>${typeof d === "string" ? d : d.error}</li>`).join("")}</ul></details>` : ""}
    `;
  }

  // ----------------------------------------------------------------- Tabs
  const tabEdit = $("#tab-edit");
  const tabCreate = $("#tab-create");
  const paneEdit = $("#pane-edit");
  const paneCreate = $("#pane-create");

  function activateTab(which) {
    const editActive = which === "edit";
    tabEdit.classList.toggle("is-active", editActive);
    tabCreate.classList.toggle("is-active", !editActive);
    tabEdit.setAttribute("aria-selected", String(editActive));
    tabCreate.setAttribute("aria-selected", String(!editActive));
    paneEdit.hidden = !editActive;
    paneEdit.classList.toggle("is-active", editActive);
    paneCreate.hidden = editActive;
    paneCreate.classList.toggle("is-active", !editActive);
  }
  tabEdit.addEventListener("click", () => activateTab("edit"));
  tabCreate.addEventListener("click", () => activateTab("create"));

  // ================================================================= EDIT
  (function initEdit() {
    const dropzone = $("#edit-dropzone");
    const fileInput = $("#edit-file-input");
    const fileCard = $("#edit-filecard");
    const fileNameEl = $("#edit-filename");
    const fileSizeEl = $("#edit-filesize");
    const removeBtn = $("#edit-file-remove");
    const uploadError = $("#edit-upload-error");
    const instructionsEl = $("#edit-instructions");
    const describeError = $("#edit-describe-error");
    const processingStatus = $("#edit-processing-status");
    const changeListEl = $("#edit-change-list");
    const warningsEl = $("#edit-warnings");
    const resultEl = $("#edit-result");

    let selectedFile = null;
    let sessionId = null;
    let stopCycle = null;

    function pickFile(file) {
      if (!file) return;
      if (!file.name.toLowerCase().endsWith(".xlsx")) {
        uploadError.hidden = false;
        uploadError.textContent = "Only .xlsx files are supported.";
        return;
      }
      if (file.size > 15 * 1024 * 1024) {
        uploadError.hidden = false;
        uploadError.textContent = "That file is larger than 15 MB.";
        return;
      }
      uploadError.hidden = true;
      selectedFile = file;
      fileNameEl.textContent = file.name;
      fileSizeEl.textContent = formatBytes(file.size);
      fileCard.hidden = false;
      dropzone.hidden = true;
      showStep(paneEdit, "describe");
    }

    dropzone.addEventListener("click", () => fileInput.click());
    dropzone.addEventListener("keydown", (e) => {
      if (e.key === "Enter" || e.key === " ") fileInput.click();
    });
    fileInput.addEventListener("change", () => pickFile(fileInput.files[0]));
    ["dragover", "dragenter"].forEach((evt) =>
      dropzone.addEventListener(evt, (e) => {
        e.preventDefault();
        dropzone.classList.add("is-dragover");
      })
    );
    ["dragleave", "drop"].forEach((evt) =>
      dropzone.addEventListener(evt, (e) => {
        e.preventDefault();
        dropzone.classList.remove("is-dragover");
      })
    );
    dropzone.addEventListener("drop", (e) => pickFile(e.dataTransfer.files[0]));

    removeBtn.addEventListener("click", () => {
      selectedFile = null;
      fileInput.value = "";
      fileCard.hidden = true;
      dropzone.hidden = false;
    });

    $("#edit-back-to-upload").addEventListener("click", () => showStep(paneEdit, "upload"));

    $all(".example-chip", paneEdit).forEach((chip) => {
      chip.addEventListener("click", () => {
        instructionsEl.value = chip.textContent;
        instructionsEl.focus();
      });
    });

    $("#edit-analyze").addEventListener("click", async () => {
      describeError.hidden = true;
      const instructions = instructionsEl.value.trim();
      if (!instructions) {
        describeError.hidden = false;
        describeError.textContent = "Please describe at least one change.";
        return;
      }
      if (!selectedFile) {
        showStep(paneEdit, "upload");
        return;
      }

      showStep(paneEdit, "processing");
      stopCycle = cycleStatus(processingStatus, [
        "Analyzing workbook…",
        "Understanding your instructions…",
        "Planning changes…",
      ]);

      try {
        const formData = new FormData();
        formData.append("file", selectedFile);
        formData.append("instructions", instructions);
        const data = await postForm("/api/edit/preview", formData);
        sessionId = data.sessionId;
        renderChangeList(changeListEl, data.summary);
        if (data.warnings && data.warnings.length) {
          warningsEl.hidden = false;
          warningsEl.textContent = data.warnings.join("\n");
        } else {
          warningsEl.hidden = true;
        }
        showStep(paneEdit, "review");
      } catch (err) {
        showStep(paneEdit, "describe");
        describeError.hidden = false;
        describeError.textContent = err.message;
      } finally {
        if (stopCycle) stopCycle();
      }
    });

    $("#edit-back-to-describe").addEventListener("click", () => showStep(paneEdit, "describe"));

    $("#edit-confirm").addEventListener("click", async () => {
      showStep(paneEdit, "processing");
      stopCycle = cycleStatus(processingStatus, [
        "Applying changes…",
        "Verifying workbook…",
        "Preparing your Excel file…",
      ]);
      try {
        const data = await postJson("/api/edit/apply", { sessionId });
        renderResult(resultEl, data);
        showStep(paneEdit, "done");
      } catch (err) {
        renderErrorResult(resultEl, err.message);
        showStep(paneEdit, "done");
      } finally {
        if (stopCycle) stopCycle();
      }
    });

    $("#edit-start-over").addEventListener("click", () => {
      selectedFile = null;
      sessionId = null;
      fileInput.value = "";
      fileCard.hidden = true;
      dropzone.hidden = false;
      instructionsEl.value = "";
      showStep(paneEdit, "upload");
    });
  })();

  // =============================================================== CREATE
  (function initCreate() {
    const instructionsEl = $("#create-instructions");
    const describeError = $("#create-describe-error");
    const processingStatus = $("#create-processing-status");
    const changeListEl = $("#create-change-list");
    const warningsEl = $("#create-warnings");
    const resultEl = $("#create-result");

    let sessionId = null;
    let stopCycle = null;

    $all(".example-chip", paneCreate).forEach((chip) => {
      chip.addEventListener("click", () => {
        instructionsEl.value = chip.textContent;
        instructionsEl.focus();
      });
    });

    $("#create-analyze").addEventListener("click", async () => {
      describeError.hidden = true;
      const instructions = instructionsEl.value.trim();
      if (!instructions) {
        describeError.hidden = false;
        describeError.textContent = "Please describe the workbook you want.";
        return;
      }

      showStep(paneCreate, "processing");
      stopCycle = cycleStatus(processingStatus, ["Planning your workbook…", "Designing structure…", "Choosing formulas…"]);

      try {
        const data = await postJson("/api/create/preview", { instructions });
        sessionId = data.sessionId;
        renderChangeList(changeListEl, data.summary);
        if (data.warnings && data.warnings.length) {
          warningsEl.hidden = false;
          warningsEl.textContent = data.warnings.join("\n");
        } else {
          warningsEl.hidden = true;
        }
        showStep(paneCreate, "review");
      } catch (err) {
        showStep(paneCreate, "describe");
        describeError.hidden = false;
        describeError.textContent = err.message;
      } finally {
        if (stopCycle) stopCycle();
      }
    });

    $("#create-back-to-describe").addEventListener("click", () => showStep(paneCreate, "describe"));

    $("#create-confirm").addEventListener("click", async () => {
      showStep(paneCreate, "processing");
      stopCycle = cycleStatus(processingStatus, ["Building workbook…", "Verifying workbook…", "Preparing your Excel file…"]);
      try {
        const data = await postJson("/api/create/apply", { sessionId });
        renderResult(resultEl, data);
        showStep(paneCreate, "done");
      } catch (err) {
        renderErrorResult(resultEl, err.message);
        showStep(paneCreate, "done");
      } finally {
        if (stopCycle) stopCycle();
      }
    });

    $("#create-start-over").addEventListener("click", () => {
      sessionId = null;
      instructionsEl.value = "";
      showStep(paneCreate, "describe");
    });
  })();
})();
