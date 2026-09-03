(() => {
  "use strict";

  const dropzone = document.getElementById("dropzone");
  const dropzoneBody = document.getElementById("dropzoneBody");
  const fileInput = document.getElementById("fileInput");

  const configPanel = document.getElementById("configPanel");
  const fileChip = document.getElementById("fileChip");
  const fileNameEl = document.getElementById("fileName");
  const clearFileBtn = document.getElementById("clearFile");

  const sheetRow = document.getElementById("sheetRow");
  const sheetSelect = document.getElementById("sheetSelect");
  const blockRow = document.getElementById("blockRow");
  const blockSelect = document.getElementById("blockSelect");
  const blockMeta = document.getElementById("blockMeta");

  const segmented = document.querySelectorAll(".segmented__option");
  const minSta = document.getElementById("minSta");
  const maxSta = document.getElementById("maxSta");
  const staHint = document.getElementById("staHint");

  const outboundChips = document.getElementById("outboundChips");
  const chipsClear = document.getElementById("chipsClear");

  const allBlocksCheck = document.getElementById("allBlocksCheck");
  const generateBtn = document.getElementById("generateBtn");
  const statusEl = document.getElementById("status");

  let currentFile = null;
  let workbookData = null;   // { sheets: [ { name, blocks: [...] } ] }
  let sortBy = "departure";
  let activeOutbound = new Set();

  // ---------- helpers ----------

  function setStatus(message, kind) {
    statusEl.textContent = message || "";
    statusEl.classList.remove("is-error", "is-success");
    if (kind) statusEl.classList.add(kind === "error" ? "is-error" : "is-success");
  }

  function resetConfig() {
    workbookData = null;
    activeOutbound = new Set();
    sortBy = "departure";
    segmented.forEach((btn) => {
      const active = btn.dataset.sort === "departure";
      btn.classList.toggle("is-active", active);
      btn.setAttribute("aria-checked", String(active));
    });
    minSta.value = "";
    maxSta.value = "";
    allBlocksCheck.checked = false;
    configPanel.hidden = true;
    setStatus("");
  }

  function currentSheet() {
    if (!workbookData) return null;
    return workbookData.sheets.find((s) => s.name === sheetSelect.value) || null;
  }

  function currentBlock() {
    const sheet = currentSheet();
    if (!sheet) return null;
    const titleRow = Number(blockSelect.value);
    return sheet.blocks.find((b) => b.title_row === titleRow) || sheet.blocks[0] || null;
  }

  function renderBlockOptions() {
    const sheet = currentSheet();
    if (!sheet) return;

    blockSelect.innerHTML = "";
    sheet.blocks.forEach((b) => {
      const opt = document.createElement("option");
      opt.value = String(b.title_row);
      opt.textContent = b.title;
      blockSelect.appendChild(opt);
    });
    blockRow.hidden = sheet.blocks.length < 2;

    renderBlockMeta();
  }

  function renderBlockMeta() {
    const block = currentBlock();
    if (!block) {
      blockMeta.textContent = "";
      return;
    }
    const range = block.sta_min && block.sta_max
      ? `STA ${block.sta_min}–${block.sta_max}`
      : "STA not detected";
    blockMeta.textContent = `${block.incoming_count} inbound flight(s) · ${range}`;
    staHint.textContent = block.sta_min && block.sta_max
      ? `Detected STA range for this block: ${block.sta_min}–${block.sta_max}. Leave blank to include all.`
      : "Leave blank to include every inbound flight in the block.";

    renderOutboundChips(block.outbound_codes || []);
  }

  function renderOutboundChips(codes) {
    activeOutbound = new Set();
    outboundChips.innerHTML = "";
    if (codes.length === 0) {
      const p = document.createElement("span");
      p.className = "field-hint";
      p.textContent = "No outbound destinations detected in this block.";
      outboundChips.appendChild(p);
      return;
    }
    codes.forEach((code) => {
      const chip = document.createElement("button");
      chip.type = "button";
      chip.className = "chip";
      chip.textContent = code;
      chip.setAttribute("aria-pressed", "false");
      chip.addEventListener("click", () => {
        if (activeOutbound.has(code)) {
          activeOutbound.delete(code);
          chip.classList.remove("is-active");
          chip.setAttribute("aria-pressed", "false");
        } else {
          activeOutbound.add(code);
          chip.classList.add("is-active");
          chip.setAttribute("aria-pressed", "true");
        }
      });
      outboundChips.appendChild(chip);
    });
  }

  function applyAllBlocksMode() {
    const all = allBlocksCheck.checked;
    sheetRow.hidden = all;
    blockRow.hidden = all || (currentSheet()?.blocks.length || 0) < 2;
    generateBtn.textContent = all
      ? "Generate every matrix block (.zip)"
      : "Generate modified matrix";
  }

  // ---------- file intake ----------

  async function handleFile(file) {
    if (!file) return;
    if (!/\.(xlsx|xlsm)$/i.test(file.name)) {
      setStatus("Please choose an .xlsx or .xlsm workbook.", "error");
      return;
    }

    currentFile = file;
    dropzone.classList.add("is-busy");
    dropzoneBody.querySelector(".dropzone__primary").textContent = "Reading workbook…";
    setStatus("");

    const formData = new FormData();
    formData.append("file", file);

    try {
      const res = await fetch("/api/inspect", { method: "POST", body: formData });
      const data = await res.json();
      if (!res.ok) throw new Error(data.error || "Could not read that workbook.");

      workbookData = data;
      fileNameEl.textContent = file.name;
      sheetSelect.innerHTML = "";
      data.sheets.forEach((s) => {
        const opt = document.createElement("option");
        opt.value = s.name;
        opt.textContent = s.name;
        sheetSelect.appendChild(opt);
      });
      sheetRow.hidden = data.sheets.length < 2;

      renderBlockOptions();
      applyAllBlocksMode();
      configPanel.hidden = false;
    } catch (err) {
      resetConfig();
      currentFile = null;
      setStatus(err.message, "error");
    } finally {
      dropzone.classList.remove("is-busy");
      dropzoneBody.querySelector(".dropzone__primary").textContent =
        "Drop a TRANSFER MATRIX workbook here, or click to browse";
    }
  }

  dropzone.addEventListener("click", () => fileInput.click());
  dropzone.addEventListener("keydown", (e) => {
    if (e.key === "Enter" || e.key === " ") {
      e.preventDefault();
      fileInput.click();
    }
  });
  fileInput.addEventListener("change", (e) => handleFile(e.target.files[0]));

  ["dragenter", "dragover"].forEach((evt) =>
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
  dropzone.addEventListener("drop", (e) => {
    const file = e.dataTransfer.files[0];
    handleFile(file);
  });

  clearFileBtn.addEventListener("click", () => {
    currentFile = null;
    fileInput.value = "";
    resetConfig();
  });

  // ---------- config interactions ----------

  sheetSelect.addEventListener("change", () => {
    renderBlockOptions();
    applyAllBlocksMode();
  });
  blockSelect.addEventListener("change", renderBlockMeta);

  segmented.forEach((btn) => {
    btn.addEventListener("click", () => {
      sortBy = btn.dataset.sort;
      segmented.forEach((b) => {
        const active = b === btn;
        b.classList.toggle("is-active", active);
        b.setAttribute("aria-checked", String(active));
      });
    });
  });

  chipsClear.addEventListener("click", () => {
    activeOutbound = new Set();
    outboundChips.querySelectorAll(".chip").forEach((c) => {
      c.classList.remove("is-active");
      c.setAttribute("aria-pressed", "false");
    });
  });

  allBlocksCheck.addEventListener("change", applyAllBlocksMode);

  // ---------- generate ----------

  generateBtn.addEventListener("click", async () => {
    if (!currentFile) {
      setStatus("Choose a workbook first.", "error");
      return;
    }

    const all = allBlocksCheck.checked;
    const block = currentBlock();

    const formData = new FormData();
    formData.append("file", currentFile);
    formData.append("mode", all ? "all" : "single");
    formData.append("sort_by", sortBy);
    if (minSta.value) formData.append("min_sta", minSta.value);
    if (maxSta.value) formData.append("max_sta", maxSta.value);
    if (activeOutbound.size > 0) {
      formData.append("outbound", Array.from(activeOutbound).join(","));
    }
    if (!all) {
      formData.append("sheet", sheetSelect.value);
      formData.append("title_row", block ? String(block.title_row) : "1");
    }

    generateBtn.disabled = true;
    generateBtn.textContent = all ? "Generating archive…" : "Generating…";
    setStatus("");

    try {
      const res = await fetch("/api/process", { method: "POST", body: formData });

      if (!res.ok) {
        const contentType = res.headers.get("content-type") || "";
        const data = contentType.includes("application/json") ? await res.json() : {};
        throw new Error(data.error || "Could not generate the modified matrix.");
      }

      const blob = await res.blob();
      const disposition = res.headers.get("content-disposition") || "";
      const match = disposition.match(/filename="?([^"]+)"?/);
      const downloadName = match ? match[1] : (all ? "modified_matrices.zip" : "modified_matrix.xlsx");

      const url = URL.createObjectURL(blob);
      const a = document.createElement("a");
      a.href = url;
      a.download = downloadName;
      document.body.appendChild(a);
      a.click();
      a.remove();
      URL.revokeObjectURL(url);

      setStatus(`Downloaded ${downloadName}.`, "success");
    } catch (err) {
      setStatus(err.message, "error");
    } finally {
      generateBtn.disabled = false;
      applyAllBlocksMode();
    }
  });
})();
