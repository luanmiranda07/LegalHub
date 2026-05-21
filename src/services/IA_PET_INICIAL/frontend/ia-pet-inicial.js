(() => {
  "use strict";

  const MAX_PDFS = 6;
  const VERSION = "1.2.14";

  const el = {
    versionLabel: document.getElementById("versionLabel"),
    helpVersion: document.getElementById("helpVersion"),

    btnSelect: document.getElementById("btnSelect"),
    fileInput: document.getElementById("fileInput"),
    fileSummary: document.getElementById("fileSummary"),
    fileList: document.getElementById("fileList"),
    btnRemove: document.getElementById("btnRemove"),

    dropzone: document.getElementById("dropzone"),

    btnRun: document.getElementById("btnRun"),
    btnCancel: document.getElementById("btnCancel"),
    progressFill: document.getElementById("progressFill"),
    progressText: document.getElementById("progressText"),
    progressBar: document.querySelector(".progress[role='progressbar']"),

    answerBox: document.getElementById("answerBox"),
    btnClear: document.getElementById("btnClear"),
    btnWord: document.getElementById("btnWord"),
    btnHelp: document.getElementById("btnHelp"),
    btnLog: document.getElementById("btnLog"),

    helpDialog: document.getElementById("helpDialog"),
    btnCloseHelp: document.getElementById("btnCloseHelp"),

    logDialog: document.getElementById("logDialog"),
    logBox: document.getElementById("logBox"),
    btnCloseLog: document.getElementById("btnCloseLog"),

    toast: document.getElementById("toast"),
  };

  const state = {
    files: [], // { id, key, file, selected }
    running: false,
    progressTimer: null,
    aiTimer: null,
    logText: "",
    runToken: 0,
  };

  // ---------- Helpers ----------
  function toast(msg) {
    el.toast.textContent = msg;
    el.toast.classList.add("show");
    window.clearTimeout(toast._t);
    toast._t = window.setTimeout(() => el.toast.classList.remove("show"), 3200);
  }

  function bytesToHuman(n) {
    const v0 = Number(n || 0);
    const units = ["B", "KB", "MB", "GB"];
    let v = v0;
    let i = 0;
    while (v >= 1024 && i < units.length - 1) {
      v = v / 1024;
      i++;
    }
    return `${v.toFixed(i === 0 ? 0 : 1)} ${units[i]}`;
  }

  function isPdf(file) {
    const name = (file?.name || "").toLowerCase();
    return file && (file.type === "application/pdf" || name.endsWith(".pdf"));
  }

  function hasPywebview() {
    return !!(window.pywebview && window.pywebview.api);
  }

  function fileKey(file) {
    // inclui path quando existir (app)
    const p = file?.path || "";
    return `${p || file.name}::${file.size || 0}::${file.lastModified || 0}`;
  }

  function randId() {
    if (window.crypto && crypto.getRandomValues) {
      const buf = new Uint32Array(2);
      crypto.getRandomValues(buf);
      return `${buf[0].toString(16)}${buf[1].toString(16)}`;
    }
    return `${Math.random().toString(16).slice(2)}${Date.now().toString(16)}`;
  }

  function setProgress(val) {
    const v = Math.max(0, Math.min(100, Number(val || 0)));
    el.progressFill.style.width = `${v}%`;
    el.progressText.textContent = `${v.toFixed(0)}%`;
    if (el.progressBar) el.progressBar.setAttribute("aria-valuenow", String(v));
  }

  function getProgress() {
    const w = (el.progressFill.style.width || "0%").replace("%", "");
    const n = Number(w);
    return Number.isFinite(n) ? n : 0;
  }

  function updateButtons() {
    const hasFiles = state.files.length > 0;
    const anySelected = state.files.some(f => f.selected);

    el.btnRemove.disabled = !anySelected || state.running;
    el.btnRun.disabled = !hasFiles || state.running;
    el.btnCancel.disabled = !state.running;

    el.btnWord.disabled = state.running || !state.logText;
    el.btnLog.disabled = !state.logText;
  }

  function renderFileList() {
    el.fileList.innerHTML = "";

    if (state.files.length === 0) {
      const li = document.createElement("li");
      li.className = "file-item";
      li.innerHTML =
        `<label class="file-row">
           <span class="file-name muted">Nenhum PDF selecionado.</span>
         </label>`;
      el.fileList.appendChild(li);
      return;
    }

    for (const item of state.files) {
      const f = item.file;

      const li = document.createElement("li");
      li.className = "file-item";

      const size = bytesToHuman(f.size || 0);
      const name = f.name || "(sem nome)";

      li.innerHTML =
        `<label class="file-row">
           <input type="checkbox" class="file-check" ${item.selected ? "checked" : ""} />
           <span class="file-name" title="${name}">${name}</span>
           <span class="file-size">${size}</span>
         </label>`;

      const chk = li.querySelector(".file-check");
      chk.addEventListener("change", () => {
        item.selected = chk.checked;
        updateButtons();
      });

      el.fileList.appendChild(li);
    }
  }

  function renderSummary() {
    const n = state.files.length;
    const total = state.files.reduce((acc, x) => acc + Number(x.file?.size || 0), 0);
    el.fileSummary.textContent = n === 0
      ? "Nenhum PDF selecionado"
      : `${n} PDF(s) • ${bytesToHuman(total)}`;
  }

  function renderAll() {
    renderFileList();
    renderSummary();
    updateButtons();
  }

  function addFiles(fileList) {
    const incoming = Array.from(fileList || []).filter(isPdf);
    if (incoming.length === 0) {
      toast("Envie apenas arquivos PDF.");
      return;
    }

    const existing = new Set(state.files.map(x => x.key));
    const newOnes = [];

    for (const f of incoming) {
      const key = fileKey(f);
      if (!existing.has(key)) {
        newOnes.push({ id: randId(), key, file: f, selected: false });
        existing.add(key);
      }
    }

    if (newOnes.length === 0) {
      toast("Todos os PDFs selecionados já estão na lista atual.");
      return;
    }

    const allowed = MAX_PDFS - state.files.length;
    if (allowed <= 0) {
      toast(`Você já selecionou o máximo de ${MAX_PDFS} PDFs.`);
      return;
    }

    if (newOnes.length > allowed) {
      newOnes.length = allowed;
      toast(`Só é possível selecionar até ${MAX_PDFS} PDFs. Os demais foram ignorados.`);
    }

    state.files.push(...newOnes);
    renderAll();
  }

  function removeSelected() {
    const before = state.files.length;
    state.files = state.files.filter(x => !x.selected);
    if (state.files.length === before) return;
    renderAll();
  }

  // ---------- “IA” (progresso) ----------
  function startProgressAnimation() {
    stopProgressAnimation();
    state.progressTimer = window.setInterval(() => {
      const val = getProgress();
      if (val < 95) setProgress(val + 1);
      else stopProgressAnimation();
    }, 1250);
  }

  function stopProgressAnimation() {
    if (state.progressTimer) {
      window.clearInterval(state.progressTimer);
      state.progressTimer = null;
    }
  }

  function extractJson(text) {
    const t = String(text || "").trim();
    const start = t.indexOf("{");
    const end = t.lastIndexOf("}");
    if (start >= 0 && end > start) return t.slice(start, end + 1);
    return t;
  }

  function finishRun(rawText) {
    if (!state.running) return;

    state.running = false;
    stopProgressAnimation();

    if (state.aiTimer) {
      window.clearTimeout(state.aiTimer);
      state.aiTimer = null;
    }

    setProgress(100);
    state.logText = rawText;

    try {
      const jsonStr = extractJson(rawText);
      const data = JSON.parse(jsonStr);

      const vinculo = data.VINCULO_COM_TRABALHO ?? "N/A";
      const cliente = data.NOME_CLIENTE ?? "N/A";
      const cidade = data.CIDADE ?? "N/A";
      const uf = data.UF ?? "N/A";

      el.answerBox.value =
        "Dados extraidos com sucesso!\n" +
        "Gere o Documento Word e realize a revisão.\n\n" +
        `VINCULO COM TRABALHO: ${vinculo}\n` +
        `CLIENTE: ${cliente}\n` +
        `CIDADE/UF: ${cidade}/${uf}`;
    } catch (e) {
      el.answerBox.value =
        "A IA respondeu, mas não foi possível formatar os dados automaticamente.\n" +
        "Verifique o botão LOG para ver a resposta original.\n\n" +
        "--- Resposta Original ---\n" +
        rawText;
    }

    updateButtons();
  }

  async function startRun() {
    if (state.running) return;
    if (state.files.length === 0) {
      toast("Selecione ou arraste pelo menos um arquivo PDF antes de executar.");
      return;
    }

    const token = ++state.runToken;

    state.running = true;
    state.logText = "";
    el.answerBox.value = "";
    setProgress(0);
    updateButtons();
    startProgressAnimation();

    try {
      if (hasPywebview() && typeof window.pywebview.api.iapet_analisar === "function") {
        const pdfPaths = state.files
          .map(x => x?.file?.path)
          .filter(p => typeof p === "string" && p.trim().length > 0);

        if (pdfPaths.length !== state.files.length) {
          throw new Error('No app, selecione os PDFs pelo botão "Selecionar PDF…" (diálogo do Windows) para obter o caminho real.');
        }

        toast("Analisando PDFs...");
        const res = await window.pywebview.api.iapet_analisar({ pdf_paths: pdfPaths });

        if (!state.running || state.runToken !== token) return;

        if (!res || res.ok === false) {
          finishRun(String(res?.error || "Erro ao consultar a IA."));
          toast("Falha ao executar.");
          return;
        }

        finishRun(String(res.text || ""));
        toast("Resposta recebida.");
        return;
      }

      // Browser fallback (simulação)
      const delay = 2200 + state.files.length * 500;
      await new Promise(r => setTimeout(r, delay));

      if (!state.running || state.runToken !== token) return;

      finishRun(simulatedRawResponse());
      toast("Resposta recebida (simulação).");
    } catch (e) {
      if (!state.running || state.runToken !== token) return;
      state.running = false;
      stopProgressAnimation();
      setProgress(0);
      el.answerBox.value = String(e?.message || e);
      state.logText = el.answerBox.value;
      updateButtons();
      toast("Falha ao executar.");
    }
  }

  function cancelRun() {
    if (!state.running) return;
    state.running = false;
    state.runToken++;
    stopProgressAnimation();
    setProgress(0);
    updateButtons();
    toast("Execução cancelada.");
  }

  function clearAll() {
    state.running = false;
    state.runToken++;
    stopProgressAnimation();
    state.files = [];
    state.logText = "";
    el.answerBox.value = "";
    setProgress(0);
    renderAll();
    toast("Interface limpa.");
  }

  // ---------- Dialogs ----------
  function openDialog(dlg) {
    if (dlg && typeof dlg.showModal === "function") dlg.showModal();
    else toast("Seu navegador não suporta dialog. Atualize ou use um navegador moderno.");
  }

  function closeDialog(dlg) {
    if (dlg && typeof dlg.close === "function") dlg.close();
  }

  function openHelp() {
    openDialog(el.helpDialog);
  }

  function openLog() {
    el.logBox.textContent = state.logText || "Sem conteúdo de LOG no momento.";
    openDialog(el.logDialog);
  }

  function safeFileName(s) {
  const txt = String(s || "").trim() || "Pet_inicial";
  return txt
    .replace(/[\\/:*?"<>|]/g, "")   // caracteres inválidos no Windows
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 80);
}

function tryGetClientNameFromLog(logText) {
  try {
    const t = String(logText || "");
    const start = t.indexOf("{");
    const end = t.lastIndexOf("}");
    const jsonStr = (start >= 0 && end > start) ? t.slice(start, end + 1) : t;
    const data = JSON.parse(jsonStr);
    return data?.NOME_CLIENTE || "";
  } catch {
    return "";
  }
}

async function generateWord() {
  if (state.running) return;

  if (!state.logText) {
    toast("Nenhuma resposta disponível. Execute primeiro.");
    return;
  }

  if (!hasPywebview() || typeof window.pywebview.api.iapet_pick_save_docx !== "function") {
    toast("Salvar Word disponível apenas no app (pywebview).");
    return;
  }

  try {
    const cliente = tryGetClientNameFromLog(state.logText);
    const suggested = `Pet_inicial_${safeFileName(cliente)}.docx`;

    // 1) pergunta onde salvar
    const pick = await window.pywebview.api.iapet_pick_save_docx({ suggested_name: suggested });
    if (!pick || pick.canceled) return;
    if (pick.ok === false) {
      toast(pick.error || "Falha ao escolher local.");
      return;
    }

    // 2) gera o Word diretamente no caminho escolhido
    toast("Gerando Word...");
    const res = await window.pywebview.api.iapet_gerar_word({
      raw_text: state.logText,
      save_path: pick.path,
    });

    if (!res || res.ok === false) {
      toast(res?.error || "Falha ao gerar Word.");
      return;
    }

    const path = res.path || pick.path || "";
    const name = res.name || pick.name || "Documento Word";

    state.logText = (state.logText || "") + `\n\n✅ Word gerado: ${path}`;
    el.answerBox.value = (el.answerBox.value || "") + `\n\n✅ Word gerado: ${name}\n${path}`;
    updateButtons();
    toast("Word gerado com sucesso.");
  } catch (e) {
    toast(String(e?.message || e));
  }
}


  let pyReady = false;
window.addEventListener("pywebviewready", () => { pyReady = true; });

function hasPywebview() {
  return pyReady && window.pywebview && window.pywebview.api;
}

  async function pickPdfs() {
  // Se estiver rodando dentro do app (pywebview), use o diálogo do Python
  if (window.pywebview && window.pywebview.api) {
    const api = window.pywebview.api;

    if (typeof api.iapet_pick_pdfs !== "function") {
      toast("Backend não disponível: iapet_pick_pdfs não encontrado.");
      return;
    }

    const res = await api.iapet_pick_pdfs();
    if (!res || res.canceled) return;

    if (res.ok === false) {
      toast(res.error || "Falha ao selecionar PDFs.");
      return;
    }

    const picked = Array.isArray(res.files) ? res.files : [];
    if (!picked.length) return;

    // transforma em objetos “File-like” que seu addFiles() já aceita
    const fileLikes = picked.map(x => ({
      name: String(x.name || ""),
      size: Number(x.size || 0),
      lastModified: Number(x.lastModified || 0),
      type: "application/pdf",
      path: String(x.path || ""),
    }));

    addFiles(fileLikes);
    return;
  }

  // Fallback: navegador
  el.fileInput.click();
}


  // Eventos:
el.btnSelect.addEventListener("click", pickPdfs);
el.dropzone.addEventListener("click", pickPdfs);



  el.fileInput.addEventListener("change", (ev) => {
    addFiles(ev.target.files);
    el.fileInput.value = "";
  });

  el.btnRemove.addEventListener("click", removeSelected);

el.dropzone.addEventListener("drop", (e) => {
  e.preventDefault();
  el.dropzone.classList.remove("dragover");

  // No app, o drag&drop não garante path real -> abre o seletor do Python
  if (window.pywebview && window.pywebview.api) {
    pickPdfs();
    return;
  }

  if (e.dataTransfer && e.dataTransfer.files) {
    addFiles(e.dataTransfer.files);
  }
});


  el.btnRun.addEventListener("click", startRun);
  el.btnCancel.addEventListener("click", cancelRun);
  el.btnClear.addEventListener("click", clearAll);

  el.btnHelp.addEventListener("click", openHelp);
  el.btnCloseHelp.addEventListener("click", () => closeDialog(el.helpDialog));

  el.btnLog.addEventListener("click", openLog);
  el.btnCloseLog.addEventListener("click", () => closeDialog(el.logDialog));

  el.btnWord.addEventListener("click", generateWord);

  // ---------- Init ----------
  if (el.versionLabel) el.versionLabel.textContent = `v${VERSION}`;
  if (el.helpVersion) el.helpVersion.textContent = `v${VERSION}`;
  setProgress(0);
  renderAll();
})();
