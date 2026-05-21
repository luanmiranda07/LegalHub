// src/services/IA_LAUDO/frontend/ialaudos.js
(() => {
  "use strict";

  const el = {
    btnHelp: document.getElementById("btnHelp"),
    helpDialog: document.getElementById("helpDialog"),
    btnCloseHelp: document.getElementById("btnCloseHelp"),

    btnSelect: document.getElementById("btnSelect"),
    fileInput: document.getElementById("fileInput"),
    fileLabel: document.getElementById("fileLabel"),

    dropzone: document.getElementById("dropzone"),

    btnAsk: document.getElementById("btnAsk"),
    btnCancel: document.getElementById("btnCancel"),

    progressFill: document.getElementById("progressFill"),
    progressTxt: document.getElementById("progressTxt"),

    answerBox: document.getElementById("answerBox"),

    btnCopy: document.getElementById("btnCopy"),
    btnClear: document.getElementById("btnClear"),
  };

  const state = {
    // Browser: File (não tem path real pro Python)
    file: /** @type {File|null} */ (null),

    // App (pywebview): path real do PDF selecionado pelo diálogo do Python
    pdfPath: "",
    pdfName: "",

    running: false,
    timer: /** @type {number|null} */ (null),
    cancelToken: 0,
  };

  function hasPywebview() {
    return !!(window.pywebview && window.pywebview.api);
  }

  function setStatusText(text) {
    el.answerBox.value = text || "";
  }

  function setProgress(p) {
    const v = Math.max(0, Math.min(100, Number(p) || 0));
    el.progressFill.style.width = `${v}%`;
    el.progressTxt.textContent = `${Math.round(v)}%`;
  }

  function updateFileLabel() {
    if (state.pdfName) {
      el.fileLabel.textContent = state.pdfName;
      return;
    }
    if (state.file) {
      el.fileLabel.textContent = state.file.name;
      return;
    }
    el.fileLabel.textContent = "Nenhum PDF selecionado";
  }

  function canAsk() {
    // No app: precisa de pdfPath (caminho real)
    if (hasPywebview()) return !!state.pdfPath;
    // No browser: aceita File (simulação)
    return !!state.file;
  }

  function lockUI(lock) {
    el.btnSelect.disabled = lock;
    el.btnAsk.disabled = lock || !canAsk();
    el.btnCancel.disabled = !lock;
    el.btnCopy.disabled = lock || !(el.answerBox.value || "").trim();
  }

  function clearTimer() {
    if (state.timer) {
      clearInterval(state.timer);
      state.timer = null;
    }
  }

  async function pickPdf() {
    if (state.running) return;

    // Preferência: no app (pywebview), pega o path real via Python
    if (hasPywebview() && typeof window.pywebview.api.ialaudos_pick_pdf === "function") {
      const res = await window.pywebview.api.ialaudos_pick_pdf();
      if (!res || res.canceled) return;

      if (res.ok === false) {
        alert(res.error || "Falha ao selecionar PDF.");
        return;
      }

      state.pdfPath = String(res.path || "");
      state.pdfName = String(res.name || "");
      state.file = null;

      updateFileLabel();
      el.btnAsk.disabled = !canAsk();
      return;
    }

    // Browser fallback: input file (não dá path real pro Python)
    el.fileInput.click();
  }

  function setBrowserFile(file) {
    state.pdfPath = "";
    state.pdfName = "";
    state.file = null;

    if (!file) {
      updateFileLabel();
      el.btnAsk.disabled = true;
      return;
    }

    const isPdf =
      (file.type || "").toLowerCase() === "application/pdf" ||
      file.name.toLowerCase().endsWith(".pdf");

    if (!isPdf) {
      alert("Selecione um arquivo PDF.");
      updateFileLabel();
      el.btnAsk.disabled = true;
      return;
    }

    state.file = file;
    updateFileLabel();
    el.btnAsk.disabled = !canAsk();
  }

  async function askAI() {
    if (state.running) return;

    // No app, só funciona se tiver pdfPath real
    if (hasPywebview()) {
      if (!state.pdfPath) {
        setStatusText('No app, selecione o PDF pelo botão "Selecionar PDF..." (diálogo do Windows).');
        return;
      }
      if (typeof window.pywebview.api.ialaudos_analisar !== "function") {
        setStatusText("Backend não disponível: método ialaudos_analisar não encontrado no pywebview.api.");
        return;
      }
    } else {
      // Browser: apenas simulação
      if (!state.file) {
        setStatusText("Selecione um PDF primeiro.");
        return;
      }
    }

    state.running = true;
    lockUI(true);
    setStatusText("Processando...");
    setProgress(0);

    const myToken = ++state.cancelToken;

    // progresso leve enquanto espera
    let fake = 0;
    clearTimer();
    state.timer = window.setInterval(() => {
      if (state.cancelToken !== myToken) return;
      fake = Math.min(fake + 2, 35);
      setProgress(fake);
    }, 120);

    try {
      if (!hasPywebview()) {
        // simulação no browser
        await new Promise((r) => setTimeout(r, 1200));
        if (state.cancelToken !== myToken) return;

        clearTimer();
        setProgress(100);
        setStatusText(
          `[SIMULAÇÃO - NAVEGADOR]\nArquivo: ${state.file.name}\n\nPara funcionar de verdade, rode dentro do Legal-Hub (pywebview).`
        );
        el.btnCopy.disabled = false;
        return;
      }

      const res = await window.pywebview.api.ialaudos_analisar({ pdf_path: state.pdfPath });
      if (state.cancelToken !== myToken) return;

      clearTimer();
      setProgress(100);

      if (!res || res.ok === false) {
        setStatusText(res?.error || "Erro ao processar o PDF.");
        el.btnCopy.disabled = true;
        return;
      }

      setStatusText(res.text || "(sem conteúdo)");
      el.btnCopy.disabled = !(res.text || "").trim();
    } catch (e) {
      if (state.cancelToken !== myToken) return;
      clearTimer();
      setProgress(0);
      setStatusText(String(e?.message || e));
      el.btnCopy.disabled = true;
    } finally {
      if (state.cancelToken === myToken) {
        state.running = false;
        lockUI(false);
      }
    }
  }

  function cancel() {
    if (!state.running) return;
    state.cancelToken++;
    clearTimer();
    state.running = false;
    setProgress(0);
    setStatusText("Cancelado.");
    lockUI(false);
  }

  function clearAll() {
    if (state.running) cancel();
    state.file = null;
    state.pdfPath = "";
    state.pdfName = "";
    updateFileLabel();
    setProgress(0);
    setStatusText("");
    el.btnCopy.disabled = true;
    el.btnAsk.disabled = !canAsk();
  }

  async function copyAnswer() {
    const txt = (el.answerBox.value || "").trim();
    if (!txt) return;

    try {
      await navigator.clipboard.writeText(txt);
    } catch {
      const ta = document.createElement("textarea");
      ta.value = txt;
      document.body.appendChild(ta);
      ta.select();
      document.execCommand("copy");
      ta.remove();
    }
  }

  // ---- bind ----
  function bind() {
    // ajuda (se existir)
    if (el.btnHelp && el.helpDialog) el.btnHelp.addEventListener("click", () => el.helpDialog.showModal());
    if (el.btnCloseHelp && el.helpDialog) el.btnCloseHelp.addEventListener("click", () => el.helpDialog.close());

    el.btnSelect.addEventListener("click", pickPdf);
    el.dropzone.addEventListener("click", pickPdf);

    el.fileInput.addEventListener("change", () => {
      const f = el.fileInput.files && el.fileInput.files[0] ? el.fileInput.files[0] : null;
      setBrowserFile(f);
      el.fileInput.value = "";
    });

    // drag/drop: no app vira “atalho” pro seletor; no browser aceita file
    el.dropzone.addEventListener("dragover", (ev) => {
      ev.preventDefault();
      el.dropzone.classList.add("dragover");
    });
    el.dropzone.addEventListener("dragleave", () => el.dropzone.classList.remove("dragover"));
    el.dropzone.addEventListener("drop", (ev) => {
      ev.preventDefault();
      el.dropzone.classList.remove("dragover");

      if (hasPywebview()) {
        // no app: forçar seletor do Python para obter path real
        pickPdf();
        return;
      }

      const f = ev.dataTransfer && ev.dataTransfer.files && ev.dataTransfer.files[0] ? ev.dataTransfer.files[0] : null;
      setBrowserFile(f);
    });

    el.btnAsk.addEventListener("click", askAI);
    el.btnCancel.addEventListener("click", cancel);
    el.btnClear.addEventListener("click", clearAll);
    el.btnCopy.addEventListener("click", copyAnswer);

    // init
    updateFileLabel();
    setProgress(0);
    setStatusText("");
    el.btnAsk.disabled = !canAsk();
    el.btnCancel.disabled = true;
    el.btnCopy.disabled = true;
  }

  bind();
})();
