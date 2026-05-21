(() => {
  "use strict";

  const $ = (id) => document.getElementById(id);

  // UI refs
  const logEl = $("log");
  const btnGerar = $("btnGerar");

  const excelInput = $("excelFile");
  const excelName = $("excelName");

  const outName = $("outName");
  const btnOutPick = $("btnOutPick");
  const outFolderFallback = $("outFolderFallback");

  const previewRows = $("previewRows");
  const excelSummary = $("excelSummary");

  const expectedChips = $("expectedChips");
  const missingNote = $("missingNote");

  // Status badges
  const dotExcel = $("dotExcel");
  const excelStatusText = $("excelStatusText");
  const dotOut = $("dotOut");
  const outStatusText = $("outStatusText");
  const dotBackend = $("dotBackend");
  const backendStatusText = $("backendStatusText");

  // Estado
  let excelFile = null;     // browser File (preview)
  let excelPath = null;     // app (pywebview) path real
  let outDirLabel = null;   // label
  let outDirPath = null;    // app (pywebview) path real


  // Colunas esperadas (mapeamento do backend) :contentReference[oaicite:11]{index=11}
  const EXPECTED_COLUMNS = [
    "NUMERO_PROCESSO","AUTOR","CUMPRIMENTO_SENTENCA","SITUACAO_PROCESSO",
    "DATA_ACAO","DATA_PERICIA","DATA_REALIZADA","DATA_LAUDO","TIPO LAUDO",
    "DATA_SENTENCA","SENTENCA","DATA_APELACAO","APE","DATA_JULGAMENTO","JULGA",
    "DATA_TRANSITO","DATA_CUMPRIMENTO","DATA_HOMOLOGACAO","DATA_PRECA","DATA_RPV",
    "DATA_OFICIO","DATA_OR_PAGAMENTO","DATA_ENCERRAMENTO"
  ];

  function log(msg) {
    const ts = new Date().toLocaleTimeString("pt-BR", { hour12: false });
    logEl.textContent += `[${ts}] ${msg}\n`;
    logEl.scrollTop = logEl.scrollHeight;
  }

  function setDot(dotEl, state) {
    dotEl.classList.remove("ok","warn","err");
    dotEl.classList.add(state);
  }

  function setExcelStatus(state, text) {
    setDot(dotExcel, state);
    excelStatusText.textContent = text;
  }

  function setOutStatus(state, text) {
    setDot(dotOut, state);
    outStatusText.textContent = text;
  }

  function setBackendStatus(state, text) {
    setDot(dotBackend, state);
    backendStatusText.textContent = text;
  }

  function humanFileName(file) {
    if (!file) return "Nenhum arquivo selecionado";
    return `${file.name} (${Math.round(file.size / 1024)} KB)`;
  }

  function escapeHtml(s) {
    return String(s ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  }

  function renderExpectedChips(foundHeaders = []) {
    const foundSet = new Set((foundHeaders || []).map(h => String(h).trim().toUpperCase()));
    expectedChips.innerHTML = "";

    const missing = [];
    for (const col of EXPECTED_COLUMNS) {
      const chip = document.createElement("span");
      chip.className = "chip";
      chip.textContent = col;

      if (!foundSet.has(col.toUpperCase())) {
        chip.classList.add("missing");
        missing.push(col);
      }

      expectedChips.appendChild(chip);
    }

    if (!foundHeaders.length) {
      missingNote.textContent = "Selecione uma planilha para validar as colunas.";
      return;
    }

    if (missing.length) {
      missingNote.textContent = `Atenção: ${missing.length} coluna(s) não encontrada(s) no cabeçalho (marcadas em vermelho).`;
    } else {
      missingNote.textContent = "OK: todas as colunas esperadas foram encontradas no cabeçalho.";
    }
  }

  function canEnableGenerate() {
  const hasExcel = !!excelFile || !!excelPath;
  const hasOut = !!outDirLabel || !!outDirPath;
  return hasExcel && hasOut;
}


  function updateGenerateButton() {
    btnGerar.disabled = !canEnableGenerate();
  }

  async function readExcelAndPreview(file) {
    if (typeof XLSX === "undefined" || !XLSX?.read) {
      setExcelStatus("err", "Excel: biblioteca XLSX ausente");
      log("Erro: biblioteca XLSX não carregou (SheetJS).");
      return;
    }

    const buf = await file.arrayBuffer();
    const wb = XLSX.read(buf, { type: "array" });
    const first = wb.SheetNames[0];
    if (!first) throw new Error("Planilha sem abas.");

    const ws = wb.Sheets[first];
    const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, raw: true, defval: "" });

    // Heurística simples: primeira linha = cabeçalho
    const header = (aoa[0] || []).map(v => String(v ?? "").trim());
    const rows = aoa.slice(1).filter(r => (r || []).some(c => String(c ?? "").trim() !== ""));

    // Render validação
    renderExpectedChips(header);

    // Contagem (espelha “Planilha carregada com X processos”) :contentReference[oaicite:12]{index=12}
    excelSummary.textContent = `Aba: "${first}". Linhas de dados detectadas: ${rows.length}. (Prévia: até 10 linhas)`;

    // Indexa colunas úteis
    const idx = (name) => header.findIndex(h => h.trim().toUpperCase() === name.toUpperCase());

    const iProc = idx("NUMERO_PROCESSO");
    const iAutor = idx("AUTOR");
    const iSit = idx("SITUACAO_PROCESSO");
    const iEnc = idx("DATA_ENCERRAMENTO");

    // Prévia
    previewRows.innerHTML = "";
    const take = Math.min(10, rows.length);

    if (take === 0) {
      previewRows.innerHTML = `<tr><td colspan="5" class="muted" style="padding: 14px 12px;">Sem linhas de dados.</td></tr>`;
      return;
    }

    for (let i = 0; i < take; i++) {
      const r = rows[i] || [];
      const procVal = (iProc >= 0 ? r[iProc] : "") || "";
      // Backend tem fallback _001 etc se não existir NUMERO_PROCESSO :contentReference[oaicite:13]{index=13}
      const obs = iProc < 0 ? "NUMERO_PROCESSO ausente (backend usa fallback)" : "";

      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td>${escapeHtml(procVal)}</td>
        <td>${escapeHtml(iAutor >= 0 ? r[iAutor] : "")}</td>
        <td>${escapeHtml(iSit >= 0 ? r[iSit] : "")}</td>
        <td>${escapeHtml(iEnc >= 0 ? r[iEnc] : "")}</td>
        <td class="muted">${escapeHtml(obs)}</td>
      `;
      previewRows.appendChild(tr);
    }
  }

  // Seleção de pasta (File System Access API, com fallback)
  async function pickOutputFolder() {
  // pywebview: pega path real
  if (window.pywebview && window.pywebview.api && window.pywebview.api.relconf_pick_out_dir) {
    const res = await window.pywebview.api.relconf_pick_out_dir();
    if (!res || res.canceled) return;
    if (res.ok === false) { log(res.error || "Falha ao selecionar pasta."); return; }

    outDirPath = res.path;
    outDirLabel = res.name || "Pasta selecionada";
    outName.textContent = outDirLabel;

    setOutStatus("ok", `Saída: ${outDirLabel}`);
    log(`Pasta de saída selecionada (pywebview): ${outDirPath}`);
    updateGenerateButton();
    return;
  }

  // browser: mantém seu fluxo atual
  if ("showDirectoryPicker" in window) {
    try {
      const handle = await window.showDirectoryPicker();
      outDirLabel = handle?.name || "Pasta selecionada";
      outName.textContent = outDirLabel;
      setOutStatus("ok", `Saída: ${outDirLabel}`);
      log(`Pasta de saída selecionada: ${outDirLabel}`);
      updateGenerateButton();
      return;
    } catch (e) {
      log("Seleção de pasta cancelada (DirectoryPicker).");
    }
  }

  outFolderFallback.click();
}


  function inferFolderNameFromFiles(fileList) {
    const files = Array.from(fileList || []);
    const f = files.find(x => (x?.webkitRelativePath || "").includes("/"));
    if (!f) return "Pasta selecionada";
    const rel = f.webkitRelativePath;
    return rel.split("/")[0] || "Pasta selecionada";
  }

  async function gerarReal() {
  if (!(window.pywebview && window.pywebview.api && window.pywebview.api.relconf_gerar)) {
    setBackendStatus("err", "Backend: API não disponível");
    log("Erro: relconf_gerar não encontrado.");
    return;
  }
  if (!excelPath || !outDirPath) {
    log("Selecione o Excel e a pasta de saída (diálogo do Python).");
    return;
  }

  setBackendStatus("ok", "Backend: conectado");
  const res = await window.pywebview.api.relconf_gerar({ excel_path: excelPath, out_dir: outDirPath });

  if (!res || res.ok === false) {
    setBackendStatus("err", "Backend: erro");
    log(res?.error || "Falha ao gerar.");
    return;
  }

  log(`OK! Saída: ${res.out_dir}`);
  log(`Processos: ${res.total_processos} | Gerados: ${res.arquivos_gerados} | Erros: ${res.erros}`);
}



  const excelPickLabel = document.querySelector("label[for='excelFile']");

  async function pickExcelPywebview(e) {
    if (!(window.pywebview && window.pywebview.api && window.pywebview.api.relconf_pick_excel)) return;

    e.preventDefault(); // impede abrir input file no app
    const res = await window.pywebview.api.relconf_pick_excel();
    if (!res || res.canceled) return;
    if (res.ok === false) { log(res.error || "Falha ao selecionar Excel."); return; }

    excelPath = res.path;
    excelFile = null; // no app não usamos File
    excelName.textContent = `${res.name} (${Math.round(res.size / 1024)} KB)`;
    setExcelStatus("ok", `Excel: ${res.name}`);
    log(`Excel selecionado (pywebview): ${res.path}`);

    // Se quiser manter a prévia no app sem SheetJS: dá pra fazer outro método python pra preview.
    // Por enquanto, só habilita o fluxo.
    excelSummary.textContent = "Excel selecionado. (Prévia no app pode depender do XLSX local)";
    renderExpectedChips([]); // ou mantenha como está
    updateGenerateButton();
  }

  if (excelPickLabel) {
    excelPickLabel.addEventListener("click", pickExcelPywebview);
  }


  // Bindings
  function bind() {
    // Estado inicial
    setExcelStatus("warn", "Excel: não selecionado");
    setOutStatus("warn", "Saída: não selecionada");
    setBackendStatus("warn", "Backend: não conectado");
    renderExpectedChips([]);

    excelInput.addEventListener("change", async (e) => {
      const f = e.target.files && e.target.files[0] ? e.target.files[0] : null;
      excelFile = f;
      excelName.textContent = humanFileName(f);

      previewRows.innerHTML = `<tr><td colspan="5" class="muted" style="padding: 14px 12px;">Carregando...</td></tr>`;

      if (!f) {
        setExcelStatus("warn", "Excel: não selecionado");
        excelSummary.textContent = "Selecione um Excel para ver contagem de linhas e uma amostra.";
        previewRows.innerHTML = `<tr><td colspan="5" class="muted" style="padding: 14px 12px;">Sem dados.</td></tr>`;
        renderExpectedChips([]);
        updateGenerateButton();
        return;
      }

      try {
        setExcelStatus("ok", `Excel: ${f.name}`);
        log(`Excel selecionado: ${f.name}`);

        await readExcelAndPreview(f);
        updateGenerateButton();
      } catch (err) {
        setExcelStatus("err", "Excel: erro ao ler");
        excelSummary.textContent = "Falha ao ler o Excel. Verifique o arquivo.";
        previewRows.innerHTML = `<tr><td colspan="5" class="muted" style="padding: 14px 12px;">Erro ao ler.</td></tr>`;
        log(`Erro ao ler Excel: ${err.message}`);
        renderExpectedChips([]);
        updateGenerateButton();
      } finally {
        // permite selecionar o mesmo arquivo novamente, se necessário
        e.target.value = "";
      }
    });

    btnOutPick.addEventListener("click", pickOutputFolder);

    outFolderFallback.addEventListener("change", (e) => {
      const files = e.target.files || [];
      if (!files.length) {
        log("Seleção de pasta cancelada (fallback).");
        return;
      }
      outDirLabel = inferFolderNameFromFiles(files);
      outName.textContent = outDirLabel;
      setOutStatus("ok", `Saída: ${outDirLabel}`);
      log(`Pasta de saída selecionada (fallback): ${outDirLabel}`);
      updateGenerateButton();
      e.target.value = "";
    });

    btnGerar.addEventListener("click", () => {
      if (window.pywebview && window.pywebview.api) gerarReal();
      else gerarStub(); // mantém no browser
    });


    log("Pronto. Selecione o Excel e a pasta de saída.");
  }

  document.addEventListener("DOMContentLoaded", bind);

;

})();
