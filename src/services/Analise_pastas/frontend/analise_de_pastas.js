(function () {
  "use strict";

  // ===== Helpers =====
  const $ = (sel) => document.querySelector(sel);
  const $$ = (sel) => Array.from(document.querySelectorAll(sel));

  function nowStamp() {
    const d = new Date();
    const pad = (n) => String(n).padStart(2, "0");
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}:${pad(d.getSeconds())}`;
  }

  function setStatus(text, kind) {
    const el = $("#logStatus");
    el.textContent = text;

    el.classList.remove("status-ok", "status-err");
    if (kind === "ok") el.classList.add("status-ok");
    if (kind === "err") el.classList.add("status-err");
  }

  function logLine(msg) {
    const box = $("#logBox");
    box.value += `[${nowStamp()}] ${msg}\n`;
    box.scrollTop = box.scrollHeight;
  }

  function clearLog() {
    $("#logBox").value = "";
    setStatus("Log limpo", "ok");
  }

  function warn(msg) {
    setStatus("Atenção", "err");
    logLine(`ERRO: ${msg}`);
    alert(msg);
  }

  function hasPyWebView() {
    return !!(window.pywebview && window.pywebview.api);
  }

  function setDisabled(el, disabled) {
    el.disabled = !!disabled;
    el.style.opacity = disabled ? "0.6" : "1";
  }

  // ===== State =====
  const state = {
    // Browser fallback
    folderFiles: [],
    folderName: "",

    // PyWebView mode
    folderPath: "",
    subfolderCount: null,

    // Preencher
    templateMode: "PADRAO", // PADRAO | ROBO
    belinePreencher: false,
    belineXlsxPath: "",
    belineXlsxLabel: "",
    belineXlsxFile: null, // browser fallback

    // Renomear
    belineRenomear: false,
    renomearXlsxPath: "",
    renomearXlsxLabel: "",
    renomearXlsxFile: null, // browser fallback
  };

  // ===== Views =====
  function showView(viewName) {
    $$(".view").forEach((v) => v.classList.add("hidden"));
    $(`#view-${viewName}`).classList.remove("hidden");
    $$(".navbtn").forEach((b) => b.classList.toggle("active", b.dataset.view === viewName));
  }

  // ===== Folder selection =====
  function extractTopFolderName(fileList) {
    const first = fileList && fileList[0];
    if (!first) return "";
    const rel = first.webkitRelativePath || "";
    if (!rel) return "";
    return rel.split("/")[0] || "";
  }

  function updateFolderUI() {
    const badge = $("#folderBadge");
    const input = $("#folderPath");

    // PyWebView
    if (state.folderPath) {
      input.value = state.folderPath;
      const qtd = typeof state.subfolderCount === "number" ? state.subfolderCount : null;
      badge.textContent = qtd === null ? `${state.folderPath}` : `${state.folderPath} • ${qtd} subpasta(s)`;
      return;
    }

    // Browser fallback
    if (!state.folderFiles.length) {
      badge.textContent = "Nenhuma pasta selecionada";
      input.value = "";
      return;
    }

    const name = state.folderName || "Pasta selecionada";
    badge.textContent = `${name} • ${state.folderFiles.length} arquivo(s)`;
    input.value = name;
  }

  async function pickFolder() {
    if (hasPyWebView()) {
      setStatus("Abrindo seletor…", null);
      try {
        const res = await window.pywebview.api.pick_folder();
        if (!res || res.canceled) {
          setStatus("Seleção cancelada", "err");
          return;
        }

        state.folderPath = res.folderPath || "";
        state.folderName = res.name || "";
        state.subfolderCount = typeof res.subfolderCount === "number" ? res.subfolderCount : null;

        // limpa modo browser
        state.folderFiles = [];
        $("#folderPicker").value = "";

        updateFolderUI();
        setStatus("Pasta selecionada", "ok");
        logLine(`Pasta base selecionada: ${state.folderPath} • ${state.subfolderCount ?? "?"} subpasta(s)`);
      } catch (e) {
        warn(String(e));
      }
      return;
    }

    $("#folderPicker").click();
  }

  // ===== Preencher controls =====
  function applyTemplateMode(mode) {
    state.templateMode = String(mode || "PADRAO").toUpperCase();
    updatePreencherControls();
  }

  function updatePreencherControls() {
    const chk = $("#chkBelinePreencher");
    const pathInput = $("#belineXlsxPath");
    const btnPick = $("#btnPickBelineXlsx");

    // label do modo (apenas visual)
    if (state.templateMode === "PADRAO") {
      // DS Beline desabilitado no Padrão (como no seu print)
      state.belinePreencher = false;
      chk.checked = false;
      setDisabled(chk, true);

      state.belineXlsxPath = "";
      state.belineXlsxFile = null;
      pathInput.value = "";
      pathInput.disabled = true;
      setDisabled(btnPick, true);
      return;
    }

    // ROBO
    setDisabled(chk, false);
    state.belinePreencher = !!chk.checked;

    if (state.belinePreencher) {
      pathInput.disabled = false;
      setDisabled(btnPick, false);
    } else {
      state.belineXlsxPath = "";
      state.belineXlsxFile = null;
      pathInput.value = "";
      pathInput.disabled = true;
      setDisabled(btnPick, true);
    }
  }

  async function pickBelineXlsx() {
    if (!state.belinePreencher) return;

    // PyWebView: reaproveita pick_template (já existe no seu backend) para selecionar .xlsx
    if (hasPyWebView()) {
      setStatus("Abrindo seletor…", null);
      try {
        const res = await window.pywebview.api.pick_template();
        if (!res || res.canceled) {
          setStatus("Seleção cancelada", "err");
          return;
        }

        state.belineXlsxPath = res.templatePath || "";
        state.belineXlsxFile = null;
        $("#belineXlsxPicker").value = "";

        state.belineXlsxLabel = res.name || state.belineXlsxPath.split(/[\\/]/).pop() || "";
        $("#belineXlsxPath").value = state.belineXlsxLabel;

        setStatus("Arquivo selecionado", "ok");
        logLine(`DS Beline (.xlsx) selecionado: ${state.belineXlsxPath}`);
      } catch (e) {
        warn(String(e));
      }
      return;
    }

    // Browser fallback
    $("#belineXlsxPicker").click();
  }

  // ===== Renomear controls =====
  function updateRenomearControls() {
    const beline = $("#chkBelineRenomear").checked;
    state.belineRenomear = beline;

    const ccsEl = $("#renameCCS");
    const parceiroEl = $("#renameParceiro");

    const xlsxInput = $("#renomearXlsxPath");
    const xlsxBtn = $("#btnPickRenomearXlsx");

    if (beline) {
      ccsEl.disabled = true;
      parceiroEl.disabled = true;
      xlsxInput.disabled = false;
      setDisabled(xlsxBtn, false);
    } else {
      ccsEl.disabled = false;
      parceiroEl.disabled = false;

      // limpa o xlsx quando sai do modo beline
      state.renomearXlsxPath = "";
      state.renomearXlsxFile = null;
      xlsxInput.value = "";
      xlsxInput.disabled = true;
      setDisabled(xlsxBtn, true);
    }
  }

  async function pickRenomearXlsx() {
    if (!state.belineRenomear) return;

    if (hasPyWebView()) {
      setStatus("Abrindo seletor…", null);
      try {
        const res = await window.pywebview.api.pick_template(); // reaproveita
        if (!res || res.canceled) {
          setStatus("Seleção cancelada", "err");
          return;
        }

        state.renomearXlsxPath = res.templatePath || "";
        state.renomearXlsxFile = null;
        $("#renomearXlsxPicker").value = "";

        state.renomearXlsxLabel = res.name || state.renomearXlsxPath.split(/[\\/]/).pop() || "";
        $("#renomearXlsxPath").value = state.renomearXlsxLabel;

        setStatus("Arquivo selecionado", "ok");
        logLine(`Renomear (DS Beline) .xlsx selecionado: ${state.renomearXlsxPath}`);
      } catch (e) {
        warn(String(e));
      }
      return;
    }

    $("#renomearXlsxPicker").click();
  }

  // ===== Corrigir controls =====
  function updateCorrigirControls(changed) {
    const padronizar = $("#chkFixPadronizarJudicial");
    const dashToDot = $("#chkFixDashToDot");
    const dotToDash = $("#chkFixDotToDash");
    const partner = $("#fixPartnerName");

    partner.disabled = !padronizar.checked;
    partner.style.opacity = padronizar.checked ? "1" : "0.6";

    if (changed === "dashToDot" && dashToDot.checked) {
      dotToDash.checked = false;
    }
    if (changed === "dotToDash" && dotToDash.checked) {
      dashToDot.checked = false;
    }
  }

  // ===== Validations =====
  function ensureFolderSelected() {
    if (state.folderPath) return true;
    if (state.folderFiles.length) return true;
    warn("Selecione uma pasta base primeiro.");
    return false;
  }

  function ensureBelineXlsxSelected(kind /* 'preencher' | 'renomear' */) {
    if (!hasPyWebView()) {
      // browser fallback
      const f = (kind === "preencher") ? state.belineXlsxFile : state.renomearXlsxFile;
      if (!f) return false;
      return (String(f.name || "").toLowerCase().endsWith(".xlsx"));
    }

    const p = (kind === "preencher") ? state.belineXlsxPath : state.renomearXlsxPath;
    if (!p) return false;
    return String(p).toLowerCase().endsWith(".xlsx");
  }

  // ===== Actions =====
  async function runPreencher() {
    if (!ensureFolderSelected()) return;

    if (state.templateMode === "ROBO" && state.belinePreencher) {
      if (!ensureBelineXlsxSelected("preencher")) {
        warn("Selecione o arquivo (.xlsx) do DS Beline.");
        return;
      }
    }

    if (hasPyWebView()) {
      if (!state.folderPath) {
        warn("No app, selecione a pasta base pelo botão 'Selecionar...' (seletor do Windows).");
        return;
      }

      setStatus("Executando…", null);
      logLine("Iniciando: Gerar Excel preenchido…");
      logLine(`Pasta base: ${state.folderPath}`);
      logLine(`Modelo: ${state.templateMode}`);
      logLine(`DS Beline: ${state.belinePreencher ? "SIM" : "NÃO"}${state.belinePreencher ? ` (${state.belineXlsxPath})` : ""}`);

      try {
        const res = await window.pywebview.api.run_preencher({
          folderPath: state.folderPath,
          templateMode: state.templateMode,
          templatePath: "",

          dsBeline: state.belinePreencher,

          // ✅ ESTE é o nome que o Python espera:
          baseXlsxPath: state.belineXlsxPath,

          // (opcional) pode manter também, não atrapalha:
          belineXlsxPath: state.belineXlsxPath,
        });


        if (!res) {
          warn("Resposta vazia do backend.");
          return;
        }
        if (res.canceled) {
          setStatus("Cancelado", "err");
          logLine("Operação cancelada pelo usuário.");
          return;
        }
        if (!res.ok) {
          warn(res.error || "Falha ao gerar Excel.");
          return;
        }

      const linhasOk = (res.linhasOk ?? res.linhas_ok ?? 0);
      const linhasProblemas = (res.linhasProblemas ?? res.linhas_problemas ?? 0);

    logLine(`OK. Arquivo gerado: ${res.outPath}`);
    logLine(`Linhas OK: ${linhasOk}`);
    logLine(`Linhas com avisos/problemas: ${linhasProblemas}`);
    logLine("Veja a aba 'VALIDAÇÃO' no Excel.");
    setStatus("Concluído", "ok");

alert(`Concluído.\nOK: ${linhasOk}\nProblemas: ${linhasProblemas}\n\nArquivo:\n${res.outPath}`);

      } catch (e) {
        warn(String(e));
      }
      return;
    }

    // Browser fallback: simulação
    setStatus("Executando…", null);
    logLine("Iniciando: Gerar Excel preenchido…");
    logLine(`Pasta base: ${state.folderName || "(selecionada via navegador)"}`);
    logLine(`Modelo: ${state.templateMode}`);
    logLine(`DS Beline: ${state.belinePreencher ? "SIM" : "NÃO"}`);
    await new Promise((r) => setTimeout(r, 450));
    const totalFiles = state.folderFiles.length;
    const ok = Math.max(0, Math.floor(totalFiles * 0.85));
    const problemas = Math.max(0, totalFiles - ok);
    logLine("OK. (Simulação) Arquivo gerado: preenchido_YYYYMMDD_HHMMSS.xlsx");
    logLine(`Linhas OK: ${ok}`);
    logLine(`Linhas com avisos/problemas: ${problemas}`);
    logLine("Veja a aba 'VALIDAÇÃO'. (Simulação)");
    setStatus("Concluído", "ok");
    alert("Simulação (no navegador). No app (pywebview) gera de verdade.");
  }

  async function runRenomearCore({ somenteCCS }) {
    if (!ensureFolderSelected()) return;

    const dryrun = $("#renameDryrun").checked;

    // DS BELINE
    if (state.belineRenomear) {
      if (!ensureBelineXlsxSelected("renomear")) {
        warn("Selecione o arquivo (.xlsx) para o modo DS Beline.");
        return;
      }

      if (!hasPyWebView()) {
        warn("Modo DS Beline não roda no navegador (somente no app/pywebview).");
        return;
      }

      setStatus("Executando…", null);
      logLine("Iniciando: Adicionar Numeração (DS Beline)…");
      logLine(`Pasta base: ${state.folderPath}`);
      logLine(`Arquivo (.xlsx): ${state.renomearXlsxPath}`);
      logLine(`Modo: ${dryrun ? "Prévia (não renomear)" : "Executar renomeação"}`);

      // IMPORTANTE: este método ainda precisa existir no backend.
      try {
        const api = window.pywebview.api;
       const res = await window.pywebview.api.run_renomear({
          folderPath: state.folderPath,
          dsBeline: true,
          xlsxPath: state.renomearXlsxPath,
          dryRun: dryrun,
          somenteCcs: !!somenteCCS,

        });


        if (!res) {
          warn("Resposta vazia do backend.");
          return;
        }
        if (!res.ok) {
          warn(res.error || "Falha ao renomear (DS Beline).");
          return;
        }

        (res.logs || []).forEach((l) => logLine(l));
        logLine(`Resumo: renomeadas=${res.renamed}, puladas=${res.skipped}, erros=${res.errors}`);
        setStatus("Concluído", res.errors ? "err" : "ok");
      } catch (e) {
        warn(String(e));
      }
      return;
    }

    // NORMAL (CCS / Parceiro)
    const rawCCS = $("#renameCCS").value.trim();
    const rawParceiro = $("#renameParceiro").value.trim();

    if (!rawCCS && !rawParceiro) {
      warn("Preencha Número CCS ou Número Parceiro.");
      return;
    }

    const ccs = rawCCS ? Number(rawCCS) : null;
    const parceiro = somenteCCS ? null : (rawParceiro ? Number(rawParceiro) : null);

    if ((ccs !== null && (!Number.isInteger(ccs) || ccs < 0)) ||
        (parceiro !== null && (!Number.isInteger(parceiro) || parceiro < 0))) {
      warn("Número CCS ou Número Parceiro inválido.");
      return;
    }

    if (hasPyWebView()) {
      if (!state.folderPath) {
        warn("No app, selecione a pasta base pelo botão 'Selecionar...' (seletor do Windows).");
        return;
      }

      setStatus("Executando…", null);
      logLine("Iniciando: Adicionar Numeração…");
      logLine(`Pasta base: ${state.folderPath}`);
      logLine(`Número CCS inicial: ${ccs === null ? "(vazio)" : ccs}`);
      logLine(`Número Parceiro inicial: ${parceiro === null ? "(vazio)" : parceiro}`);
      logLine(`Modo: ${dryrun ? "Prévia (não renomear)" : "Executar renomeação"}`);

      try {
        const res = await window.pywebview.api.run_renomear({
          folderPath: state.folderPath,
          ccsStart: ccs,
          parceiroStart: parceiro,
          dryRun: dryrun,

          // ✅ necessário pro backend:
          somenteCcs: !!somenteCCS,
        });


        if (!res) {
          warn("Resposta vazia do backend.");
          return;
        }
        if (!res.ok) {
          warn(res.error || "Falha ao renomear.");
          return;
        }

        (res.logs || []).forEach((l) => logLine(l));
        logLine(`Resumo: renomeadas=${res.renamed}, puladas=${res.skipped}, erros=${res.errors}`);
        setStatus("Concluído", res.errors ? "err" : "ok");
      } catch (e) {
        warn(String(e));
      }
      return;
    }

    // Browser fallback: simulação
    setStatus("Executando…", null);
    logLine("Iniciando: Adicionar Numeração…");
    logLine(`Pasta base: ${state.folderName || "(selecionada via navegador)"}`);
    logLine(`Número CCS inicial: ${ccs === null ? "(vazio)" : ccs}`);
    logLine(`Número Parceiro inicial: ${parceiro === null ? "(vazio)" : parceiro}`);
    logLine(`Modo: ${dryrun ? "Prévia (não renomear)" : "Executar renomeação"}`);
    await new Promise((r) => setTimeout(r, 450));
    const renamed = dryrun ? 0 : Math.min(25, state.folderFiles.length);
    const skipped = Math.max(0, Math.min(10, state.folderFiles.length - renamed));
    const errors = 0;
    logLine(`(Simulação) Resumo: renomeadas=${renamed}, puladas=${skipped}, erros=${errors}`);
    setStatus("Concluído", "ok");
  }

  async function runRenomear() {
    return runRenomearCore({ somenteCCS: false });
  }

  async function runSomenteCCS() {
    return runRenomearCore({ somenteCCS: true });
  }

  async function runLimparNumeracao() {
    if (!ensureFolderSelected()) return;

    const dryrun = $("#renameDryrun").checked;

    if (hasPyWebView()) {
      if (!state.folderPath) {
        warn("No app, selecione a pasta base pelo botão 'Selecionar...' (seletor do Windows).");
        return;
      }

      setStatus("Executando…", null);
      logLine("Iniciando: Limpar Numeração…");
      logLine(`Pasta base: ${state.folderPath}`);
      logLine(`Modo: ${dryrun ? "Prévia (não renomear)" : "Executar renomeação"}`);

      try {
        const res = await window.pywebview.api.run_renomear({
          folderPath: state.folderPath,
          dryRun: dryrun,
          limparNumeracao: true,
        });

        if (!res) {
          warn("Resposta vazia do backend.");
          return;
        }
        if (!res.ok) {
          warn(res.error || "Falha ao limpar numeração.");
          return;
        }

        (res.logs || []).forEach((l) => logLine(l));
        logLine(`Resumo: renomeadas=${res.renamed}, puladas=${res.skipped}, erros=${res.errors}`);
        setStatus("Concluído", res.errors ? "err" : "ok");
      } catch (e) {
        warn(String(e));
      }
      return;
    }

    setStatus("Executando…", null);
    logLine("Iniciando: Limpar Numeração…");
    logLine(`Pasta base: ${state.folderName || "(selecionada via navegador)"}`);
    logLine(`Modo: ${dryrun ? "Prévia (não renomear)" : "Executar renomeação"}`);
    await new Promise((r) => setTimeout(r, 450));
    const renamed = dryrun ? 0 : Math.min(25, state.folderFiles.length);
    logLine(`(Simulação) Resumo: renomeadas=${renamed}, puladas=0, erros=0`);
    setStatus("Concluído", "ok");
  }

  async function runCorrigir() {
    if (!ensureFolderSelected()) return;

    const raw = $("#fixPrefix").value.trim();
    const padronizarJudicial = $("#chkFixPadronizarJudicial").checked;
    const parceiroJudicial = $("#fixPartnerName").value.trim();
    const trocarHifenPorPonto = $("#chkFixDashToDot").checked;
    const trocarPontoPorHifen = $("#chkFixDotToDash").checked;
    const usarModoEspecial = padronizarJudicial || trocarHifenPorPonto || trocarPontoPorHifen;

    if (!usarModoEspecial && raw === ".") {
      warn("O prefixo não pode ser apenas '.'");
      return;
    }
    if (padronizarJudicial && !parceiroJudicial) {
      warn("Informe o parceiro para padronizar o nome judicial.");
      return;
    }
    if (trocarHifenPorPonto && trocarPontoPorHifen) {
      warn("Escolha apenas uma conversão: '-' por '.' ou '.' por '-'.");
      return;
    }

    const dryrun = $("#fixDryrun").checked;

    if (hasPyWebView()) {
      if (!state.folderPath) {
        warn("No app, selecione a pasta base pelo botão 'Selecionar...' (seletor do Windows).");
        return;
      }

      setStatus("Executando…", null);
      logLine("Iniciando: Corrigir numeração…");
      logLine(`Pasta base: ${state.folderPath}`);
      if (padronizarJudicial) {
        logLine(`Modo: Padronizar parceiro judicial (${parceiroJudicial})`);
      } else if (trocarHifenPorPonto) {
        logLine("Modo: Trocar '-' por '.' no prefixo numérico");
      } else if (trocarPontoPorHifen) {
        logLine("Modo: Trocar '.' por '-' no prefixo numérico");
      } else {
        logLine(`Novo prefixo (antes do primeiro '.'): ${raw === "" ? "(vazio = remover prefixo)" : raw}`);
      }
      logLine(`Modo: ${dryrun ? "Prévia (não renomear)" : "Executar renomeação"}`);

      try {
        const res = await window.pywebview.api.run_corrigir({
          folderPath: state.folderPath,
          prefix: raw,
          dryRun: dryrun,
          padronizarJudicial,
          parceiroJudicial,
          trocarHifenPorPonto,
          trocarPontoPorHifen,
        });

        if (!res) {
          warn("Resposta vazia do backend.");
          return;
        }
        if (!res.ok) {
          warn(res.error || "Falha ao corrigir numeração.");
          return;
        }

        (res.logs || []).forEach((l) => logLine(l));
        logLine(`Resumo: renomeadas=${res.renamed}, puladas=${res.skipped}, erros=${res.errors}`);
        setStatus("Concluído", res.errors ? "err" : "ok");
      } catch (e) {
        warn(String(e));
      }
      return;
    }

    // Browser fallback: simulação
    setStatus("Executando…", null);
    logLine("Iniciando: Corrigir numeração…");
    logLine(`Pasta base: ${state.folderName || "(selecionada via navegador)"}`);
    if (padronizarJudicial) {
      logLine(`Modo: Padronizar parceiro judicial (${parceiroJudicial})`);
    } else if (trocarHifenPorPonto) {
      logLine("Modo: Trocar '-' por '.' no prefixo numérico");
    } else if (trocarPontoPorHifen) {
      logLine("Modo: Trocar '.' por '-' no prefixo numérico");
    } else {
      logLine(`Novo prefixo (antes do primeiro '.'): ${raw === "" ? "(vazio = remover prefixo)" : raw}`);
    }
    logLine(`Modo: ${dryrun ? "Prévia (não renomear)" : "Executar renomeação"}`);
    await new Promise((r) => setTimeout(r, 450));
    const renamed = dryrun ? 0 : Math.min(25, state.folderFiles.length);
    const skipped = Math.max(0, Math.min(10, state.folderFiles.length - renamed));
    const errors = 0;
    logLine(`(Simulação) Resumo: renomeadas=${renamed}, puladas=${skipped}, erros=${errors}`);
    setStatus("Concluído", "ok");
  }

  // ===== Wire up UI =====
  function init() {
    // Navigation
    $$(".navbtn").forEach((btn) => {
      btn.addEventListener("click", () => showView(btn.dataset.view));
    });

    // Folder picker
    $("#btnPickFolder").addEventListener("click", pickFolder);
    $("#folderPicker").addEventListener("change", (e) => {
      const files = Array.from(e.target.files || []);
      state.folderFiles = files;
      state.folderName = extractTopFolderName(files);

      // limpa modo pywebview
      state.folderPath = "";
      state.subfolderCount = null;

      updateFolderUI();

      if (files.length) {
        setStatus("Pasta selecionada", "ok");
        logLine(`Pasta base selecionada: ${state.folderName || "(sem nome detectável)"} • ${files.length} arquivo(s)`);
      } else {
        setStatus("Nenhuma pasta", "err");
      }
    });

    // Template mode radios
    $$('input[name="templateMode"]').forEach((r) => {
      r.addEventListener("change", () => applyTemplateMode(r.value));
    });

    // DS Beline (Preencher)
    $("#chkBelinePreencher").addEventListener("change", () => updatePreencherControls());
    $("#btnPickBelineXlsx").addEventListener("click", pickBelineXlsx);
    $("#belineXlsxPicker").addEventListener("change", (e) => {
      const f = (e.target.files && e.target.files[0]) ? e.target.files[0] : null;
      state.belineXlsxFile = f;
      state.belineXlsxPath = "";
      state.belineXlsxLabel = f ? f.name : "";
      $("#belineXlsxPath").value = state.belineXlsxLabel;
      if (f) logLine(`DS Beline (.xlsx) selecionado (browser): ${f.name}`);
    });

    // DS Beline (Renomear)
    $("#chkBelineRenomear").addEventListener("change", updateRenomearControls);
    $("#btnPickRenomearXlsx").addEventListener("click", pickRenomearXlsx);
    $("#renomearXlsxPicker").addEventListener("change", (e) => {
      const f = (e.target.files && e.target.files[0]) ? e.target.files[0] : null;
      state.renomearXlsxFile = f;
      state.renomearXlsxPath = "";
      state.renomearXlsxLabel = f ? f.name : "";
      $("#renomearXlsxPath").value = state.renomearXlsxLabel;
      if (f) logLine(`Renomear (DS Beline) .xlsx selecionado (browser): ${f.name}`);
    });

    // Run buttons
    $("#btnRunPreencher").addEventListener("click", runPreencher);
    $("#btnRunRenomear").addEventListener("click", runRenomear);
    $("#btnRunSomenteCCS").addEventListener("click", runSomenteCCS);
    $("#btnLimparNumeracao").addEventListener("click", runLimparNumeracao);
    $("#btnRunCorrigir").addEventListener("click", runCorrigir);

    $("#chkFixPadronizarJudicial").addEventListener("change", () => updateCorrigirControls("padronizar"));
    $("#chkFixDashToDot").addEventListener("change", () => updateCorrigirControls("dashToDot"));
    $("#chkFixDotToDash").addEventListener("change", () => updateCorrigirControls("dotToDash"));

    // Clear log buttons
    $("#btnClearLogGlobal").addEventListener("click", clearLog);
    $("#btnClearLog1").addEventListener("click", clearLog);
    $("#btnClearLog2").addEventListener("click", clearLog);
    $("#btnClearLog3").addEventListener("click", clearLog);

    // Initial state
    applyTemplateMode("PADRAO");
    $("#chkBelineRenomear").checked = false;
    updateRenomearControls();
    updateCorrigirControls();
    updateFolderUI();
    showView("preencher");
    setStatus("Pronto", "ok");
    logLine("Interface carregada.");
    logLine(hasPyWebView() ? "Modo app detectado (pywebview): backend conectado." : "Modo navegador: ações ficam em simulação.");
  }

  // Start
  document.addEventListener("DOMContentLoaded", () => {
    init();
  });

})();
