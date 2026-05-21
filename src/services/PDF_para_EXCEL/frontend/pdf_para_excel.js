(() => {
  "use strict";

  const $ = (id) => document.getElementById(id);

  const elRows = $("rows");
  const elStatus = $("statusText");
  const elDotApp = $("dotApp");
  const elCount = $("fileCount");
  const elInclude = $("includeSubfolders");

  const btnAbrirMultiplos = $("btnAbrirMultiplos");
  const btnAbrirPasta = $("btnAbrirPasta");

  // Fallback web (se abrir no navegador sem pywebview)
  const elPdfFiles = $("pdfFiles");
  const elPdfFolder = $("pdfFolder");

  // Estado: no DESKTOP a gente guarda paths reais em row.path
  let rows = [];
  let nextId = 1;

  // Último resultado do backend (para salvar Excel via Python)
  let lastBackendRows = null;

  function hasBackend() {
    return !!(window.pywebview && window.pywebview.api);
  }

  function setStatus(type, msg) {
    elDotApp.classList.remove("ok", "warn", "err");
    elDotApp.classList.add(type);
    elStatus.textContent = msg;
  }

  function updateCount() {
    elCount.textContent = `${rows.length} arquivo(s)`;
  }

  function escapeHtml(s) {
    return String(s ?? "")
      .replaceAll("&", "&amp;")
      .replaceAll("<", "&lt;")
      .replaceAll(">", "&gt;")
      .replaceAll('"', "&quot;")
      .replaceAll("'", "&#039;");
  }

  function filenameOnly(pathOrName) {
    const s = String(pathOrName || "");
    const parts = s.split(/[\\/]/g);
    return parts[parts.length - 1] || s;
  }

  function render() {
    elRows.innerHTML = "";

    if (!rows.length) {
      const tr = document.createElement("tr");
      tr.innerHTML = `
        <td colspan="6" style="padding: 14px 12px; color: rgba(255,255,255,0.72);">
          Nenhum arquivo selecionado.
        </td>
      `;
      elRows.appendChild(tr);
      return;
    }

    for (const r of rows) {
      const tr = document.createElement("tr");
      if (r.selected) tr.classList.add("selected");

      tr.innerHTML = `
        <td title="${escapeHtml(r.path || r.file?.name || "")}">${escapeHtml(filenameOnly(r.path || r.file?.name || ""))}</td>
        <td>${escapeHtml(r.nome)}</td>
        <td>${escapeHtml(r.processo)}</td>
        <td>${escapeHtml(r.dataEnc)}</td>
        <td>${escapeHtml(r.total)}</td>
        <td><span class="pill">${escapeHtml(r.status)}</span></td>
      `;

      tr.addEventListener("click", () => {
        r.selected = !r.selected;
        render();
      });

      elRows.appendChild(tr);
    }
  }

  function addPaths(paths) {
    const list = Array.from(paths || []).filter(Boolean);
    const pdfs = list.filter(p => String(p).toLowerCase().endsWith(".pdf"));
    if (!pdfs.length) {
      setStatus("warn", "Nenhum PDF encontrado.");
      return;
    }

    const existing = new Set(rows.map(r => String(r.path || "").toLowerCase()));
    let added = 0;

    for (const p of pdfs) {
      const key = String(p).toLowerCase();
      if (existing.has(key)) continue;

      rows.push({
        id: nextId++,
        path: String(p),      // <-- no desktop isso é essencial
        file: null,
        nome: "",
        processo: "",
        dataEnc: "",
        total: "",
        status: "Pendente",
        selected: false
      });

      existing.add(key);
      added++;
    }

    updateCount();
    render();
    setStatus("ok", `${added} arquivo(s) adicionado(s). Total: ${rows.length}.`);
  }

  // -----------------------------
  // AÇÕES DESKTOP (pywebview)
  // -----------------------------

  async function abrirMultiplosBackend() {
    if (!hasBackend()) {
      setStatus("warn", "Backend não disponível (abrindo no navegador?).");
      return;
    }
    const api = window.pywebview.api;
    const picked = await api.pdfexcel_pick_files();
    if (!picked || picked.canceled) {
      setStatus("warn", "Seleção cancelada.");
      return;
    }
    addPaths(picked.files || []);
  }

  async function abrirPastaBackend() {
    if (!hasBackend()) {
      setStatus("warn", "Backend não disponível (abrindo no navegador?).");
      return;
    }

    const api = window.pywebview.api;

    // 1) escolher a pasta via dialog do backend
    const picked = await api.pick_folder();
    if (!picked || picked.canceled) {
      setStatus("warn", "Seleção cancelada.");
      return;
    }

    // 2) listar PDFs da pasta via backend
    const folderPath = picked.folderPath;
    const recursive = !!elInclude.checked;

    const listed = await api.pdfexcel_list_folder({ folder_path: folderPath, recursive });
    if (!listed?.ok) {
      setStatus("err", listed?.error || "Falha ao listar PDFs da pasta.");
      return;
    }

    addPaths(listed.files || []);
  }

  async function processar() {
    if (!rows.length) {
      setStatus("warn", "Selecione arquivos ou uma pasta primeiro.");
      return;
    }
    if (!hasBackend()) {
      setStatus("err", "Sem backend. Rode dentro do app (pywebview).");
      return;
    }

    const api = window.pywebview.api;

    // Python precisa de caminhos reais
    const pdfPaths = rows.map(r => r.path).filter(Boolean);
    if (pdfPaths.length !== rows.length) {
      setStatus("warn", "Alguns itens não têm path real. Use Abrir/Abrir pasta (backend).");
      return;
    }

    for (const r of rows) r.status = "Processando...";
    render();
    setStatus("warn", "Processando PDFs...");

    const result = await api.pdfexcel_process({ pdf_paths: pdfPaths });
    if (!result?.ok) {
      setStatus("err", result?.error || "Erro desconhecido no backend.");
      for (const r of rows) r.status = "Erro";
      render();
      return;
    }

    lastBackendRows = result.rows || [];

    // Atualiza por índice (mesma ordem do input)
    for (let i = 0; i < rows.length; i++) {
      const r = rows[i];
      const br = lastBackendRows[i] || {};

      r.nome = br["Nome"] || "";
      r.processo = br["Número do processo"] || "";
      r.dataEnc = br["Data Encerramento"] || "";
      r.total = br["Total"] || "";

      const err = br["Erro"] || "";
      r.status = err ? `Erro: ${err}` : "OK";
    }

    render();
    const s = result.summary || {};
    setStatus("ok", `Concluído. Total: ${s.total ?? rows.length} | OK: ${s.ok ?? "?"} | Erros: ${s.errors ?? "?"}`);
  }

  async function salvarExcel() {
    if (!rows.length) {
      setStatus("warn", "Nada para salvar.");
      return;
    }
    if (!hasBackend()) {
      setStatus("err", "Sem backend para salvar Excel. Rode dentro do app.");
      return;
    }
    if (!lastBackendRows || !lastBackendRows.length) {
      setStatus("warn", "Clique em Processar antes de salvar.");
      return;
    }

    const api = window.pywebview.api;
    const picked = await api.pdfexcel_pick_save_xlsx();
    if (!picked || picked.canceled) {
      setStatus("warn", "Salvar cancelado.");
      return;
    }

    const outPath = picked.path;
    const r = await api.pdfexcel_save_xlsx({ output_path: outPath, rows: lastBackendRows });
    if (!r?.ok) {
      setStatus("err", r?.error || "Falha ao salvar Excel.");
      return;
    }

    setStatus("ok", `Excel salvo em: ${outPath}`);
  }

  function removerSelecao() {
    const before = rows.length;
    rows = rows.filter(r => !r.selected);
    updateCount();
    render();
    setStatus("ok", `Removidos ${before - rows.length} arquivo(s).`);
  }

  function limpar() {
    rows = [];
    lastBackendRows = null;
    updateCount();
    render();
    setStatus("ok", "Pronto.");
  }

  // -----------------------------
  // EVENTOS
  // -----------------------------
  // Botões do HTML (agora funcionam)
  if (btnAbrirMultiplos) btnAbrirMultiplos.addEventListener("click", abrirMultiplosBackend);
  if (btnAbrirPasta) btnAbrirPasta.addEventListener("click", abrirPastaBackend);

  $("btnProcessar").addEventListener("click", processar);
  $("btnSalvarExcel").addEventListener("click", salvarExcel);
  $("btnLimpar").addEventListener("click", limpar);
  $("btnRemoverSelecao").addEventListener("click", removerSelecao);

  // Fallback navegador (se alguém abrir fora do app)
  if (elPdfFiles) {
    elPdfFiles.addEventListener("change", (e) => {
      // no navegador isso não vira path real; só pra UI
      const files = Array.from(e.target.files || []);
      for (const f of files) {
        rows.push({
          id: nextId++,
          path: "",
          file: f,
          nome: "",
          processo: "",
          dataEnc: "",
          total: "",
          status: "Pendente",
          selected: false
        });
      }
      updateCount(); render();
      e.target.value = "";
    });
  }
  if (elPdfFolder) {
    elPdfFolder.addEventListener("change", (e) => {
      const files = Array.from(e.target.files || []);
      for (const f of files) {
        rows.push({
          id: nextId++,
          path: "",
          file: f,
          nome: "",
          processo: "",
          dataEnc: "",
          total: "",
          status: "Pendente",
          selected: false
        });
      }
      updateCount(); render();
      e.target.value = "";
    });
  }

  // Init
  updateCount();
  render();
  setStatus("ok", hasBackend() ? "Pronto (backend detectado)." : "Pronto (sem backend).");
})();
