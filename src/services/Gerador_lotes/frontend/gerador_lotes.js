/* gerador_lotes.js (pywebview)
   - Seleciona o LOTE via diálogo nativo (pywebview)
   - Usa o backend Python para ler/gravar Excel (pandas/openpyxl)
   - Chama:
       - window.pywebview.api.lotes_default_model()
       - window.pywebview.api.pick_lote_file()
       - window.pywebview.api.pick_model_file()   (opcional)
       - window.pywebview.api.pick_folder()       (pasta de saída opcional)
       - window.pywebview.api.lotes_preview_columns({caminho_lote,...}) (opcional)
       - window.pywebview.api.run_lotes(payload)
*/

(() => {
  "use strict";

  // ----------------------------
  // Helpers de UI
  // ----------------------------
  const $ = (id) => document.getElementById(id);

  const logEl = $("log");
  function log(msg) {
    const ts = new Date().toLocaleTimeString("pt-BR", { hour12: false });
    logEl.textContent += `[${ts}] ${msg}\n`;
    logEl.scrollTop = logEl.scrollHeight;
  }

  function setBadge(which, state, text) {
    // state: "ok" | "warn" | "err"
    const dot = which === "model" ? $("dotModel") : $("dotLote");
    const txt = which === "model" ? $("modelStatusText") : $("loteStatusText");

    dot.classList.remove("ok", "warn", "err");
    dot.classList.add(state);
    txt.textContent = text;
  }

  function setBox(id, text) {
    const el = $(id);
    if (el) el.textContent = text;
  }

  function hasPyWebView() {
    return !!(window.pywebview && window.pywebview.api);
  }

  // ----------------------------
  // Estado
  // ----------------------------
  let lotePath = null;   // string
  let modelPath = null;  // string (opcional)
  let outputDir = null;  // string (opcional)

  // ----------------------------
  // Modelo (backend)
  // ----------------------------
  async function loadDefaultModel() {
    if (!hasPyWebView()) {
      setBadge("model", "err", "Modelo: rode via app (pywebview)");
      log("Erro: window.pywebview.api não encontrado. Abra esta página pelo app.");
      return;
    }

    try {
      const resp = await window.pywebview.api.lotes_default_model();
      if (resp && resp.ok) {
        setBadge("model", "ok", `Modelo: ${resp.name}`);
        const hint = $("modelHint");
        if (hint) {
          hint.innerHTML =
            `Usando modelo padrão do app: <strong>${resp.name}</strong><br>` +
            `<span style="opacity:.75">(${resp.path})</span><br>` +
            `Se quiser, você pode selecionar um modelo manualmente abaixo.`;
        }
        log(`Modelo padrão OK: ${resp.path}`);
      } else {
        setBadge("model", "warn", "Modelo: não encontrado (ver src/docs)");
        log(`Aviso: modelo padrão não disponível. Motivo: ${resp?.error || "desconhecido"}`);
      }
    } catch (e) {
      setBadge("model", "err", "Modelo: erro ao consultar backend");
      log(`Erro: falha ao consultar modelo padrão no backend: ${e.message}`);
    }
  }

  // ----------------------------
  // Diálogos (backend)
  // ----------------------------
  async function pickLote() {
    if (!hasPyWebView()) return;

    try {
      const resp = await window.pywebview.api.pick_lote_file();
      if (resp?.canceled) {
        log("Seleção do LOTE cancelada.");
        return;
      }
      if (resp && resp.ok === false) {
        setBadge("lote", "err", "LOTE: arquivo inválido");
        log(`Erro ao selecionar LOTE: ${resp.error || "arquivo inválido"}`);
        return;
      }

      lotePath = resp.lotePath;
      setBox("loteName", resp.name || lotePath);
      setBadge("lote", "ok", `LOTE: ${resp.name || "selecionado"}`);
      log(`LOTE selecionado: ${lotePath}`);

      // (opcional) prévia de colunas, útil pra debug
      try {
        const preview = await window.pywebview.api.lotes_preview_columns({
          caminho_lote: lotePath,
          coluna_processo_hint: $("colProcesso")?.value || "Número do Processo",
        });
        if (preview?.ok) {
          log(`Colunas detectadas no LOTE: ${preview.columns.length}`);
        } else {
          log(`Aviso: não consegui listar colunas do LOTE. Motivo: ${preview?.error || "desconhecido"}`);
        }
      } catch (e) {
        log(`Aviso: falha ao pedir preview de colunas: ${e.message}`);
      }
    } catch (e) {
      setBadge("lote", "err", "LOTE: erro no diálogo");
      log(`Erro ao abrir diálogo do LOTE: ${e.message}`);
    }
  }

  async function pickModel() {
    if (!hasPyWebView()) return;

    try {
      const resp = await window.pywebview.api.pick_model_file();
      if (resp?.canceled) {
        log("Seleção do modelo cancelada.");
        return;
      }
      if (resp && resp.ok === false) {
        setBadge("model", "err", "Modelo: arquivo inválido");
        log(`Erro ao selecionar modelo: ${resp.error || "arquivo inválido"}`);
        return;
      }

      modelPath = resp.modelPath;
      setBox("modelName", resp.name || modelPath);
      setBadge("model", "ok", `Modelo: ${resp.name || "manual"}`);
      log(`Modelo manual selecionado: ${modelPath}`);
    } catch (e) {
      setBadge("model", "err", "Modelo: erro no diálogo");
      log(`Erro ao abrir diálogo do modelo: ${e.message}`);
    }
  }

  function clearModel() {
    modelPath = null;
    setBox("modelName", "Nenhum modelo selecionado");
    // volta a mostrar o default (se existir)
    loadDefaultModel().catch(() => {});
    log("Modelo manual removido (voltando ao padrão, se existir).");
  }

  async function pickOutDir() {
    if (!hasPyWebView()) return;

    try {
      const resp = await window.pywebview.api.pick_folder();
      if (resp?.canceled) {
        log("Seleção da pasta de saída cancelada.");
        return;
      }
      if (resp && resp.ok === false) {
        log(`Erro ao selecionar pasta: ${resp.error || "pasta inválida"}`);
        return;
      }

      outputDir = resp.folderPath;
      setBox("outDirName", outputDir);
      log(`Pasta de saída selecionada: ${outputDir}`);
    } catch (e) {
      log(`Erro ao abrir diálogo de pasta: ${e.message}`);
    }
  }

  function clearOutDir() {
    outputDir = null;
    setBox("outDirName", "Mesma pasta do LOTE");
    log("Pasta de saída removida (usando a mesma pasta do LOTE).");
  }

  // ----------------------------
  // Execução
  // ----------------------------
  async function gerarArquivos() {
    log("Iniciando geração (backend)...");

    if (!hasPyWebView()) {
      log("Erro: window.pywebview.api não encontrado. Abra esta página pelo app.");
      return;
    }

    if (!lotePath) {
      setBadge("lote", "err", "LOTE: selecione um arquivo");
      log("Erro: selecione o arquivo LOTE.");
      return;
    }

    const payload = {
      caminho_lote: lotePath,
      coluna_processo: ($("colProcesso")?.value || "Número do Processo").trim(),
      solicitado_por: ($("solicitadoPor")?.value || "45270").trim(),
      evento_map: {
        "CALCP": ($("colCalcp")?.value || "").trim(),
        "HC30%": ($("colHc30")?.value || "").trim(),
        "HCP": ($("colHcp")?.value || "").trim(),
        "CALCS": ($("colCalcs")?.value || "").trim(),
        "HSP": ($("colHsp")?.value || "").trim(),
      },
      modelo_path: modelPath,   // null => backend usa padrão
      output_dir: outputDir,    // null => backend usa pasta do lote
      // eventos: null,         // opcional: lista custom
    };

    try {
      const resp = await window.pywebview.api.run_lotes(payload);
      if (!resp?.ok) {
        setBadge("lote", "err", "Execução: erro");
        log(`Erro no backend: ${resp?.error || "desconhecido"}`);
        return;
      }

      setBadge("lote", "ok", "LOTE: processado");
      log(`OK. Pasta de saída: ${resp.output_dir}`);
      if (resp.warnings && resp.warnings.length) {
        log("Avisos:");
        for (const w of resp.warnings) log(` - ${w}`);
      }
      log("Arquivos gerados:");
      for (const f of resp.files || []) log(` - ${f}`);

      log("Concluído.");
    } catch (e) {
      setBadge("lote", "err", "Execução: erro inesperado");
      log(`Erro inesperado: ${e.message}`);
    }
  }

  // ----------------------------
  // Bind eventos
  // ----------------------------
  function bind() {
    $("btnPickLote")?.addEventListener("click", () => pickLote());
    $("btnPickModel")?.addEventListener("click", () => pickModel());
    $("btnClearModel")?.addEventListener("click", () => clearModel());
    $("btnPickOutDir")?.addEventListener("click", () => pickOutDir());
    $("btnClearOutDir")?.addEventListener("click", () => clearOutDir());

    $("btnGerar")?.addEventListener("click", () => {
      gerarArquivos().catch((err) => {
        log(`Erro inesperado: ${err.message}`);
        setBadge("lote", "err", "Execução: erro inesperado");
      });
    });
  }

  // ----------------------------
  // Init
  // ----------------------------
  document.addEventListener("DOMContentLoaded", async () => {
    bind();

    setBadge("model", "warn", "Modelo: aguardando backend");
    setBadge("lote", "warn", "LOTE: não selecionado");
    setBox("loteName", "Nenhum arquivo selecionado");
    setBox("modelName", "Nenhum modelo selecionado");
    setBox("outDirName", "Mesma pasta do LOTE");

    log("Pronto. Selecione o LOTE e clique em 'Gerar Arquivos'.");
    await loadDefaultModel();
  });
})();
