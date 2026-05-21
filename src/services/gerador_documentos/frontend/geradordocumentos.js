const elementos = {
  arquivoExcel: document.getElementById("arquivoExcel"),
  arquivoExcelTexto: document.getElementById("arquivoExcelTexto"),
  pastaSaidaTexto: document.getElementById("pastaSaidaTexto"),
  gerarPdf: document.getElementById("gerarPdf"),

  btnSelecionarArquivo: document.getElementById("btnSelecionarArquivo"),
  btnSelecionarPasta: document.getElementById("btnSelecionarPasta"),
  btnGerar: document.getElementById("btnGerar"),
  btnCancelar: document.getElementById("btnCancelar"),
  btnLimpar: document.getElementById("btnLimpar"),

  statusText: document.getElementById("statusText"),
  dotApp: document.getElementById("dotApp"),
  hintNote: document.getElementById("hintNote")
};

const estado = {
  arquivo: null,
  arquivoPath: "",
  pastaSaida: "",
  gerando: false,
  jobId: "",
  pollTimer: null,
  cancelamentoSolicitado: false,
  ultimaPastaFinal: ""
};

const extensoesPermitidas = [".xlsx", ".xls", ".xlsm", ".csv"];

function iniciar() {
  elementos.btnSelecionarArquivo.addEventListener("click", selecionarArquivo);
  elementos.arquivoExcel.addEventListener("change", carregarArquivo);

  elementos.btnSelecionarPasta.addEventListener("click", selecionarPasta);
  elementos.gerarPdf.addEventListener("change", atualizarStatusPorCampos);

  elementos.btnGerar.addEventListener("click", gerarDocumentos);
  elementos.btnCancelar.addEventListener("click", cancelarGeracao);
  elementos.btnLimpar.addEventListener("click", limparTela);
  elementos.hintNote.addEventListener("click", abrirPastaFinal);

  carregarModeloInfo();
}

function apiDisponivel(nome) {
  return !!(window.pywebview && window.pywebview.api && typeof window.pywebview.api[nome] === "function");
}

async function carregarModeloInfo() {
  if (!apiDisponivel("gerdoc_modelo_info")) {
    atualizarStatus("Abra esta tela pelo app para conectar ao backend.", "warn");
    elementos.hintNote.textContent = "Backend pywebview não encontrado.";
    return;
  }

  try {
    const resposta = await window.pywebview.api.gerdoc_modelo_info();
    if (!resposta || resposta.ok === false) {
      atualizarStatus("Backend indisponível para o gerador.", "err");
      elementos.hintNote.textContent = resposta?.error || "Falha ao consultar o backend.";
      return;
    }

    if (!resposta.exists) {
      atualizarStatus("Modelo Word fixo não encontrado.", "err");
      elementos.hintNote.textContent = resposta.path || "Coloque MODELO JUNTADA.docx em src/docs.";
      return;
    }

    atualizarStatusPorCampos();
    elementos.hintNote.textContent = `Modelo: ${resposta.name}`;
  } catch (erro) {
    atualizarStatus("Erro ao consultar o backend.", "err");
    elementos.hintNote.textContent = erro.message;
  }
}

async function selecionarArquivo() {
  limparPastaFinal();

  if (apiDisponivel("gerdoc_pick_planilha")) {
    try {
      const resposta = await window.pywebview.api.gerdoc_pick_planilha();
      if (!resposta || resposta.canceled) {
        return;
      }

      if (resposta.ok === false) {
        atualizarStatus(resposta.error || "Arquivo inválido.", "err");
        return;
      }

      estado.arquivo = null;
      estado.arquivoPath = resposta.path;
      elementos.arquivoExcel.value = "";
      elementos.arquivoExcelTexto.value = resposta.name || resposta.path;

      atualizarStatusPorCampos();
      return;
    } catch (erro) {
      atualizarStatus("Não foi possível selecionar a planilha.", "err");
      elementos.hintNote.textContent = erro.message;
      return;
    }
  }

  elementos.arquivoExcel.click();
}

function carregarArquivo(evento) {
  limparPastaFinal();

  const arquivo = evento.target.files[0];
  if (!arquivo) {
    return;
  }

  const extensao = obterExtensao(arquivo.name);

  if (!extensoesPermitidas.includes(extensao)) {
    estado.arquivo = null;
    estado.arquivoPath = "";
    elementos.arquivoExcel.value = "";
    elementos.arquivoExcelTexto.value = "";

    atualizarStatus("Arquivo inválido. Selecione Excel ou CSV.", "err");
    return;
  }

  estado.arquivo = arquivo;
  estado.arquivoPath = "";
  elementos.arquivoExcelTexto.value = `${arquivo.name} (${formatarTamanho(arquivo.size)})`;

  atualizarStatus("Selecione a planilha pelo diálogo do app.", "warn");
  elementos.hintNote.textContent = "O navegador não fornece o caminho real do arquivo.";
}

async function selecionarPasta() {
  limparPastaFinal();

  if (apiDisponivel("gerdoc_pick_saida")) {
    try {
      const resposta = await window.pywebview.api.gerdoc_pick_saida();
      if (!resposta || resposta.canceled) {
        return;
      }

      if (resposta.ok === false) {
        atualizarStatus(resposta.error || "Pasta inválida.", "err");
        return;
      }

      estado.pastaSaida = resposta.path;
      elementos.pastaSaidaTexto.value = resposta.path;

      atualizarStatusPorCampos();
      return;
    } catch (erro) {
      atualizarStatus("Não foi possível selecionar a pasta.", "err");
      elementos.hintNote.textContent = erro.message;
      return;
    }
  }

  atualizarStatus("Abra esta tela pelo app para selecionar a pasta.", "err");
  elementos.hintNote.textContent = "Backend pywebview não encontrado.";
}

async function gerarDocumentos() {
  if (estado.gerando) {
    return;
  }

  if (!apiDisponivel("gerdoc_validar") || !apiDisponivel("gerdoc_gerar")) {
    atualizarStatus("Backend pywebview não encontrado.", "err");
    elementos.hintNote.textContent = "Abra esta tela pelo app principal.";
    return;
  }

  if (!estado.arquivoPath || !estado.pastaSaida) {
    atualizarStatus("Selecione a planilha e a pasta de saída.", "err");
    return;
  }

  const payload = {
    planilha: estado.arquivoPath,
    pasta_saida: estado.pastaSaida,
    gerar_pdf: elementos.gerarPdf.checked
  };

  estado.gerando = true;
  estado.jobId = "";
  estado.cancelamentoSolicitado = false;
  limparPastaFinal();
  bloquearTela(true);

  const formato = elementos.gerarPdf.checked ? "Word e PDF" : "Word";
  atualizarStatus(`Validando geração em ${formato}...`, "warn");
  elementos.hintNote.textContent = "Validando planilha, pasta de saída e modelo.";

  try {
    const validacao = await window.pywebview.api.gerdoc_validar(payload);
    if (!validacao || validacao.ok === false) {
      throw new Error(validacao?.error || "Falha na validação.");
    }

    atualizarStatus(`Gerando documentos em ${formato}...`, "warn");
    elementos.hintNote.textContent = `${validacao.info?.linhas ?? 0} linha(s) para processar.`;

    const resposta = await window.pywebview.api.gerdoc_gerar(payload);
    if (!resposta || resposta.ok === false) {
      throw new Error(resposta?.error || "Falha ao iniciar geração.");
    }

    estado.jobId = resposta.job_id;
    renderizarJob(resposta.job);
    agendarPolling();
  } catch (erro) {
    estado.gerando = false;
    estado.jobId = "";
    bloquearTela(false);
    atualizarStatus("Erro ao gerar documentos.", "err");
    elementos.hintNote.textContent = erro.message;
  }
}

async function consultarStatus() {
  if (!estado.jobId || !apiDisponivel("gerdoc_status")) {
    return;
  }

  try {
    const resposta = await window.pywebview.api.gerdoc_status({ job_id: estado.jobId });
    if (!resposta || resposta.ok === false) {
      throw new Error(resposta?.error || "Falha ao consultar status.");
    }

    renderizarJob(resposta.job);

    if (resposta.job?.status === "processando") {
      agendarPolling();
    }
  } catch (erro) {
    estado.gerando = false;
    estado.jobId = "";
    bloquearTela(false);
    atualizarStatus("Erro ao consultar a geração.", "err");
    elementos.hintNote.textContent = erro.message;
  }
}

function agendarPolling() {
  if (estado.pollTimer) {
    window.clearTimeout(estado.pollTimer);
  }

  estado.pollTimer = window.setTimeout(consultarStatus, 900);
}

function renderizarJob(job) {
  if (!job) {
    return;
  }

  const progresso = job.linhas ? `${job.total || 0}/${job.linhas}` : `${job.total || 0}`;
  const logs = Array.isArray(job.logs) ? job.logs : [];
  const ultimoLog = logs.length ? logs[logs.length - 1] : "";

  if (job.status === "processando") {
    atualizarStatus(`Gerando documentos... ${progresso}`, "warn");
    elementos.hintNote.textContent = ultimoLog || "Processando documentos.";
    return;
  }

  estado.gerando = false;
  estado.jobId = "";
  estado.cancelamentoSolicitado = false;
  bloquearTela(false);

  if (estado.pollTimer) {
    window.clearTimeout(estado.pollTimer);
    estado.pollTimer = null;
  }

  if (job.status === "concluido") {
    atualizarStatus(`Concluído. ${job.total || 0} linha(s) processada(s).`, "ok");
    marcarPastaFinal(job.pasta_final);
    return;
  }

  if (job.status === "cancelado") {
    atualizarStatus(`Geração cancelada. ${job.total || 0} linha(s) processada(s).`, "warn");
    marcarPastaFinal(job.pasta_final);
    return;
  }

  if (job.status === "erro") {
    atualizarStatus("Erro na geração.", "err");
    elementos.hintNote.textContent = job.erro || ultimoLog || "Falha ao gerar documentos.";
  }
}

async function cancelarGeracao() {
  if (!estado.gerando) {
    return;
  }

  estado.cancelamentoSolicitado = true;
  elementos.btnCancelar.disabled = true;

  if (!estado.jobId || !apiDisponivel("gerdoc_cancelar")) {
    estado.gerando = false;
    estado.jobId = "";
    bloquearTela(false);
    atualizarStatus("Geração cancelada.", "warn");
    elementos.hintNote.textContent = "Operação cancelada pelo usuário.";
    return;
  }

  try {
    const resposta = await window.pywebview.api.gerdoc_cancelar({ job_id: estado.jobId });
    if (!resposta || resposta.ok === false) {
      throw new Error(resposta?.error || "Falha ao cancelar.");
    }

    atualizarStatus("Cancelamento solicitado.", "warn");
    elementos.hintNote.textContent = "Aguardando finalizar o documento atual.";
  } catch (erro) {
    estado.cancelamentoSolicitado = false;
    elementos.btnCancelar.disabled = false;
    atualizarStatus("Erro ao solicitar cancelamento.", "err");
    elementos.hintNote.textContent = erro.message;
  }
}

function limparTela() {
  if (estado.gerando) {
    return;
  }

  if (estado.pollTimer) {
    window.clearTimeout(estado.pollTimer);
    estado.pollTimer = null;
  }

  estado.arquivo = null;
  estado.arquivoPath = "";
  estado.pastaSaida = "";
  estado.jobId = "";
  estado.cancelamentoSolicitado = false;
  limparPastaFinal();

  elementos.arquivoExcel.value = "";
  elementos.arquivoExcelTexto.value = "";
  elementos.pastaSaidaTexto.value = "";
  elementos.gerarPdf.checked = false;

  atualizarStatus("Selecione a planilha e a pasta de saída.", "warn");
  elementos.hintNote.textContent = "Aguardando integração com o backend.";
}

function bloquearTela(bloquear) {
  elementos.btnGerar.disabled = bloquear;
  elementos.btnCancelar.disabled = !bloquear || estado.cancelamentoSolicitado;
  elementos.btnSelecionarArquivo.disabled = bloquear;
  elementos.btnSelecionarPasta.disabled = bloquear;
  elementos.btnLimpar.disabled = bloquear;
  elementos.gerarPdf.disabled = bloquear;
}

function atualizarStatusPorCampos() {
  if (estado.arquivoPath && estado.pastaSaida) {
    atualizarStatus("Pronto para gerar documentos.", "ok");
    elementos.hintNote.textContent = elementos.gerarPdf.checked
      ? "Saída configurada para Word e PDF."
      : "Saída configurada para Word.";
    return;
  }

  if (estado.arquivoPath) {
    atualizarStatus("Arquivo selecionado. Selecione a pasta de saída.", "warn");
    elementos.hintNote.textContent = "Falta selecionar a pasta base de saída.";
    return;
  }

  if (estado.pastaSaida) {
    atualizarStatus("Pasta selecionada. Selecione a planilha.", "warn");
    elementos.hintNote.textContent = "Falta selecionar a planilha Excel ou CSV.";
    return;
  }

  atualizarStatus("Selecione a planilha e a pasta de saída.", "warn");
  elementos.hintNote.textContent = "Aguardando integração com o backend.";
}

function atualizarStatus(mensagem, tipo = "warn") {
  elementos.statusText.textContent = mensagem;

  elementos.dotApp.classList.remove("ok", "warn", "err");
  elementos.dotApp.classList.add(tipo);
}

function marcarPastaFinal(caminho) {
  if (!caminho) {
    elementos.hintNote.textContent = "Nenhuma pasta final informada pelo backend.";
    limparPastaFinal();
    return;
  }

  estado.ultimaPastaFinal = caminho;
  elementos.hintNote.textContent = `Pasta final: ${caminho}`;
  elementos.hintNote.title = "Clique para abrir a pasta final";
  elementos.hintNote.style.cursor = "pointer";
}

function limparPastaFinal() {
  estado.ultimaPastaFinal = "";
  elementos.hintNote.title = "";
  elementos.hintNote.style.cursor = "";
}

async function abrirPastaFinal() {
  if (!estado.ultimaPastaFinal || !apiDisponivel("gerdoc_abrir_pasta")) {
    return;
  }

  try {
    const resposta = await window.pywebview.api.gerdoc_abrir_pasta({ caminho: estado.ultimaPastaFinal });
    if (!resposta || resposta.ok === false) {
      throw new Error(resposta?.error || "Falha ao abrir pasta.");
    }
  } catch (erro) {
    atualizarStatus("Não foi possível abrir a pasta.", "err");
    elementos.hintNote.textContent = erro.message;
  }
}

function obterExtensao(nomeArquivo) {
  const indice = nomeArquivo.lastIndexOf(".");

  if (indice === -1) {
    return "";
  }

  return nomeArquivo.slice(indice).toLowerCase();
}

function formatarTamanho(bytes) {
  if (!Number.isFinite(bytes)) {
    return "0 KB";
  }

  return `${Math.max(1, Math.round(bytes / 1024))} KB`;
}

document.addEventListener("DOMContentLoaded", iniciar);
