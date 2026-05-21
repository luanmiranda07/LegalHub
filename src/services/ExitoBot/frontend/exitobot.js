const form = document.getElementById("formExitoBot");

const btnSelecionarPasta = document.getElementById("btnSelecionarPasta");
const btnSelecionarExcel = document.getElementById("btnSelecionarExcel");
const btnValidar = document.getElementById("btnValidar");
const btnExecutar = document.getElementById("btnExecutar");
const btnLimparLog = document.getElementById("btnLimparLog");

const inputPasta = document.getElementById("inputPasta");
const inputExcel = document.getElementById("inputExcel");

const pasta = document.getElementById("pasta");
const excel = document.getElementById("excel");

const mensagem = document.getElementById("mensagem");
const resultado = document.getElementById("resultado");
const statusBadge = document.getElementById("statusBadge");

const estado = {
  pastaPath: "",
  excelPath: "",
  executando: false
};

function apiDisponivel(nome) {
  return !!(window.pywebview && window.pywebview.api && typeof window.pywebview.api[nome] === "function");
}

function atualizarStatus(texto, tipo = "info") {
  statusBadge.textContent = texto;
  statusBadge.classList.remove("status-ok", "status-err");

  if (tipo === "ok") {
    statusBadge.classList.add("status-ok");
  }

  if (tipo === "erro") {
    statusBadge.classList.add("status-err");
  }
}

function escreverMensagem(texto, tipo = "info") {
  mensagem.textContent = texto;

  if (tipo === "erro") {
    atualizarStatus("Atencao", "erro");
    return;
  }

  if (tipo === "ok") {
    atualizarStatus("Pronto", "ok");
    return;
  }

  atualizarStatus("Atualizado");
}

function escreverLog(texto) {
  const data = new Date();
  const hora = data.toLocaleTimeString("pt-BR", {
    hour: "2-digit",
    minute: "2-digit",
    second: "2-digit"
  });

  resultado.value += `[${hora}] ${texto}\n`;
  resultado.scrollTop = resultado.scrollHeight;
}

function escreverLogs(logs) {
  for (const linha of logs || []) {
    escreverLog(linha);
  }
}

function limparLog() {
  resultado.value = "";
  escreverMensagem("Log limpo. Selecione a pasta e o Excel.");
}

function bloquearTela(bloquear) {
  estado.executando = bloquear;
  btnSelecionarPasta.disabled = bloquear;
  btnSelecionarExcel.disabled = bloquear;
  btnValidar.disabled = bloquear;
  btnExecutar.disabled = bloquear;
  btnLimparLog.disabled = bloquear;
}

function validarCampos() {
  if (!estado.pastaPath) {
    escreverMensagem("Selecione uma pasta antes de continuar.", "erro");
    escreverLog("Validacao interrompida: pasta nao selecionada.");
    return false;
  }

  if (!estado.excelPath) {
    escreverMensagem("Selecione um arquivo Excel antes de continuar.", "erro");
    escreverLog("Validacao interrompida: Excel nao selecionado.");
    return false;
  }

  return true;
}

function montarPayload() {
  return {
    pasta: estado.pastaPath,
    excel: estado.excelPath
  };
}

function mostrarBackendIndisponivel() {
  escreverMensagem("Abra esta tela pelo app para conectar ao backend.", "erro");
  escreverLog("Backend pywebview nao encontrado. Execute pelo aplicativo Legal Hub.");
}

async function selecionarPasta() {
  if (estado.executando) {
    return;
  }

  if (!apiDisponivel("pick_folder")) {
    mostrarBackendIndisponivel();
    return;
  }

  try {
    const resposta = await window.pywebview.api.pick_folder();
    if (!resposta || resposta.canceled) {
      escreverMensagem("Selecao de pasta cancelada.");
      return;
    }

    if (resposta.ok === false) {
      escreverMensagem(resposta.error || "Nao foi possivel selecionar a pasta.", "erro");
      return;
    }

    estado.pastaPath = resposta.folderPath || resposta.path || "";
    pasta.value = resposta.name ? `${resposta.name} - ${estado.pastaPath}` : estado.pastaPath;
    escreverMensagem("Pasta selecionada.", "ok");
    escreverLog(`Pasta selecionada: ${estado.pastaPath}`);
  } catch (erro) {
    escreverMensagem("Nao foi possivel selecionar a pasta.", "erro");
    escreverLog(`[ERRO] ${erro.message}`);
  }
}

async function selecionarExcel() {
  if (estado.executando) {
    return;
  }

  if (!apiDisponivel("exitobot_pick_excel")) {
    mostrarBackendIndisponivel();
    return;
  }

  try {
    const resposta = await window.pywebview.api.exitobot_pick_excel();
    if (!resposta || resposta.canceled) {
      escreverMensagem("Selecao de Excel cancelada.");
      return;
    }

    if (resposta.ok === false) {
      escreverMensagem(resposta.error || "Nao foi possivel selecionar o Excel.", "erro");
      return;
    }

    estado.excelPath = resposta.excelPath || resposta.path || "";
    excel.value = resposta.name ? `${resposta.name} - ${estado.excelPath}` : estado.excelPath;
    escreverMensagem("Excel selecionado.", "ok");
    escreverLog(`Excel selecionado: ${estado.excelPath}`);
  } catch (erro) {
    escreverMensagem("Nao foi possivel selecionar o Excel.", "erro");
    escreverLog(`[ERRO] ${erro.message}`);
  }
}

async function processar(tipo) {
  if (estado.executando || !validarCampos()) {
    return;
  }

  const metodo = tipo === "validar" ? "exitobot_validar" : "exitobot_executar";
  if (!apiDisponivel(metodo)) {
    mostrarBackendIndisponivel();
    return;
  }

  bloquearTela(true);
  resultado.value = "";
  escreverMensagem(tipo === "validar" ? "Validando..." : "Executando...");
  escreverLog(tipo === "validar" ? "Iniciando validacao no backend." : "Iniciando execucao no backend.");

  try {
    const resposta = await window.pywebview.api[metodo](montarPayload());

    if (!resposta) {
      escreverMensagem("Backend nao retornou resposta.", "erro");
      escreverLog("[ERRO] Backend nao retornou resposta.");
      return;
    }

    escreverLogs(resposta.logs || []);

    if (resposta.ok === false) {
      escreverMensagem(resposta.mensagem || resposta.error || "Falha ao processar.", "erro");
      return;
    }

    escreverMensagem(resposta.mensagem || "Operacao concluida.", "ok");
    escreverLog(
      `Resumo: total=${resposta.total || 0}, prontos=${resposta.prontos || 0}, ` +
      `ignorados=${resposta.ignorados || 0}, renomeados=${resposta.renomeados || 0}, erros=${resposta.erros || 0}`
    );
  } catch (erro) {
    escreverMensagem("Erro ao processar.", "erro");
    escreverLog(`[ERRO] ${erro.message}`);
  } finally {
    bloquearTela(false);
  }
}

btnSelecionarPasta.addEventListener("click", selecionarPasta);
btnSelecionarExcel.addEventListener("click", selecionarExcel);
btnValidar.addEventListener("click", () => processar("validar"));
btnLimparLog.addEventListener("click", limparLog);

form.addEventListener("submit", (event) => {
  event.preventDefault();
  processar("executar");
});

if (inputPasta) {
  inputPasta.disabled = true;
}

if (inputExcel) {
  inputExcel.disabled = true;
}

escreverLog("Tela Adicionar Chance de Exito carregada. Aguardando selecao da pasta e do Excel.");
