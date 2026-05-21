const form = document.getElementById("formDivisorPdf");

const btnSelecionarPdf = document.getElementById("btnSelecionarPdf");
const btnSelecionarPasta = document.getElementById("btnSelecionarPasta");
const btnAbrirPasta = document.getElementById("btnAbrirPasta");
const btnFechar = document.getElementById("btnFechar");
const btnDividirPdf = document.getElementById("btnDividirPdf");

const inputPdf = document.getElementById("inputPdf");
const inputPasta = document.getElementById("inputPasta");

const pdfEntrada = document.getElementById("pdfEntrada");
const pastaSaida = document.getElementById("pastaSaida");
const nomeBase = document.getElementById("nomeBase");

const paginasPorArquivo = document.getElementById("paginasPorArquivo");
const intervalos = document.getElementById("intervalos");
const palavraChave = document.getElementById("palavraChave");

const descricaoModo = document.getElementById("descricaoModo");
const resultado = document.getElementById("resultado");
const statusBadge = document.getElementById("statusBadge");

const state = {
  pdfPath: "",
  outputDir: ""
};

const descricoesModo = {
  pagina: "Cria um PDF para cada pagina do arquivo original.",
  quantidade: "Cria um PDF a cada quantidade de paginas definida.",
  intervalos: "Cria PDFs conforme os intervalos personalizados informados.",
  palavra: "Cria PDFs separando o documento a partir da palavra-chave informada."
};

function hasPywebviewApi() {
  return Boolean(window.pywebview && window.pywebview.api);
}

function api() {
  return hasPywebviewApi() ? window.pywebview.api : null;
}

function obterModoSelecionado() {
  const radioSelecionado = document.querySelector('input[name="modoDivisao"]:checked');
  return radioSelecionado ? radioSelecionado.value : "pagina";
}

function atualizarDescricaoModo() {
  const modo = obterModoSelecionado();
  descricaoModo.textContent = descricoesModo[modo] || descricoesModo.pagina;
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

function mostrarResultado(mensagem, tipo = "info") {
  resultado.textContent = mensagem;

  if (tipo === "erro") {
    atualizarStatus("Atencao", "erro");
    return;
  }

  if (tipo === "ok") {
    atualizarStatus("Pronto", "ok");
    return;
  }

  if (tipo === "processando") {
    atualizarStatus("Processando");
    return;
  }

  atualizarStatus("Atualizado");
}

function formatarSucesso(res) {
  const files = Array.isArray(res.output_files) ? res.output_files : [];
  const nomes = files.slice(0, 12).map((filePath) => {
    const partes = String(filePath).split(/[\\/]/);
    return `- ${partes[partes.length - 1]}`;
  });

  if (files.length > 12) {
    nomes.push(`... e mais ${files.length - 12} arquivo(s).`);
  }

  return [
    res.details || `PDF dividido em ${res.count || files.length} arquivo(s).`,
    "",
    `Pasta de saida: ${res.output_dir || state.outputDir}`,
    "",
    "Arquivos gerados:",
    nomes.join("\n") || "- Nenhum arquivo listado."
  ].join("\n");
}

function validarFormulario() {
  const modo = obterModoSelecionado();

  if (!state.pdfPath) {
    mostrarResultado("Selecione um PDF de entrada pelo botao Selecionar.", "erro");
    return false;
  }

  if (!state.outputDir) {
    mostrarResultado("Selecione uma pasta de saida pelo botao Selecionar.", "erro");
    return false;
  }

  if (!nomeBase.value.trim()) {
    mostrarResultado("Informe o nome base dos arquivos.", "erro");
    return false;
  }

  if (modo === "quantidade" && Number(paginasPorArquivo.value) < 1) {
    mostrarResultado("Informe uma quantidade de paginas maior ou igual a 1.", "erro");
    return false;
  }

  if (modo === "intervalos" && !intervalos.value.trim()) {
    mostrarResultado("Informe pelo menos um intervalo. Exemplo: 1-3, 4-6.", "erro");
    return false;
  }

  if (modo === "palavra" && !palavraChave.value.trim()) {
    mostrarResultado("Informe uma palavra-chave.", "erro");
    return false;
  }

  return true;
}

function montarPayloadFrontEnd() {
  return {
    pdfEntrada: state.pdfPath,
    pastaSaida: state.outputDir,
    nomeBase: nomeBase.value.trim(),
    modoDivisao: obterModoSelecionado(),
    paginasPorArquivo: Number(paginasPorArquivo.value),
    intervalos: intervalos.value.trim(),
    palavraChave: palavraChave.value.trim()
  };
}

btnSelecionarPdf.addEventListener("click", async () => {
  const bridge = api();

  if (!bridge || typeof bridge.divisorpdf_pick_pdf !== "function") {
    inputPdf.click();
    mostrarResultado("Abra esta tela pelo aplicativo desktop para selecionar o caminho real do PDF.", "erro");
    return;
  }

  mostrarResultado("Abrindo seletor de PDF...");
  const res = await bridge.divisorpdf_pick_pdf();

  if (res && res.canceled) {
    mostrarResultado("Selecao de PDF cancelada.");
    return;
  }

  if (!res || res.ok === false) {
    mostrarResultado((res && res.error) || "Nao foi possivel selecionar o PDF.", "erro");
    return;
  }

  state.pdfPath = String(res.path || "");
  pdfEntrada.value = res.name || state.pdfPath;

  if (res.suggested_output && !state.outputDir) {
    state.outputDir = String(res.suggested_output);
    pastaSaida.value = state.outputDir;
  }

  if (res.stem && (!nomeBase.value.trim() || nomeBase.value.trim() === "arquivo_dividido")) {
    nomeBase.value = res.stem;
  }

  mostrarResultado(`PDF selecionado: ${pdfEntrada.value}.`, "ok");
});

inputPdf.addEventListener("change", () => {
  inputPdf.value = "";
  mostrarResultado("No navegador o arquivo nao fornece caminho real. Use o aplicativo desktop Legal Hub.", "erro");
});

btnSelecionarPasta.addEventListener("click", async () => {
  const bridge = api();

  if (!bridge || typeof bridge.divisorpdf_pick_saida !== "function") {
    inputPasta.click();
    mostrarResultado("Abra esta tela pelo aplicativo desktop para selecionar a pasta real de saida.", "erro");
    return;
  }

  mostrarResultado("Abrindo seletor de pasta...");
  const res = await bridge.divisorpdf_pick_saida();

  if (res && res.canceled) {
    mostrarResultado("Selecao de pasta cancelada.");
    return;
  }

  if (!res || res.ok === false) {
    mostrarResultado((res && res.error) || "Nao foi possivel selecionar a pasta de saida.", "erro");
    return;
  }

  state.outputDir = String(res.path || "");
  pastaSaida.value = state.outputDir;
  mostrarResultado(`Pasta de saida selecionada: ${state.outputDir}.`, "ok");
});

inputPasta.addEventListener("change", () => {
  inputPasta.value = "";
  mostrarResultado("No navegador a pasta nao fornece caminho real. Use o aplicativo desktop Legal Hub.", "erro");
});

document.querySelectorAll('input[name="modoDivisao"]').forEach((radio) => {
  radio.addEventListener("change", () => {
    atualizarDescricaoModo();
    mostrarResultado("Modo de divisao atualizado.");
  });
});

btnAbrirPasta.addEventListener("click", async () => {
  if (!state.outputDir) {
    mostrarResultado("Selecione uma pasta de saida antes de continuar.", "erro");
    return;
  }

  const bridge = api();
  if (!bridge || typeof bridge.divisorpdf_abrir_pasta !== "function") {
    mostrarResultado("Abra esta tela pelo aplicativo desktop para abrir a pasta de saida.", "erro");
    return;
  }

  const res = await bridge.divisorpdf_abrir_pasta({ caminho: state.outputDir });
  if (!res || res.ok === false) {
    mostrarResultado((res && res.error) || "Nao foi possivel abrir a pasta de saida.", "erro");
    return;
  }

  mostrarResultado(res.mensagem || "Pasta de saida aberta.", "ok");
});

btnFechar.addEventListener("click", () => {
  window.location.href = "../../../web/pages/index.html";
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();

  const bridge = api();
  if (!bridge || typeof bridge.divisorpdf_processar !== "function") {
    mostrarResultado("Abra esta tela pelo aplicativo desktop Legal Hub para dividir PDFs.", "erro");
    return;
  }

  if (!validarFormulario()) {
    return;
  }

  btnDividirPdf.disabled = true;
  mostrarResultado("Dividindo PDF. Aguarde...", "processando");

  try {
    const res = await bridge.divisorpdf_processar(montarPayloadFrontEnd());

    if (!res || res.ok === false) {
      mostrarResultado((res && res.error) || "Nao foi possivel dividir o PDF.", "erro");
      return;
    }

    if (res.output_dir) {
      state.outputDir = String(res.output_dir);
      pastaSaida.value = state.outputDir;
    }

    mostrarResultado(formatarSucesso(res), "ok");
  } catch (error) {
    mostrarResultado(error && error.message ? error.message : String(error), "erro");
  } finally {
    btnDividirPdf.disabled = false;
  }
});

atualizarDescricaoModo();
