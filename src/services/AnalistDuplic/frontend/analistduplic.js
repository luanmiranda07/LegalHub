const elementos = {
  pastaTexto: document.getElementById("pastaTexto"),
  numeroInicial: document.getElementById("numeroInicial"),
  ano: document.getElementById("ano"),
  usarSoNumeracao: document.getElementById("usarSoNumeracao"),

  btnSelecionarPasta: document.getElementById("btnSelecionarPasta"),
  btnExecutar: document.getElementById("btnExecutar"),
  btnIdentificarDuplicados: document.getElementById("btnIdentificarDuplicados"),
  btnCancelar: document.getElementById("btnCancelar"),
  btnLimpar: document.getElementById("btnLimpar"),

  statusText: document.getElementById("statusText"),
  dotApp: document.getElementById("dotApp"),
  hintNote: document.getElementById("hintNote")
};

const estado = {
  pasta: "",
  executando: false,
  cancelamentoSolicitado: false
};

function iniciar() {
  elementos.btnSelecionarPasta.addEventListener("click", selecionarPasta);
  elementos.btnExecutar.addEventListener("click", executar);
  elementos.btnIdentificarDuplicados.addEventListener("click", identificarDuplicados);
  elementos.btnCancelar.addEventListener("click", cancelar);
  elementos.btnLimpar.addEventListener("click", limparTela);

  elementos.numeroInicial.addEventListener("input", atualizarStatusPorCampos);
  elementos.ano.addEventListener("input", atualizarStatusPorCampos);
  elementos.usarSoNumeracao.addEventListener("change", atualizarHintConfiguracao);

  preencherAnoAtual();
  atualizarStatusPorCampos();
}

function apiDisponivel(nome) {
  return !!(window.pywebview && window.pywebview.api && typeof window.pywebview.api[nome] === "function");
}

function preencherAnoAtual() {
  elementos.ano.value = new Date().getFullYear();
}

async function selecionarPasta() {
  if (apiDisponivel("analisduplic_pick_folder")) {
    try {
      const resposta = await window.pywebview.api.analisduplic_pick_folder();
      if (!resposta || resposta.canceled) {
        return;
      }

      if (resposta.ok === false) {
        atualizarStatus(resposta.error || "Não foi possível selecionar a pasta.", "err");
        return;
      }

      estado.pasta = resposta.path || resposta.folderPath || "";
      elementos.pastaTexto.value = resposta.name ? `${resposta.name} - ${estado.pasta}` : estado.pasta;

      atualizarStatusPorCampos();
      return;
    } catch (erro) {
      atualizarStatus("Não foi possível selecionar a pasta.", "err");
      elementos.hintNote.textContent = erro.message;
      return;
    }
  }

  if ("showDirectoryPicker" in window) {
    try {
      const pasta = await window.showDirectoryPicker();

      estado.pasta = pasta.name;
      elementos.pastaTexto.value = pasta.name;

      atualizarStatus("Pasta selecionada apenas no navegador.", "warn");
      elementos.hintNote.textContent = "Abra pelo app para executar no backend.";
      return;
    } catch (erro) {
      if (erro.name !== "AbortError") {
        atualizarStatus("Não foi possível selecionar a pasta.", "err");
      }

      return;
    }
  }

  const pastaDigitada = window.prompt("Informe o caminho completo da pasta:");

  if (!pastaDigitada) {
    return;
  }

  estado.pasta = pastaDigitada.trim();
  elementos.pastaTexto.value = estado.pasta;

  atualizarStatusPorCampos();
}

async function executar() {
  if (estado.executando) {
    return;
  }

  if (!apiDisponivel("analisduplic_executar")) {
    atualizarStatus("Abra esta tela pelo app para conectar ao backend.", "err");
    elementos.hintNote.textContent = "Backend pywebview não encontrado.";
    return;
  }

  const validacao = validarCamposExecucao();

  if (!validacao.valido) {
    atualizarStatus(validacao.mensagem, "err");
    return;
  }

  await iniciarExecucao("executar");
}

async function identificarDuplicados() {
  if (estado.executando) {
    return;
  }

  if (!apiDisponivel("analisduplic_identificar_duplicados")) {
    atualizarStatus("Abra esta tela pelo app para conectar ao backend.", "err");
    elementos.hintNote.textContent = "Backend pywebview não encontrado.";
    return;
  }

  if (!estado.pasta) {
    atualizarStatus("Selecione a pasta antes de identificar duplicados.", "err");
    return;
  }

  await iniciarExecucao("duplicados");
}

async function iniciarExecucao(tipo) {
  estado.executando = true;
  estado.cancelamentoSolicitado = false;
  bloquearTela(true);

  try {
    let resposta;

    if (tipo === "executar") {
      const numeroInicial = elementos.numeroInicial.value.trim();
      const ano = elementos.ano.value.trim();
      const modo = elementos.usarSoNumeracao.checked ? "somente numeração" : "com CASO";
      const descricao = numeroInicial || ano
        ? `Executando numeração a partir de ${numeroInicial}${ano ? `/${ano}` : ""} (${modo})...`
        : "Removendo numeração das pastas...";

      atualizarStatus(descricao, "warn");
      elementos.hintNote.textContent = "Aguarde o retorno do backend.";

      resposta = await window.pywebview.api.analisduplic_executar(montarPayload());
    } else {
      atualizarStatus("Identificando duplicados na pasta selecionada...", "warn");
      elementos.hintNote.textContent = "Aguarde a geração do relatório.";

      resposta = await window.pywebview.api.analisduplic_identificar_duplicados({
        pasta: estado.pasta
      });
    }

    if (estado.cancelamentoSolicitado) {
      return;
    }

    tratarResposta(tipo, resposta);
  } catch (erro) {
    if (!estado.cancelamentoSolicitado) {
      atualizarStatus("Erro ao executar a operação.", "err");
      elementos.hintNote.textContent = erro.message;
    }
  } finally {
    estado.executando = false;
    estado.cancelamentoSolicitado = false;
    bloquearTela(false);
  }
}

function montarPayload() {
  return {
    pasta: estado.pasta,
    numeroInicial: elementos.numeroInicial.value.trim(),
    ano: elementos.ano.value.trim(),
    usarSoNumeracao: elementos.usarSoNumeracao.checked
  };
}

function tratarResposta(tipo, resposta) {
  if (!resposta) {
    atualizarStatus("Backend não retornou resposta.", "err");
    elementos.hintNote.textContent = "Tente novamente pelo app.";
    return;
  }

  if (resposta.ok === false) {
    atualizarStatus(resposta.mensagem || resposta.error || "Falha ao executar.", "err");
    elementos.hintNote.textContent = montarResumoResposta(resposta);
    return;
  }

  atualizarStatus(resposta.mensagem || "Operação concluída.", "ok");

  if (tipo === "duplicados" && resposta.arquivo_saida) {
    elementos.hintNote.textContent = `Arquivo salvo em: ${resposta.arquivo_saida}`;
    return;
  }

  elementos.hintNote.textContent = montarResumoResposta(resposta);
}

function montarResumoResposta(resposta) {
  const partes = [];

  if (typeof resposta.alteradas === "number") {
    partes.push(`Alteradas: ${resposta.alteradas}`);
  }

  if (typeof resposta.total_duplicados === "number") {
    partes.push(`Duplicados: ${resposta.total_duplicados}`);
  }

  if (Array.isArray(resposta.ignoradas)) {
    partes.push(`Ignoradas: ${resposta.ignoradas.length}`);
  }

  if (Array.isArray(resposta.erros)) {
    partes.push(`Erros: ${resposta.erros.length}`);
  }

  if (resposta.arquivo_saida) {
    partes.push(`Arquivo: ${resposta.arquivo_saida}`);
  }

  return partes.length ? partes.join(" | ") : "Operação concluída pelo backend.";
}

function cancelar() {
  if (!estado.executando) {
    return;
  }

  estado.cancelamentoSolicitado = true;
  atualizarStatus("Operação cancelada na tela.", "warn");
  elementos.hintNote.textContent = "A chamada em andamento pode terminar no backend.";
}

function limparTela() {
  if (estado.executando) {
    return;
  }

  estado.pasta = "";

  elementos.pastaTexto.value = "";
  elementos.numeroInicial.value = "";
  preencherAnoAtual();
  elementos.usarSoNumeracao.checked = false;

  atualizarStatus("Selecione a pasta e informe os dados para executar.", "warn");
  elementos.hintNote.textContent = apiDisponivel("analisduplic_executar")
    ? "Backend conectado."
    : "Aguardando integração com o backend.";
}

function validarCamposExecucao() {
  if (!estado.pasta) {
    return {
      valido: false,
      mensagem: "Selecione a pasta antes de executar."
    };
  }

  const numeroInicialTexto = elementos.numeroInicial.value.trim();
  const anoTexto = elementos.ano.value.trim();

  if (!numeroInicialTexto && !anoTexto) {
    return {
      valido: true,
      mensagem: ""
    };
  }

  if (!numeroInicialTexto) {
    return {
      valido: false,
      mensagem: "Informe o número inicial ou deixe número e ano vazios para limpar."
    };
  }

  const numeroInicial = Number(numeroInicialTexto);

  if (!Number.isInteger(numeroInicial) || numeroInicial <= 0) {
    return {
      valido: false,
      mensagem: "Informe um número inicial válido."
    };
  }

  if (anoTexto) {
    const ano = Number(anoTexto);

    if (!Number.isInteger(ano) || ano < 1900 || ano > 2999) {
      return {
        valido: false,
        mensagem: "Informe um ano válido."
      };
    }
  }

  return {
    valido: true,
    mensagem: ""
  };
}

function bloquearTela(bloquear) {
  elementos.btnExecutar.disabled = bloquear;
  elementos.btnIdentificarDuplicados.disabled = bloquear;
  elementos.btnSelecionarPasta.disabled = bloquear;
  elementos.btnLimpar.disabled = bloquear;
  elementos.numeroInicial.disabled = bloquear;
  elementos.ano.disabled = bloquear;
  elementos.usarSoNumeracao.disabled = bloquear;
  elementos.btnCancelar.disabled = !bloquear;
}

function atualizarStatusPorCampos() {
  if (!estado.pasta) {
    atualizarStatus("Selecione a pasta e informe os dados para executar.", "warn");
    elementos.hintNote.textContent = apiDisponivel("analisduplic_executar")
      ? "Backend conectado."
      : "Falta selecionar a pasta.";
    return;
  }

  const numeroInicial = elementos.numeroInicial.value.trim();
  const ano = elementos.ano.value.trim();

  if (!numeroInicial && !ano) {
    atualizarStatus("Pronto para remover numeração existente.", "ok");
    elementos.hintNote.textContent = "Número e ano vazios ativam a limpeza.";
    return;
  }

  if (!numeroInicial) {
    atualizarStatus("Informe o número inicial.", "warn");
    elementos.hintNote.textContent = "Para limpar, deixe número e ano vazios.";
    return;
  }

  atualizarStatus("Pronto para executar.", "ok");
  atualizarHintConfiguracao();
}

function atualizarHintConfiguracao() {
  if (!estado.pasta) {
    return;
  }

  elementos.hintNote.textContent = elementos.usarSoNumeracao.checked
    ? "Modo: somente numeração, sem CASO."
    : "Modo: numeração com CASO.";
}

function atualizarStatus(mensagem, tipo = "warn") {
  elementos.statusText.textContent = mensagem;

  elementos.dotApp.classList.remove("ok", "warn", "err");
  elementos.dotApp.classList.add(tipo);
}

document.addEventListener("DOMContentLoaded", iniciar);
window.addEventListener("pywebviewready", atualizarStatusPorCampos);
