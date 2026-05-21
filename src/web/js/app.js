// Proteção simples: se não estiver logado, volta para login
if (localStorage.getItem("cc_logged") !== "1") {
  window.location.href = "login.html";
}

document.getElementById("btnLogout").addEventListener("click", () => {
  localStorage.removeItem("cc_logged");
  window.location.href = "../pages/login.html";
});

const automacoes = [
 
  {
    id: "consolidação-Processual",
    titulo: "Consolidação Processual",
    descricao: "Geração de documentos a partir de modelos e dados preenchidos.",
    status: "active", // active | draft
    pagina: "../../services/PDF_para_EXCEL/frontend/pdf_para_excel.html"
  },
 
  {
    id: "analise-de-pastas",
    titulo: "Analise de Pastas",
    descricao: "Organizado de pastas e unificando todas as pasta em um modelo Excel",
    status: "active",
    pagina: "../../services/Analise_pastas/frontend/analise_de_pastas.html"
  },

  {
    id: "gerador-lotes",
    titulo: "Gerador de Lotes",
    descricao: "Gerador de Lotes Preenchendo todas as planilhas necessaria para o ROBO do CPJ",
    status: "active",
    pagina: "../../services/Gerador_lotes/frontend/gerador_lotes.html"
  },

  {
    id: "relatorios-conformidade",
    titulo: "Relatório Conformidade",
    descricao: "Gera os Modelos preenchido tanto RPV , quanto PRECA e prevendo as datas de termino do processo.",
    status: "active",
    pagina: "../../services/Relatorio_Conformidade/frontend/relatorio_conformidade.html"
  },

  {
    id: "ia-laudos",
    titulo: "IA Laudos",
    descricao: "IA Analisadora de Laudos.",
    status: "draft", // active | draft
    pagina: "../../services/IA_LAUDO/frontend/ialaudos.html"
  },

  {
    id: "ia-pet-inicial",
    titulo: "IA PET INICIAL",
    descricao: "IA Geradora de Petição Inicial.",
    status: "draft", // active | draft
    pagina: "../../services/IA_PET_INICIAL/frontend/ia-pet-inicial.html"
  },

  {
    id: "calculadora-previdenciaria",
    titulo: "Calculadora Previdenciaria",
    descricao: "Calculadora para Calcular Beneficio Previdenciario .",
    status: "draft", // active | draft
    pagina: "../pages/desenvolvimento.html"
  },

   {
    id: "gerador-documentos",
    titulo: "Gerador de Documentos",
    descricao: "Gerador de documentos a partir de modelos e dados preenchidos.",
    status: "active", // active | draft
    pagina: "../../services/gerador_documentos/frontend/geradordocumentos.html"
  },

  {
    id: "analistduplic",
    titulo: "AnalistDuplic",
    descricao: "Análise de duplicidade em processos.",
    status: "active", // active | draft
    pagina: "../../services/AnalistDuplic/frontend/analisduplic.html"
  },

  {
    id: "divisor-pdf",
    titulo: "Divisor de PDF",
    descricao: "Divisão de documentos PDF em arquivos menores.",
    status: "active", // active | draft
    pagina: "../../services/divisor_pdf/frontend/divisorpdf.html"
  },

  {
    id: "exitobot",
    titulo: "Exitobot",
    descricao: "Bot para validação e execução de processos.",
    status: "active", // active | draft
    pagina: "../../services/ExitoBot/frontend/exitobot.html",
    manualPdf: "../../docs/manual-exitobot.pdf",
  
  },

];

let filtroStatus = "all";

const elCards = document.getElementById("cards");
const elEmpty = document.getElementById("emptyState");
const elSearch = document.getElementById("searchInput");
const elDetailsModal = document.getElementById("detailsModal");
const elDetailsTitle = document.getElementById("detailsTitle");
const elDetailsStatus = document.getElementById("detailsStatus");
const elDetailsDescription = document.getElementById("detailsDescription");
const elDetailsPage = document.getElementById("detailsPage");
const elManualLink = document.getElementById("manualLink");
const elManualEmpty = document.getElementById("manualEmpty");
const elDetailsOpenPage = document.getElementById("detailsOpenPage");
const elDetailsClose = document.getElementById("detailsClose");

let automacaoSelecionada = null;

if (elDetailsPage) {
  elDetailsPage.closest(".modal-row").style.display = "none";
}

function render() {
  const q = (elSearch.value || "").trim().toLowerCase();

  const itens = automacoes.filter(a => {
    const matchTexto =
      a.titulo.toLowerCase().includes(q) ||
      a.descricao.toLowerCase().includes(q);

    const matchStatus =
      filtroStatus === "all" ? true : a.status === filtroStatus;

    return matchTexto && matchStatus;
  });

  elCards.innerHTML = "";

  if (itens.length === 0) {
    elEmpty.style.display = "block";
    return;
  }

  elEmpty.style.display = "none";

  for (const a of itens) {
    const statusLabel = a.status === "active" ? "Ativa" : "Em construção";
    const statusClass = a.status;
    const documentacaoAction = a.documentacaoPdf
      ? `<button class="btn" data-doc="${escapeHtml(a.documentacaoPdf)}">Abrir documentação</button>`
      : "";

    const card = document.createElement("div");
    card.className = "card";

    card.innerHTML = `
      <div class="card-top">
        <div><h3>${escapeHtml(a.titulo)}</h3></div>
        <div class="status ${statusClass}">${statusLabel}</div>
      </div>

      <p class="desc">${escapeHtml(a.descricao)}</p>

      <div class="actions">
        <button class="btn primary" data-open="${a.pagina}">Abrir</button>
        <button class="btn" data-info="${a.id}">Detalhes</button>
        ${documentacaoAction}
      </div>
    `;

    elCards.appendChild(card);
  }

  elCards.querySelectorAll("[data-open]").forEach(btn => {
    btn.addEventListener("click", () => {
      const pagina = btn.getAttribute("data-open");
      window.location.href = pagina;
    });
  });

  elCards.querySelectorAll("[data-info]").forEach(btn => {
    btn.addEventListener("click", () => {
      const id = btn.getAttribute("data-info");
      const a = automacoes.find(x => x.id === id);
      if (!a) return;
      openDetails(a);
    });
  });

  elCards.querySelectorAll("[data-doc]").forEach(btn => {
    btn.addEventListener("click", () => {
      const pdf = btn.getAttribute("data-doc");
      if (!pdf) return;
      window.open(pdf, "_blank", "noopener");
    });
  });
}

function getStatusLabel(status) {
  return status === "active" ? "Ativa" : "Em construção";
}

function openDetails(a) {
  automacaoSelecionada = a;

  elDetailsTitle.textContent = a.titulo;
  elDetailsStatus.textContent = getStatusLabel(a.status);
  elDetailsDescription.textContent = a.descricao;

  if (a.manualPdf) {
    elManualLink.href = a.manualPdf;
    elManualLink.style.display = "inline-flex";
    elManualEmpty.style.display = "none";
  } else {
    elManualLink.removeAttribute("href");
    elManualLink.style.display = "none";
    elManualEmpty.style.display = "block";
  }

  elDetailsModal.classList.add("open");
  elDetailsModal.setAttribute("aria-hidden", "false");
}

function closeDetails() {
  automacaoSelecionada = null;
  elDetailsModal.classList.remove("open");
  elDetailsModal.setAttribute("aria-hidden", "true");
}

function escapeHtml(str) {
  return String(str)
    .replaceAll("&", "&amp;")
    .replaceAll("<", "&lt;")
    .replaceAll(">", "&gt;")
    .replaceAll('"', "&quot;")
    .replaceAll("'", "&#039;");
}

// Filtros
document.getElementById("btnAll").addEventListener("click", () => { filtroStatus = "all"; render(); });
document.getElementById("btnActive").addEventListener("click", () => { filtroStatus = "active"; render(); });
document.getElementById("btnDraft").addEventListener("click", () => { filtroStatus = "draft"; render(); });

// Busca
elSearch.addEventListener("input", render);

// Modal de detalhes
elDetailsClose.addEventListener("click", closeDetails);
elDetailsModal.addEventListener("click", (e) => {
  if (e.target === elDetailsModal) closeDetails();
});
elDetailsOpenPage.addEventListener("click", () => {
  if (!automacaoSelecionada) return;
  window.location.href = automacaoSelecionada.pagina;
});

// Atalhos
document.addEventListener("keydown", (e) => {
  if (e.key === "Escape" && elDetailsModal.classList.contains("open")) {
    closeDetails();
    return;
  }
  if (e.key === "/" && document.activeElement !== elSearch) {
    e.preventDefault();
    elSearch.focus();
  }
  if (e.key === "Escape") {
    elSearch.value = "";
    render();
  }
});

// Relógio
function tickClock() {
  const now = new Date();
  const hh = String(now.getHours()).padStart(2, "0");
  const mm = String(now.getMinutes()).padStart(2, "0");
  document.getElementById("clock").textContent = `${hh}:${mm}`;
}
setInterval(tickClock, 1000);
tickClock();
  // Inicialização: aplica filtro ao abrir a página (ex.: vindo do login)
  (function init() {
    const tab = (new URLSearchParams(window.location.search).get("tab") || "").toLowerCase();

    if (tab === "active") filtroStatus = "active";
    else if (tab === "draft") filtroStatus = "draft";
    else filtroStatus = "all";

    render();
  })();


