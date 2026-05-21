const form = document.getElementById("loginForm");
const errorBox = document.getElementById("errorBox");

document.getElementById("year").textContent = new Date().getFullYear();

// Se já estiver logado, vai direto para a home (mesma pasta pages)
if (localStorage.getItem("cc_logged") === "1") {
  window.location.href = "../pages/index.html?tab=active";
}

// Se seu FRONT estiver em outro host/porta, deixe assim:
const API_URL = "http://localhost:8000/api/login";
// Se futuramente o front for servido pelo mesmo backend, você pode trocar por:
// const API_URL = `${window.location.origin}/api/login`;
const DEFAULT_AUTH_ERROR = "Credenciais inválidas. Verifique usuário e senha.";
const DEFAULT_CONNECTIVITY_ERROR = "Erro ao conectar no servidor de login.";

function hasPywebviewLoginApi() {
  return Boolean(
    window.pywebview &&
      window.pywebview.api &&
      typeof window.pywebview.api.login === "function"
  );
}

async function loginViaPywebview(user, pass) {
  try {
    const res = await window.pywebview.api.login({ usuario: user, senha: pass });
    const code = Number(res?.code ?? 500);

    return {
      ok: Boolean(res?.ok),
      code,
      erro: res?.erro || (code === 200 ? "" : DEFAULT_AUTH_ERROR),
      data: res || {},
    };
  } catch (err) {
    console.error(err);
    return {
      ok: false,
      code: 0,
      erro: DEFAULT_CONNECTIVITY_ERROR,
      data: {},
    };
  }
}

async function loginViaHttp(user, pass) {
  try {
    const resp = await fetch(API_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({ usuario: user, senha: pass }),
    });

    const data = await resp.json().catch(() => ({}));
    if (!resp.ok) {
      // FastAPI manda {detail: {...}}
      const detail = data.detail ?? data;
      return {
        ok: false,
        code: resp.status,
        erro: detail?.erro || DEFAULT_AUTH_ERROR,
        data,
      };
    }

    return {
      ok: true,
      code: resp.status,
      erro: "",
      data,
    };
  } catch (err) {
    console.error(err);
    return {
      ok: false,
      code: 0,
      erro: DEFAULT_CONNECTIVITY_ERROR,
      data: {},
    };
  }
}

async function authenticate(user, pass) {
  if (hasPywebviewLoginApi()) {
    return loginViaPywebview(user, pass);
  }

  return loginViaHttp(user, pass);
}

form.addEventListener("submit", async (e) => {
  e.preventDefault();

  const user = document.getElementById("user").value.trim();
  const pass = document.getElementById("pass").value;

  errorBox.style.display = "none";
  errorBox.textContent = "";

  if (!user || !pass) {
    showError("Preencha usuário e senha.");
    return;
  }

  const result = await authenticate(user, pass);
  if (!result.ok) {
    showError(result.erro || DEFAULT_AUTH_ERROR);
    return;
  }

  // Login OK (sem token, sem nada)
  try {
    localStorage.setItem("cc_logged", "1");
    window.location.href = "../pages/index.html?tab=active";
  } catch (err) {
    console.error(err);
    showError(DEFAULT_AUTH_ERROR);
  }
});

document.getElementById("clearSession").addEventListener("click", (e) => {
  e.preventDefault();
  localStorage.removeItem("cc_logged");
  showError("Sessão limpa. Faça login novamente.");
});

function showError(msg) {
  errorBox.textContent = msg;
  errorBox.style.display = "block";
}
