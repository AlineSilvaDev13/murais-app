const msalConfig = {
  auth: {
    clientId: CONFIG.clientId,
    authority: `https://login.microsoftonline.com/${CONFIG.tenantId}`,
    redirectUri: CONFIG.redirectUri
  },
  cache: {
    cacheLocation: "localStorage",
    storeAuthStateInCookie: false
  }
};

const msalInstance = new msal.PublicClientApplication(msalConfig);
const loginRequest = { scopes: ["User.Read", "Sites.ReadWrite.All", "Files.ReadWrite.All"] };

let currentAccount = null;
let currentSetor = null;
let siteId = null;
let pedidosListId = null;
let programacaoListId = null;

const loginView = document.getElementById("loginView");
const demandasView = document.getElementById("demandasView");
const loginBtn = document.getElementById("loginBtn");
const logoutBtn = document.getElementById("logoutBtn");
const loginError = document.getElementById("loginError");
const setorNomeEl = document.getElementById("setorNome");
const demandasCardsEl = document.getElementById("demandasCards");
const pdfModal = document.getElementById("pdfModal");
const pdfFrame = document.getElementById("pdfFrame");
const closePdfBtn = document.getElementById("closePdfBtn");

loginBtn.addEventListener("click", handleLogin);
logoutBtn.addEventListener("click", handleLogout);
closePdfBtn.addEventListener("click", closePdfModal);

async function getToken() {
  const accounts = msalInstance.getAllAccounts();
  const account = accounts[0];
  if (!account) throw new Error("Nenhuma conta autenticada");

  const request = { ...loginRequest, account };
  try {
    const result = await msalInstance.acquireTokenSilent(request);
    return result.accessToken;
  } catch (err) {
    await msalInstance.acquireTokenRedirect(request);
  }
}

async function graphFetch(path, options = {}) {
  const token = await getToken();
  const url = path.startsWith("https://") ? path : `https://graph.microsoft.com/v1.0${path}`;
  const response = await fetch(url, {
    ...options,
    headers: {
      Authorization: `Bearer ${token}`,
      Prefer: "HonorNonIndexedQueriesWarningMayFailRandomly",
      ...(options.headers || {})
    }
  });
  if (!response.ok) {
    const text = await response.text();
    throw new Error(`Graph error ${response.status}: ${text}`);
  }
  return response;
}

async function handleLogin() {
  loginError.hidden = true;
  try {
    await msalInstance.loginRedirect(loginRequest);
  } catch (err) {
    showLoginError("Falha ao entrar com Microsoft. Tente novamente.");
    console.error(err);
  }
}

function handleLogout() {
  msalInstance.logoutRedirect();
}

function showLoginError(message) {
  loginError.textContent = message;
  loginError.hidden = false;
}

async function afterLogin() {
  const email = (currentAccount.username || "").toLowerCase();
  const setor = CONFIG.acessoPorSetor[email];

  if (!setor) {
    showLoginError("Este e-mail não tem um setor configurado");
    return;
  }

  currentSetor = setor;
  setorNomeEl.textContent = setor;

  loginView.hidden = true;
  demandasView.hidden = false;

  await resolveSiteAndLists();
  await carregarDemandas();
}

async function resolveSiteAndLists() {
  const siteRes = await graphFetch(`/sites/${CONFIG.hostname}:${CONFIG.sitePath}`);
  const site = await siteRes.json();
  siteId = site.id;

  const listsRes = await graphFetch(`/sites/${siteId}/lists?$select=id,displayName`);
  const lists = (await listsRes.json()).value;

  const pedidosList = lists.find(l => l.displayName === CONFIG.pedidosListDisplayName);
  const programacaoList = lists.find(l => l.displayName === CONFIG.programacaoListDisplayName);

  if (!pedidosList || !programacaoList) {
    throw new Error("Não foi possível localizar as listas do SharePoint");
  }

  pedidosListId = pedidosList.id;
  programacaoListId = programacaoList.id;
}

async function carregarDemandas() {
  demandasCardsEl.innerHTML = "";

  const filter = `fields/Setor eq '${currentSetor}'`;
  const programacaoRes = await graphFetch(
    `/sites/${siteId}/lists/${programacaoListId}/items?$expand=fields&$filter=${encodeURIComponent(filter)}`
  );
  const registros = (await programacaoRes.json()).value;

  // Transição automática: registros "Programado" cuja DataProgramada já chegou
  // (hoje ou antes) viram "Em Andamento" antes de filtrar o que será exibido.
  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);

  for (const registro of registros) {
    if (registro.fields.Status === "Programado" && registro.fields.DataProgramada) {
      const dataProgramada = new Date(registro.fields.DataProgramada);
      dataProgramada.setHours(0, 0, 0, 0);
      if (dataProgramada <= hoje) {
        await graphFetch(`/sites/${siteId}/lists/${programacaoListId}/items/${registro.id}/fields`, {
          method: "PATCH",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ Status: "Em Andamento" })
        });
        registro.fields.Status = "Em Andamento";
      }
    }
  }

  const programacaoItems = registros.filter((registro) => registro.fields.Status === "Em Andamento");

  if (programacaoItems.length === 0) {
    demandasCardsEl.innerHTML = '<p class="empty-message">Nenhuma demanda em andamento para o seu setor.</p>';
    return;
  }

  for (const item of programacaoItems) {
    const pedidoLookupId = item.fields.PedidoLookupId;
    if (!pedidoLookupId) continue;

    try {
      const pedidoRes = await graphFetch(`/sites/${siteId}/lists/${pedidosListId}/items/${pedidoLookupId}?$expand=fields`);
      const pedido = (await pedidoRes.json()).fields;
      renderCard(pedido, item.fields);
    } catch (err) {
      console.error("Erro ao buscar pedido", pedidoLookupId, err);
    }
  }
}

function renderCard(pedido, programacaoFields) {
  const f = CONFIG.fields;
  const card = document.createElement("article");
  card.className = "demanda-card";

  const numero = pedido[f.title] || "-";
  const cliente = pedido[f.cliente] || "-";
  const categoria = pedido[f.categoria2] || "-";
  const dataFabrica = pedido[f.dataFabrica];

  let prazoHtml = "";
  if (dataFabrica) {
    const prazo = new Date(dataFabrica);
    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);
    const prazoSemHora = new Date(prazo);
    prazoSemHora.setHours(0, 0, 0, 0);
    const atrasado = prazoSemHora < hoje;

    prazoHtml = `Prazo: ${prazo.toLocaleDateString("pt-BR")}`;
    if (atrasado) {
      prazoHtml += ' <span class="card-atrasado">• Atrasado</span>';
    }
  }

  card.innerHTML = `
    <div class="card-header">
      <span class="card-numero">#${numero}</span>
      <span class="card-categoria">${categoria}</span>
    </div>
    <div class="card-body">
      <p class="card-cliente">${cliente}</p>
      ${prazoHtml ? `<p class="card-prazo">${prazoHtml}</p>` : ""}
    </div>
    <div class="card-footer">
      <button class="btn-ver-pdf">Ver PDF</button>
    </div>
  `;

  const btnPdf = card.querySelector(".btn-ver-pdf");
  btnPdf.addEventListener("click", () => abrirPdf(pedido));

  demandasCardsEl.appendChild(card);
}

async function abrirPdf(pedido) {
  const f = CONFIG.fields;
  const linkField = pedido[f.url];
  if (!linkField) {
    alert("Este pedido não possui PDF vinculado.");
    return;
  }

  try {
    const itemId = typeof linkField === "object" ? linkField.Id || linkField.id : linkField;
    const contentRes = await graphFetch(
      `/sites/${siteId}/lists/${pedidosListId}/items/${itemId}/driveItem/content`
    );
    const blob = await contentRes.blob();
    const blobUrl = URL.createObjectURL(blob);
    pdfFrame.src = blobUrl;
    pdfModal.hidden = false;
  } catch (err) {
    console.error(err);
    alert("Não foi possível carregar o PDF.");
  }
}

function closePdfModal() {
  pdfModal.hidden = true;
  if (pdfFrame.src) {
    URL.revokeObjectURL(pdfFrame.src);
    pdfFrame.src = "";
  }
}

(async function init() {
  const redirectResponse = await msalInstance.handleRedirectPromise();
  if (redirectResponse) {
    currentAccount = redirectResponse.account;
    msalInstance.setActiveAccount(currentAccount);
  }

  const accounts = msalInstance.getAllAccounts();
  if (accounts.length > 0) {
    currentAccount = accounts[0];
    msalInstance.setActiveAccount(currentAccount);
    try {
      await afterLogin();
    } catch (err) {
      console.error(err);
      showLoginError("Erro ao carregar demandas. Tente novamente.");
    }
  }
})();
