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
let acessoTodos = false;
let siteId = null;
let pedidosListId = null;
let programacaoListId = null;
let documentosListId = null;

const loginView = document.getElementById("loginView");
const demandasView = document.getElementById("demandasView");
const loginBtn = document.getElementById("loginBtn");
const logoutBtn = document.getElementById("logoutBtn");
const loginError = document.getElementById("loginError");
const setorNomeEl = document.getElementById("setorNome");
const setorFiltroEl = document.getElementById("setorFiltro");
const filtrosAdminEl = document.getElementById("filtrosAdmin");
const buscaPedidoEl = document.getElementById("buscaPedido");
const filtroTipoEl = document.getElementById("filtroTipo");
const contadorPedidosEl = document.getElementById("contadorPedidos");
const semResultadosEl = document.getElementById("semResultados");
const demandasCardsEl = document.getElementById("demandasCards");

const observadorPdf = new IntersectionObserver((entradas, observador) => {
  entradas.forEach((entrada) => {
    if (entrada.isIntersecting) {
      observador.unobserve(entrada.target);
      entrada.target.carregarPdf();
    }
  });
}, { rootMargin: "600px 0px" });

loginBtn.addEventListener("click", handleLogin);
logoutBtn.addEventListener("click", handleLogout);

let tokenEmMemoria = null;
let tokenExpiraEm = 0;
let tokenEmAndamento = null;

async function getToken() {
  if (tokenEmMemoria && Date.now() < tokenExpiraEm - 60000) return tokenEmMemoria;
  if (tokenEmAndamento) return tokenEmAndamento;

  const accounts = msalInstance.getAllAccounts();
  const account = accounts[0];
  if (!account) throw new Error("Nenhuma conta autenticada");

  const request = { ...loginRequest, account };
  tokenEmAndamento = (async () => {
    try {
      const result = await msalInstance.acquireTokenSilent(request);
      tokenEmMemoria = result.accessToken;
      tokenExpiraEm = result.expiresOn ? result.expiresOn.getTime() : Date.now() + 5 * 60000;
      return tokenEmMemoria;
    } catch (err) {
      await msalInstance.acquireTokenRedirect(request);
    } finally {
      tokenEmAndamento = null;
    }
  })();
  return tokenEmAndamento;
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
  acessoTodos = setor === "Todos";

  if (acessoTodos) {
    configurarSeletorDeSetor();
  } else {
    setorNomeEl.textContent = setor;
  }

  loginView.hidden = true;
  demandasView.hidden = false;
  demandasCardsEl.innerHTML = '<p class="empty-message">Carregando pedidos...</p>';

  await iniciarMural();
}

async function iniciarMural() {
  try {
    await resolveSiteAndLists();
    await carregarDemandas();
  } catch (err) {
    // IDs em cache podem ter ficado velhos (lista recriada): limpa e tenta uma vez do zero.
    if (idsVieramDoCache) {
      try { localStorage.removeItem(CACHE_IDS_KEY); } catch (e) { /* sem acesso ao storage */ }
      try {
        await resolveSiteAndLists();
        await carregarDemandas();
        return;
      } catch (err2) {
        err = err2;
      }
    }
    console.error(err);
    mostrarErroNoMural(err);
  }
}

function mostrarErroNoMural(err) {
  const detalhe = String((err && err.message) || err).slice(0, 300);
  demandasCardsEl.innerHTML = "";
  const aviso = document.createElement("div");
  aviso.className = "empty-message";
  aviso.innerHTML = `<p>Não foi possível carregar os pedidos.</p><p class="erro-detalhe"></p><button class="btn-secondary">Tentar novamente</button>`;
  aviso.querySelector(".erro-detalhe").textContent = detalhe;
  aviso.querySelector("button").onclick = () => {
    demandasCardsEl.innerHTML = '<p class="empty-message">Carregando pedidos...</p>';
    iniciarMural();
  };
  demandasCardsEl.appendChild(aviso);
  semResultadosEl.hidden = true;
}

function configurarSeletorDeSetor() {
  preencherSetores([]);
  filtrosAdminEl.hidden = false;
  setorNomeEl.hidden = true;

  setorFiltroEl.onchange = () => {
    currentSetor = setorFiltroEl.value;
    aplicarFiltros();
  };
  buscaPedidoEl.oninput = aplicarFiltros;
  filtroTipoEl.onchange = aplicarFiltros;
}

// Setores da config + qualquer outro setor que apareça nos registros (inclui os que não são de fábrica).
function preencherSetores(registros) {
  const encontrados = registros.map((r) => r.fields.Setor).filter(Boolean);
  const extras = [...new Set(encontrados)]
    .filter((s) => !CONFIG.setores.includes(s))
    .sort((a, b) => a.localeCompare(b, "pt-BR"));
  const setores = CONFIG.setores.concat(extras);

  setorFiltroEl.innerHTML = [`<option value="Todos">Todos os setores</option>`]
    .concat(setores.map((s) => `<option value="${s}">${s}</option>`))
    .join("");
  setorFiltroEl.value = currentSetor;
}

function normalizarTexto(texto) {
  return String(texto || "")
    .normalize("NFD")
    .replace(/[̀-ͯ]/g, "")
    .toLowerCase()
    .trim();
}

function aplicarFiltros() {
  const termo = normalizarTexto(buscaPedidoEl.value);
  const tipo = filtroTipoEl.value;
  let visiveis = 0;

  demandasCardsEl.querySelectorAll(".demanda-card").forEach((card) => {
    const okSetor = currentSetor === "Todos" || card.dataset.setor === currentSetor;
    const okBusca = !termo || card.dataset.busca.includes(termo);
    const okTipo = tipo === "todos" || card.dataset.tipo === tipo;
    const mostrar = okSetor && okBusca && okTipo;

    card.hidden = !mostrar;
    if (mostrar) {
      visiveis++;
      card.querySelector(".badge-prioridade").textContent = visiveis;
    }
  });

  contadorPedidosEl.textContent = `${visiveis} pedido(s)`;
  semResultadosEl.hidden = visiveis !== 0;
}

const CACHE_IDS_KEY = "murais_ids_v1";
let idsVieramDoCache = false;

function lerIdsEmCache() {
  try {
    const c = JSON.parse(localStorage.getItem(CACHE_IDS_KEY));
    if (c && c.hostname === CONFIG.hostname && c.sitePath === CONFIG.sitePath
        && c.pedidosNome === CONFIG.pedidosListDisplayName && c.programacaoNome === CONFIG.programacaoListDisplayName
        && c.siteId && c.pedidosListId && c.programacaoListId && c.documentosListId) {
      return c;
    }
  } catch (e) { /* cache ausente ou inválido */ }
  return null;
}

// Site e listas quase nunca mudam: guarda os IDs pra não refazer 4-5 chamadas a cada abertura.
async function resolveSiteAndLists() {
  const cache = lerIdsEmCache();
  if (cache) {
    siteId = cache.siteId;
    pedidosListId = cache.pedidosListId;
    programacaoListId = cache.programacaoListId;
    documentosListId = cache.documentosListId;
    idsVieramDoCache = true;
    return;
  }

  idsVieramDoCache = false;
  const siteRes = await graphFetch(`/sites/${CONFIG.hostname}:${CONFIG.sitePath}`);
  siteId = (await siteRes.json()).id;

  const listas = await graphFetchTodasPaginas(`/sites/${siteId}/lists?$select=id,displayName`);
  const achar = (nome) => {
    const lista = listas.find((l) => l.displayName === nome);
    if (!lista) throw new Error(`Não foi possível localizar a lista "${nome}" no SharePoint`);
    return lista.id;
  };

  pedidosListId = achar(CONFIG.pedidosListDisplayName);
  programacaoListId = achar(CONFIG.programacaoListDisplayName);
  documentosListId = achar("Documentos");

  try {
    localStorage.setItem(CACHE_IDS_KEY, JSON.stringify({
      hostname: CONFIG.hostname,
      sitePath: CONFIG.sitePath,
      pedidosNome: CONFIG.pedidosListDisplayName,
      programacaoNome: CONFIG.programacaoListDisplayName,
      siteId, pedidosListId, programacaoListId, documentosListId
    }));
  } catch (e) { /* sem acesso ao storage */ }
}

async function graphFetchTodasPaginas(path) {
  const itens = [];
  let url = path;
  while (url) {
    const res = await graphFetch(url);
    const json = await res.json();
    itens.push(...json.value);
    url = json["@odata.nextLink"] || null;
  }
  return itens;
}

// Busca vários itens de lista com o endpoint $batch (até 20 por chamada, 3 chamadas em paralelo).
async function buscarPedidosEmLote(ids) {
  const porId = new Map();
  const tamanhoLote = 20;
  const lotes = [];
  for (let i = 0; i < ids.length; i += tamanhoLote) lotes.push(ids.slice(i, i + tamanhoLote));

  async function buscarIndividual(id) {
    try {
      const res = await graphFetch(`/sites/${siteId}/lists/${pedidosListId}/items/${id}?$expand=fields`);
      porId.set(id, (await res.json()).fields);
    } catch (err) {
      console.error("Erro ao buscar pedido", id, err);
    }
  }

  async function processarLote(lote) {
    try {
      const res = await graphFetch("https://graph.microsoft.com/v1.0/$batch", {
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          requests: lote.map((id) => ({
            id: String(id),
            method: "GET",
            url: `/sites/${siteId}/lists/${pedidosListId}/items/${id}?$expand=fields`
          }))
        })
      });
      const { responses } = await res.json();
      const falhas = [];
      const respondidos = new Set();
      for (const r of responses) {
        respondidos.add(r.id);
        if (r.status === 200 && r.body && r.body.fields) {
          porId.set(Number(r.id), r.body.fields);
        } else if (r.status === 429 || r.status >= 500) {
          falhas.push(Number(r.id));
        } else {
          console.error("Erro ao buscar pedido", r.id, r.status, r.body);
        }
      }
      lote.filter((id) => !respondidos.has(String(id))).forEach((id) => falhas.push(id));
      await Promise.all(falhas.map(buscarIndividual));
    } catch (err) {
      console.warn("Falha no $batch, buscando individualmente", err);
      await Promise.all(lote.map(buscarIndividual));
    }
  }

  const fila = lotes.slice();
  await Promise.all(Array.from({ length: Math.min(3, fila.length) }, async () => {
    while (fila.length) await processarLote(fila.shift());
  }));

  return porId;
}

async function buscarPedidos(programacaoItems) {
  // O mesmo pedido costuma ter um registro por setor: busca cada pedido uma vez só.
  const ids = [...new Set(programacaoItems.map((i) => Number(i.fields.PedidoLookupId)).filter(Boolean))];
  const pedidos = await buscarPedidosEmLote(ids);

  return programacaoItems
    .map((item) => ({ pedido: pedidos.get(Number(item.fields.PedidoLookupId)), item }))
    .filter((d) => d.pedido);
}

async function buscarRegistros() {
  // Acesso "Todos" carrega todos os setores (inclusive os "Programado", ainda com data marcada) e filtra no navegador.
  const statusBuscados = acessoTodos
    ? "(fields/Status eq 'Em Andamento' or fields/Status eq 'Parado' or fields/Status eq 'Programado')"
    : "(fields/Status eq 'Em Andamento' or fields/Status eq 'Parado')";
  const base = `/sites/${siteId}/lists/${programacaoListId}/items?$expand=fields&$filter=`;
  const consulta = (filtro) => graphFetchTodasPaginas(base + encodeURIComponent(filtro));

  if (!acessoTodos) {
    return consulta(`fields/Setor eq '${currentSetor}' and ${statusBuscados}`);
  }

  try {
    return await consulta(statusBuscados);
  } catch (err) {
    // Consulta sem filtro de setor pode estourar o limite de lista grande do SharePoint: busca setor a setor.
    console.warn("Consulta geral falhou, buscando setor a setor", err);
    const grupos = await Promise.all(
      CONFIG.setores.map((setor) => consulta(`fields/Setor eq '${setor}' and ${statusBuscados}`))
    );
    return grupos.flat();
  }
}

async function carregarDemandas() {
  demandasCardsEl.innerHTML = "";

  const registros = await buscarRegistros();

  if (acessoTodos) {
    preencherSetores(registros);
  }

  // Transição automática: registros "Programado" cuja DataProgramada já chegou
  // (hoje ou antes) viram "Em Andamento" antes de filtrar o que será exibido.
  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);

  for (const registro of registros) {
    if (!acessoTodos && registro.fields.Status === "Programado" && registro.fields.DataProgramada) {
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

  const statusExibidos = acessoTodos ? ["Em Andamento", "Parado", "Programado"] : ["Em Andamento", "Parado"];
  const programacaoItems = registros.filter((registro) => statusExibidos.includes(registro.fields.Status));

  if (programacaoItems.length === 0 && !acessoTodos) {
    demandasCardsEl.innerHTML = '<p class="empty-message">Nenhuma demanda em andamento para o seu setor.</p>';
    return;
  }

  const f = CONFIG.fields;
  const demandas = await buscarPedidos(programacaoItems);

  demandas.sort((a, b) => {
    const dataA = a.pedido[f.dataFabrica] ? new Date(a.pedido[f.dataFabrica]).getTime() : Infinity;
    const dataB = b.pedido[f.dataFabrica] ? new Date(b.pedido[f.dataFabrica]).getTime() : Infinity;
    return dataA - dataB;
  });

  demandas.forEach((demanda, index) => {
    renderCard(demanda.pedido, demanda.item, index + 1);
  });

  if (acessoTodos) {
    aplicarFiltros();
  }
}

function renderCard(pedido, item, prioridade) {
  const f = CONFIG.fields;
  const card = document.createElement("article");
  card.className = "demanda-card";

  const numero = pedido[f.title] || "-";
  const cliente = pedido[f.cliente] || "-";
  const categoria = pedido[f.categoria2] || "-";
  const dataFabrica = pedido[f.dataFabrica];

  const isAt = categoria === "AT";
  const badgeTexto = categoria && categoria !== "-" ? categoria : "-";

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

  const travado = item.fields.Status === "Parado";
  const programado = item.fields.Status === "Programado";
  const dataProgramada = item.fields.DataProgramada
    ? new Date(item.fields.DataProgramada).toLocaleDateString("pt-BR")
    : "";

  const mostrarSetor = acessoTodos;

  card.dataset.setor = item.fields.Setor || "";
  card.dataset.tipo = isAt ? "AT" : "normal";
  card.dataset.busca = normalizarTexto(`${numero} ${cliente}`);

  card.innerHTML = `
    <div class="card-travado-banner" ${travado ? "" : "hidden"}>TRAVADO: <span class="card-travado-motivo">${item.fields.MotivoParada || ""}</span></div>
    ${programado ? `<div class="card-programado-banner">PROGRAMADO${dataProgramada ? ` PARA ${dataProgramada}` : ""}</div>` : ""}
    <div class="card-header">
      <span class="badge-prioridade">${prioridade}</span>
      <span class="card-numero">#${numero}</span>
      <span class="badge-categoria ${isAt ? "badge-at" : ""}">${badgeTexto}</span>
      ${mostrarSetor ? `<span class="card-setor">${item.fields.Setor || ""}</span>` : ""}
    </div>
    <div class="card-body">
      <p class="card-cliente">${cliente}</p>
      ${prazoHtml ? `<p class="card-prazo">${prazoHtml}</p>` : ""}
    </div>
    <div class="card-pdf-wrap">
      <div class="pdf-placeholder">Carregando PDF...</div>
    </div>
    <div class="card-acoes">
      <button class="btn-finalizar">✓ FINALIZAR</button>
      ${programado ? "" : `<button class="${travado ? "btn-destravar" : "btn-travar"}">${travado ? "🔓 DESTRAVAR" : "🔒 TRAVAR"}</button>`}
    </div>
  `;

  if (travado) {
    card.classList.add("card-travado");
  }

  card.querySelector(".btn-finalizar").addEventListener("click", () => handleFinalizar(item, card));

  const btnTravar = card.querySelector(".btn-travar, .btn-destravar");
  if (btnTravar) {
    vincularBotaoTravar(btnTravar, item, card, travado);
  }

  card.carregarPdf = () => carregarPdfDoCard(pedido, card);
  demandasCardsEl.appendChild(card);
  observadorPdf.observe(card);
}

function vincularBotaoTravar(btn, item, card, travado) {
  if (travado) {
    btn.className = "btn-destravar";
    btn.textContent = "🔓 DESTRAVAR";
    btn.onclick = () => handleDestravar(item, card, btn);
  } else {
    btn.className = "btn-travar";
    btn.textContent = "🔒 TRAVAR";
    btn.onclick = () => handleTravar(item, card, btn);
  }
}

async function configurarVisualizadorPdf(blob, containerEl) {
  const arrayBuffer = await blob.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  let paginaAtual = 1;

  containerEl.innerHTML = "";

  const canvas = document.createElement("canvas");
  canvas.className = "card-pdf-canvas";
  containerEl.appendChild(canvas);

  const navegacao = document.createElement("div");
  navegacao.className = "pdf-navegacao";
  navegacao.innerHTML = `
    <button class="pdf-nav-btn pdf-anterior">‹</button>
    <span class="pdf-pagina-indicador">1 / ${pdf.numPages}</span>
    <button class="pdf-nav-btn pdf-proxima">›</button>
  `;
  containerEl.appendChild(navegacao);

  const indicador = navegacao.querySelector(".pdf-pagina-indicador");
  const btnAnterior = navegacao.querySelector(".pdf-anterior");
  const btnProxima = navegacao.querySelector(".pdf-proxima");

  async function desenharPagina(numero) {
    const page = await pdf.getPage(numero);
    const containerWidth = containerEl.clientWidth;
    const viewportOriginal = page.getViewport({ scale: 1 });
    const escala = containerWidth / viewportOriginal.width;
    const viewport = page.getViewport({ scale: escala });
    canvas.width = viewport.width;
    canvas.height = viewport.height;
    const context = canvas.getContext("2d");
    await page.render({ canvasContext: context, viewport: viewport }).promise;
    indicador.textContent = `${numero} / ${pdf.numPages}`;
    btnAnterior.disabled = numero <= 1;
    btnProxima.disabled = numero >= pdf.numPages;
  }

  btnAnterior.addEventListener("click", () => {
    if (paginaAtual > 1) {
      paginaAtual--;
      desenharPagina(paginaAtual);
    }
  });

  btnProxima.addEventListener("click", () => {
    if (paginaAtual < pdf.numPages) {
      paginaAtual++;
      desenharPagina(paginaAtual);
    }
  });

  await desenharPagina(paginaAtual);
}

// O mesmo PDF aparece em vários cards (um por setor): baixa uma vez só, guardando os últimos 30.
const pdfBlobCache = new Map();

function obterPdfBlob(itemId) {
  if (!pdfBlobCache.has(itemId)) {
    const promessa = graphFetch(`/sites/${siteId}/lists/${documentosListId}/items/${itemId}/driveItem/content`)
      .then((res) => res.blob())
      .catch((err) => {
        pdfBlobCache.delete(itemId);
        throw err;
      });
    pdfBlobCache.set(itemId, promessa);
    if (pdfBlobCache.size > 30) {
      pdfBlobCache.delete(pdfBlobCache.keys().next().value);
    }
  }
  return pdfBlobCache.get(itemId);
}

async function carregarPdfDoCard(pedido, card) {
  const f = CONFIG.fields;
  const linkField = pedido[f.url];
  const wrap = card.querySelector(".card-pdf-wrap");
  const placeholder = card.querySelector(".pdf-placeholder");

  if (!linkField) {
    placeholder.textContent = "Este pedido não possui PDF vinculado.";
    return;
  }

  try {
    const itemId = typeof linkField === "object" ? linkField.Id || linkField.id : linkField;
    const blob = await obterPdfBlob(itemId);
    await configurarVisualizadorPdf(blob, wrap);
  } catch (err) {
    console.error(err);
    placeholder.textContent = "Não foi possível carregar o PDF.";
  }
}

async function handleFinalizar(item, card) {
  const codigo = window.prompt("Digite o código de confirmação:");
  if (codigo === null) return;
  if (codigo !== "000") {
    alert("Código incorreto");
    return;
  }

  try {
    await graphFetch(`/sites/${siteId}/lists/${programacaoListId}/items/${item.id}/fields`, {
      method: "PATCH",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        Status: "Concluído",
        DataConclusao: new Date().toISOString()
      })
    });
    card.remove();
    if (acessoTodos) {
      aplicarFiltros();
    }
  } catch (err) {
    console.error(err);
    alert("Não foi possível finalizar o pedido.");
  }
}

async function handleTravar(item, card, btn) {
  const codigo = window.prompt("Digite o código de confirmação:");
  if (codigo === null) return;
  if (codigo !== "000") {
    alert("Código incorreto");
    return;
  }

  const motivo = window.prompt("Motivo da parada:");
  if (motivo === null) return;

  try {
    await graphFetch(`/sites/${siteId}/lists/${programacaoListId}/items/${item.id}/fields`, {
      method: "PATCH",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        Status: "Parado",
        MotivoParada: motivo
      })
    });

    item.fields.Status = "Parado";
    item.fields.MotivoParada = motivo;

    card.classList.add("card-travado");
    const banner = card.querySelector(".card-travado-banner");
    banner.hidden = false;
    banner.querySelector(".card-travado-motivo").textContent = motivo;

    vincularBotaoTravar(btn, item, card, true);
  } catch (err) {
    console.error(err);
    alert("Não foi possível travar o pedido.");
  }
}

async function handleDestravar(item, card, btn) {
  const codigo = window.prompt("Digite o código:");
  if (codigo === null) return;
  if (codigo !== "000") {
    alert("Código incorreto");
    return;
  }

  try {
    await graphFetch(`/sites/${siteId}/lists/${programacaoListId}/items/${item.id}/fields`, {
      method: "PATCH",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        Status: "Em Andamento",
        MotivoParada: ""
      })
    });

    item.fields.Status = "Em Andamento";
    item.fields.MotivoParada = "";

    card.classList.remove("card-travado");
    const banner = card.querySelector(".card-travado-banner");
    banner.hidden = true;
    banner.querySelector(".card-travado-motivo").textContent = "";

    vincularBotaoTravar(btn, item, card, false);
  } catch (err) {
    console.error(err);
    alert("Não foi possível destravar o pedido.");
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
