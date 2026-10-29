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
    const okTipo = tipo === "todos" || card.dataset.tipos.split(" ").includes(tipo);
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
  // "Programado" entra na busca de todos: pra um setor, o pedido aparece a partir da DataProgramada (ver carregarDemandas).
  // Acesso "Todos" carrega todos os setores e filtra no navegador.
  const statusBuscados = "(fields/Status eq 'Em Andamento' or fields/Status eq 'Parado' or fields/Status eq 'Programado')";
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

// Cada pedido tem um registro por etapa/setor, quase todos "Programado" pro futuro: só entram no mural
// os que estão em andamento/parados, os Programado cuja data já chegou e, pro acesso Todos, os Programado
// de setores fora da fábrica (que não têm mural próprio).
function deveExibirNoMural(registro) {
  const { Status: status, Setor: setor, DataProgramada: dataProgramada } = registro.fields;
  if (status === "Em Andamento" || status === "Parado") return true;
  if (status !== "Programado" || !acessoTodos) return false;
  if (!CONFIG.setores.includes(setor)) return true;
  if (!dataProgramada) return false;

  const data = new Date(dataProgramada);
  data.setHours(0, 0, 0, 0);
  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);
  return data <= hoje;
}

async function carregarDemandas() {
  demandasCardsEl.innerHTML = "";

  const registros = await buscarRegistros();

  if (acessoTodos) {
    preencherSetores(registros);
  }

  // Mural de setor: registros "Programado" cuja DataProgramada chegou (hoje ou qualquer dia antes)
  // viram "Em Andamento" e continuam aparecendo nos dias seguintes. Não bloqueia o mural se a gravação falhar.
  if (!acessoTodos) {
    const hoje = new Date();
    hoje.setHours(0, 0, 0, 0);

    const aIniciar = registros.filter((registro) => {
      if (registro.fields.Status !== "Programado" || !registro.fields.DataProgramada) return false;
      const dataProgramada = new Date(registro.fields.DataProgramada);
      dataProgramada.setHours(0, 0, 0, 0);
      return dataProgramada <= hoje;
    });

    await Promise.all(aIniciar.map(async (registro) => {
      registro.fields.Status = "Em Andamento";
      try {
        await graphFetch(`/sites/${siteId}/lists/${programacaoListId}/items/${registro.id}/fields`, {
          method: "PATCH",
          headers: { "Content-Type": "application/json" },
          body: JSON.stringify({ Status: "Em Andamento" })
        });
      } catch (err) {
        console.error("Não foi possível gravar a mudança para Em Andamento", registro.id, err);
      }
    }));
  }

  const programacaoItems = registros.filter((registro) => deveExibirNoMural(registro));

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

  // Pedidos do mesmo GrupoPedido (2 ou mais no mesmo setor) viram UM card só, na posição do mais urgente,
  // com botão pra passar de um PDF pro outro. Um pedido sozinho num grupo é um card normal.
  const unidades = [];
  const gruposPorChave = new Map();
  demandas.forEach((demanda) => {
    const nomeGrupo = String(demanda.pedido[f.grupo] || "").trim();
    if (!nomeGrupo) {
      unidades.push({ nomeGrupo: "", membros: [demanda] });
      return;
    }
    const chave = `${nomeGrupo}||${demanda.item.fields.Setor || ""}`;
    if (!gruposPorChave.has(chave)) {
      const unidade = { nomeGrupo, membros: [] };
      gruposPorChave.set(chave, unidade);
      unidades.push(unidade);
    }
    gruposPorChave.get(chave).membros.push(demanda);
  });

  unidades.forEach((unidade, index) => {
    renderCard(unidade, index + 1);
  });

  if (acessoTodos) {
    aplicarFiltros();
  }
}

function preencherPrazo(el, dataFabrica) {
  el.textContent = "";
  el.hidden = !dataFabrica;
  if (!dataFabrica) return;

  const prazo = new Date(dataFabrica);
  const hoje = new Date();
  hoje.setHours(0, 0, 0, 0);
  const prazoSemHora = new Date(prazo);
  prazoSemHora.setHours(0, 0, 0, 0);

  el.textContent = `Prazo: ${prazo.toLocaleDateString("pt-BR")}`;
  if (prazoSemHora < hoje) {
    const atrasado = document.createElement("span");
    atrasado.className = "card-atrasado";
    atrasado.textContent = " • Atrasado";
    el.appendChild(atrasado);
  }
}

function renderCard(unidade, prioridade) {
  const f = CONFIG.fields;
  const { membros, nomeGrupo } = unidade;
  const itens = membros.map((m) => m.item);
  const primeiro = itens[0];
  const emGrupo = membros.length >= 2;

  const card = document.createElement("article");
  card.className = "demanda-card";

  const travado = itens.some((i) => i.fields.Status === "Parado");
  const programado = itens.every((i) => i.fields.Status === "Programado");
  const motivo = (itens.find((i) => i.fields.Status === "Parado") || primeiro).fields.MotivoParada || "";
  const dataProgramada = primeiro.fields.DataProgramada
    ? new Date(primeiro.fields.DataProgramada).toLocaleDateString("pt-BR")
    : "";

  const tipos = new Set(membros.map((m) => (m.pedido[f.categoria2] === "AT" ? "AT" : "normal")));
  card.dataset.setor = primeiro.fields.Setor || "";
  card.dataset.tipos = [...tipos].join(" ");
  card.dataset.busca = normalizarTexto(
    membros.map((m) => `${m.pedido[f.title] || ""} ${m.pedido[f.cliente] || ""}`).join(" ") + (emGrupo ? ` ${nomeGrupo}` : "")
  );

  card.innerHTML = `
    <div class="card-travado-banner" hidden>TRAVADO: <span class="card-travado-motivo"></span></div>
    ${programado ? `<div class="card-programado-banner">PROGRAMADO${dataProgramada ? ` PARA ${dataProgramada}` : ""}</div>` : ""}
    <div class="card-header">
      <span class="badge-prioridade">${prioridade}</span>
      <span class="card-numero"></span>
      <span class="badge-categoria"></span>
      ${emGrupo ? `<span class="badge-grupo"></span>` : ""}
      ${acessoTodos ? `<span class="card-setor"></span>` : ""}
    </div>
    <div class="card-body">
      <p class="card-cliente"></p>
      <p class="card-prazo"></p>
    </div>
    ${emGrupo ? `
    <div class="grupo-nav">
      <button class="grupo-nav-btn grupo-anterior">‹ Pedido anterior</button>
      <span class="grupo-indicador"></span>
      <button class="grupo-nav-btn grupo-proximo">Próximo pedido ›</button>
    </div>` : ""}
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
    const banner = card.querySelector(".card-travado-banner");
    banner.hidden = false;
    banner.querySelector(".card-travado-motivo").textContent = motivo;
  }

  if (acessoTodos) card.querySelector(".card-setor").textContent = primeiro.fields.Setor || "";
  if (emGrupo) card.querySelector(".badge-grupo").textContent = `Grupo ${nomeGrupo} · ${membros.length} pedidos`;

  card.indiceMembro = 0;
  card.pdfSeq = 0;

  // Mostra os dados do pedido da vez (número, categoria, cliente, prazo) e, se pedido, carrega o PDF dele.
  function mostrarMembro(indice, carregarPdf) {
    const { pedido } = membros[indice];
    const categoria = pedido[f.categoria2] || "-";
    const badge = card.querySelector(".badge-categoria");

    card.indiceMembro = indice;
    card.querySelector(".card-numero").textContent = `#${pedido[f.title] || "-"}`;
    badge.textContent = categoria;
    badge.classList.toggle("badge-at", categoria === "AT");
    card.querySelector(".card-cliente").textContent = pedido[f.cliente] || "-";
    preencherPrazo(card.querySelector(".card-prazo"), pedido[f.dataFabrica]);

    if (emGrupo) {
      card.querySelector(".grupo-indicador").textContent = `Pedido ${indice + 1} de ${membros.length}`;
      card.querySelector(".grupo-anterior").disabled = indice === 0;
      card.querySelector(".grupo-proximo").disabled = indice === membros.length - 1;
    }

    if (carregarPdf) {
      card.pdfSeq++;
      card.querySelector(".card-pdf-wrap").innerHTML = '<div class="pdf-placeholder">Carregando PDF...</div>';
      carregarPdfDoCard(pedido, card);
    }
  }

  mostrarMembro(0, false);

  if (emGrupo) {
    card.querySelector(".grupo-anterior").onclick = () => mostrarMembro(card.indiceMembro - 1, true);
    card.querySelector(".grupo-proximo").onclick = () => mostrarMembro(card.indiceMembro + 1, true);
  }

  card.querySelector(".btn-finalizar").addEventListener("click", () => handleFinalizar(itens, card));

  const btnTravar = card.querySelector(".btn-travar, .btn-destravar");
  if (btnTravar) {
    vincularBotaoTravar(btnTravar, itens, card, travado);
  }

  card.carregarPdf = () => carregarPdfDoCard(membros[card.indiceMembro].pedido, card);
  demandasCardsEl.appendChild(card);
  observadorPdf.observe(card);
}

function vincularBotaoTravar(btn, itens, card, travado) {
  if (travado) {
    btn.className = "btn-destravar";
    btn.textContent = "🔓 DESTRAVAR";
    btn.onclick = () => handleDestravar(itens, card, btn);
  } else {
    btn.className = "btn-travar";
    btn.textContent = "🔒 TRAVAR";
    btn.onclick = () => handleTravar(itens, card, btn);
  }
}

async function configurarVisualizadorPdf(blob, containerEl, aindaValido = () => true) {
  const arrayBuffer = await blob.arrayBuffer();
  const pdf = await pdfjsLib.getDocument({ data: arrayBuffer }).promise;
  if (!aindaValido()) return;
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
  const seq = card.pdfSeq;
  const aindaValido = () => card.pdfSeq === seq;

  if (!linkField) {
    placeholder.textContent = "Este pedido não possui PDF vinculado.";
    return;
  }

  try {
    const itemId = typeof linkField === "object" ? linkField.Id || linkField.id : linkField;
    const blob = await obterPdfBlob(itemId);
    if (!aindaValido()) return;
    await configurarVisualizadorPdf(blob, wrap, aindaValido);
  } catch (err) {
    console.error(err);
    placeholder.textContent = "Não foi possível carregar o PDF.";
  }
}

// Num card de grupo, a ação vale pra todos os pedidos do grupo (a programação deles é a mesma).
async function atualizarRegistros(itens, campos) {
  const resultados = await Promise.allSettled(itens.map((item) =>
    graphFetch(`/sites/${siteId}/lists/${programacaoListId}/items/${item.id}/fields`, {
      method: "PATCH",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(campos)
    }).then(() => Object.assign(item.fields, campos))
  ));

  const falhas = resultados.filter((r) => r.status === "rejected");
  if (falhas.length > 0) {
    falhas.forEach((r) => console.error(r.reason));
    throw new Error(`${falhas.length} de ${itens.length} registros não foram atualizados`);
  }
}

async function handleFinalizar(itens, card) {
  const codigo = window.prompt("Digite o código de confirmação:");
  if (codigo === null) return;
  if (codigo !== "000") {
    alert("Código incorreto");
    return;
  }

  try {
    await atualizarRegistros(itens, {
      Status: "Concluído",
      DataConclusao: new Date().toISOString()
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

async function handleTravar(itens, card, btn) {
  const codigo = window.prompt("Digite o código de confirmação:");
  if (codigo === null) return;
  if (codigo !== "000") {
    alert("Código incorreto");
    return;
  }

  const motivo = window.prompt("Motivo da parada:");
  if (motivo === null) return;

  try {
    await atualizarRegistros(itens, { Status: "Parado", MotivoParada: motivo });

    card.classList.add("card-travado");
    const banner = card.querySelector(".card-travado-banner");
    banner.hidden = false;
    banner.querySelector(".card-travado-motivo").textContent = motivo;

    vincularBotaoTravar(btn, itens, card, true);
  } catch (err) {
    console.error(err);
    alert("Não foi possível travar o pedido.");
  }
}

async function handleDestravar(itens, card, btn) {
  const codigo = window.prompt("Digite o código:");
  if (codigo === null) return;
  if (codigo !== "000") {
    alert("Código incorreto");
    return;
  }

  try {
    await atualizarRegistros(itens, { Status: "Em Andamento", MotivoParada: "" });

    card.classList.remove("card-travado");
    const banner = card.querySelector(".card-travado-banner");
    banner.hidden = true;
    banner.querySelector(".card-travado-motivo").textContent = "";

    vincularBotaoTravar(btn, itens, card, false);
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
