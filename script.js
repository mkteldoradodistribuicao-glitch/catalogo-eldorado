let produtos = [];
let produtosFiltrados = [];
let cotacao = [];
let paginaAtual = 1;
let menuAtual = "catalogo";
let ultimasBuscas = JSON.parse(localStorage.getItem("ultimasBuscasCatalogo") || "[]");

const produtosPorPagina = 150;
const codigosOcultosTernura = ["93217", "89817"];
const codigosIgnoradosNovidades = ["999999", "116451"];

const colunas = {
  codigo: 0,
  ean: 1,
  descricao: 2,
  embalagem: 3,
  qtdMaster: 4,
  estoque: 5,
  imagem: 6,
  codFornecedor: 7,
  fornecedor: 8
};

const sinonimosBusca = {
  fh: ["papel", "higienico"],
  ph: ["papel", "higienico"],
  pt: ["papel", "toalha"],
  fg: ["fralda", "geriatrica"],
  fb: ["fralda", "baby"],
  fradla: ["fralda"],
  higienco: ["higienico"],
  higiênico: ["higienico"],
  sabao: ["sabao", "sabão"],
  deterg: ["detergente"],
  amac: ["amaciante"]
};

const $ = id => document.getElementById(id);

function texto(valor) {
  return String(valor ?? "").trim();
}

function normalizar(valor) {
  return texto(valor)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .toLowerCase();
}

function numero(valor) {
  const n = Number(String(valor ?? "0").replace(",", "."));
  return isNaN(n) ? 0 : n;
}

function tokensBusca(busca) {
  const limpo = normalizar(busca).replace(/%/g, " ");
  let tokens = limpo.split(/\s+/).filter(Boolean);
  let expandidos = [];

  tokens.forEach(t => {
    expandidos.push(t);
    if (sinonimosBusca[t]) expandidos.push(...sinonimosBusca[t]);
  });

  return [...new Set(expandidos)];
}

function basePesquisaProduto(p) {
  return normalizar([
    p.codigo,
    p.ean,
    p.descricao,
    p.embalagem,
    p.qtdMaster,
    p.codFornecedor,
    p.fornecedor
  ].join(" "));
}

function camposPesquisa(p) {
  return {
    codigo: normalizar(p.codigo),
    ean: normalizar(p.ean),
    descricao: normalizar(p.descricao),
    embalagem: normalizar(p.embalagem),
    qtdMaster: normalizar(p.qtdMaster),
    codFornecedor: normalizar(p.codFornecedor),
    fornecedor: normalizar(p.fornecedor),
    tudo: basePesquisaProduto(p)
  };
}

fetch("Produtos.xlsx")
  .then(resposta => resposta.arrayBuffer())
  .then(buffer => {
    const workbook = XLSX.read(buffer, { type: "array" });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const linhas = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    produtos = linhas.slice(1).map(linha => ({
      codigo: texto(linha[colunas.codigo]),
      ean: texto(linha[colunas.ean]),
      descricao: texto(linha[colunas.descricao]),
      embalagem: texto(linha[colunas.embalagem]),
      qtdMaster: numero(linha[colunas.qtdMaster]) || 1,
      estoque: numero(linha[colunas.estoque]),
      imagem: texto(linha[colunas.imagem]),
      codFornecedor: texto(linha[colunas.codFornecedor]),
      fornecedor: texto(linha[colunas.fornecedor])
    })).filter(p => p.codigo && p.descricao);

    aplicarFiltros();
  })
  .catch(() => {
    $("mensagemResultado").innerHTML =
      "Erro ao carregar a planilha Produtos.xlsx. Verifique se o arquivo está na raiz do projeto.";
  });

function produtosBase() {
  if (menuAtual === "ternura") {
    return produtos.filter(p =>
      normalizar(p.descricao).includes("ternura") &&
      !codigosOcultosTernura.includes(p.codigo)
    );
  }

  if (menuAtual === "novidades") return obterNovidades();

  return [...produtos];
}

function obterNovidades() {
  return [...produtos]
    .filter(p => !codigosIgnoradosNovidades.includes(p.codigo))
    .sort((a, b) => numero(b.codigo) - numero(a.codigo))
    .slice(0, 150);
}

function ehNovidade(produto) {
  return obterNovidades().some(p => p.codigo === produto.codigo);
}

function aplicarFiltros() {
  let lista = produtosBase();

  const busca = texto($("buscaPrincipal").value);
  const codFornecedor = normalizar($("buscaCodFornecedor").value);
  const fornecedor = normalizar($("buscaFornecedor").value);
  const estoque = $("filtroEstoque").value;

  lista = filtrarBuscaInteligente(lista, busca);

  if (codFornecedor) {
    lista = lista.filter(p => normalizar(p.codFornecedor).includes(codFornecedor));
  }

  if (fornecedor) {
    lista = lista.filter(p => normalizar(p.fornecedor).includes(fornecedor));
  }

  if (estoque === "com") lista = lista.filter(p => p.estoque > 0);
  if (estoque === "sem") lista = lista.filter(p => p.estoque <= 0);

  ordenarProdutos(lista, busca);

  produtosFiltrados = lista;
  paginaAtual = 1;

  renderizarProdutos();
  renderizarNovidades();
  renderizarHistoricoBusca();
}

function filtrarBuscaInteligente(lista, busca) {
  const termoOriginal = texto(busca);
  if (!termoOriginal) return lista;

  const termo = normalizar(termoOriginal);
  const tokens = tokensBusca(termoOriginal);

  const resultados = lista.map(p => {
    const campos = camposPesquisa(p);
    let pontos = 0;
    let encontrados = 0;

    if (campos.codigo === termo) pontos += 1000;
    if (campos.ean === termo) pontos += 1000;
    if (campos.descricao === termo) pontos += 850;
    if (campos.descricao.startsWith(termo)) pontos += 700;
    if (campos.descricao.includes(termo)) pontos += 420;
    if (campos.tudo.includes(termo)) pontos += 180;

    tokens.forEach(t => {
      const termoNumerico = /^\d+$/.test(t);

      if (campos.descricao.includes(t)) {
        pontos += termoNumerico ? 120 : 180;
        encontrados++;
      }

      if (campos.descricao.startsWith(t)) pontos += 180;
      if (campos.codigo.includes(t)) pontos += 160;
      if (campos.ean.includes(t)) pontos += 150;
      if (campos.embalagem.includes(t)) pontos += 80;
      if (campos.qtdMaster === t) pontos += 65;
      if (campos.codFornecedor.includes(t)) pontos += 60;
      if (campos.fornecedor.includes(t)) pontos += 50;

      if (!campos.tudo.includes(t) && termoNumerico) {
        const numeroDescricao = campos.descricao.match(/\d+/g) || [];
        if (numeroDescricao.some(n => n.includes(t))) pontos += 90;
      }
    });

    if (tokens.length && tokens.every(t => campos.tudo.includes(t))) pontos += 350;
    if (encontrados === tokens.length) pontos += 240;
    if (p.estoque > 0) pontos += 20;
    pontos += Math.min(p.estoque / 100, 25);

    return { produto: p, pontos };
  }).filter(x => x.pontos > 0);

  resultados.sort((a, b) => {
    if (b.pontos !== a.pontos) return b.pontos - a.pontos;
    return b.produto.estoque - a.produto.estoque;
  });

  return resultados.map(x => x.produto);
}

function ordenarProdutos(lista, busca) {
  const tipo = $("ordenacao").value;

  if (tipo === "relevancia" && texto(busca)) return;

  if (tipo === "codigoAsc") lista.sort((a, b) => numero(a.codigo) - numero(b.codigo));
  if (tipo === "codigoDesc") lista.sort((a, b) => numero(b.codigo) - numero(a.codigo));
  if (tipo === "descricaoAsc") lista.sort((a, b) => a.descricao.localeCompare(b.descricao));
  if (tipo === "descricaoDesc") lista.sort((a, b) => b.descricao.localeCompare(a.descricao));
  if (tipo === "estoqueDesc") lista.sort((a, b) => b.estoque - a.estoque);
}

function renderizarProdutos() {
  const inicio = (paginaAtual - 1) * produtosPorPagina;
  const fim = inicio + produtosPorPagina;
  const pagina = produtosFiltrados.slice(inicio, fim);

  $("listaProdutos").innerHTML = pagina.map(criarCardProduto).join("");
  renderizarPaginacao();
  renderizarMensagem();
}

function destacarTexto(valor) {
  const busca = texto($("buscaPrincipal").value);
  if (!busca) return valor;

  let resultado = texto(valor);
  const tokens = tokensBusca(busca).filter(t => t.length > 1).slice(0, 8);

  tokens.forEach(t => {
    const re = new RegExp(`(${escaparRegExp(t)})`, "ig");
    resultado = resultado.replace(re, "<mark>$1</mark>");
  });

  return resultado;
}

function escaparRegExp(valor) {
  return valor.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
}

function criarCardProduto(produto) {
  const imagem = produto.imagem ? `Imagens/${produto.imagem}` : "";
  const selo = ehNovidade(produto) ? `<span class="selo-novidade">Novidade</span>` : "";

  return `
    <article class="card" id="card-${produto.codigo}">
      ${selo}
      <div class="imagem-produto">
        ${imagem ? `<img src="${imagem}" alt="${produto.descricao}" onerror="this.style.display='none'">` : ""}
      </div>

      <div class="codigo-destaque">Código: ${destacarTexto(produto.codigo)}</div>

      <h3>${destacarTexto(produto.descricao)}</h3>

      <div class="card-info">
        <div class="info-linha">
          <span class="info-label">EAN</span>
          <span class="info-valor">${destacarTexto(produto.ean)}</span>
        </div>

        <div class="info-linha">
          <span class="info-label">Embalagem</span>
          <span class="info-valor">${destacarTexto(produto.embalagem)}</span>
        </div>

        <div class="info-linha">
          <span class="info-label">QTD Master</span>
          <span class="info-valor">${produto.qtdMaster}</span>
        </div>

        <div class="info-linha">
          <span class="info-label">Cód. Fornecedor</span>
          <span class="info-valor">${destacarTexto(produto.codFornecedor)}</span>
        </div>

        <div class="info-linha">
          <span class="info-label">Fornecedor</span>
          <span class="info-valor fornecedor-valor">${destacarTexto(produto.fornecedor)}</span>
        </div>
      </div>

      <button onclick="adicionarCotacao('${produto.codigo}', this)">
        Adicionar à Cotação
      </button>
    </article>
  `;
}

function renderizarNovidades() {
  if (menuAtual === "novidades") {
    $("areaNovidades").innerHTML = "";
    return;
  }

  const novidades = obterNovidades().slice(0, 12);

  $("areaNovidades").innerHTML = `
    <div class="faixa-novidades">
      <h2>⭐ Novidades</h2>
      <div class="produtos">
        ${novidades.map(criarCardProduto).join("")}
      </div>
    </div>
  `;
}

function renderizarPaginacao() {
  const totalPaginas = Math.max(1, Math.ceil(produtosFiltrados.length / produtosPorPagina));
  const inicio = produtosFiltrados.length === 0 ? 0 : (paginaAtual - 1) * produtosPorPagina + 1;
  const fim = Math.min(paginaAtual * produtosPorPagina, produtosFiltrados.length);

  const html = `
    <button onclick="mudarPagina(${paginaAtual - 1})" ${paginaAtual === 1 ? "disabled" : ""}>Anterior</button>
    <span>Página ${paginaAtual} de ${totalPaginas}</span>
    <button onclick="mudarPagina(${paginaAtual + 1})" ${paginaAtual === totalPaginas ? "disabled" : ""}>Próxima</button>
    <span>Exibindo ${inicio} - ${fim} de ${produtosFiltrados.length}</span>
    <input type="number" min="1" max="${totalPaginas}" value="${paginaAtual}" class="ir-pagina">
    <button onclick="irParaPagina(this)">Ir</button>
  `;

  $("paginacaoTopo").innerHTML = html;
  $("paginacaoBase").innerHTML = html;
}

function mudarPagina(pagina) {
  const totalPaginas = Math.max(1, Math.ceil(produtosFiltrados.length / produtosPorPagina));
  if (pagina < 1 || pagina > totalPaginas) return;

  paginaAtual = pagina;
  renderizarProdutos();

  window.scrollTo({ top: 0, behavior: "smooth" });
}

function irParaPagina(botao) {
  const input = botao.parentElement.querySelector(".ir-pagina");
  mudarPagina(numero(input.value));
}

function renderizarMensagem() {
  if (produtosFiltrados.length > 0) {
    $("mensagemResultado").textContent = "";
    return;
  }

  const busca = texto($("buscaPrincipal").value);
  const fornecedor = texto($("buscaFornecedor").value);
  const codFornecedor = texto($("buscaCodFornecedor").value);

  let mensagem = "Nenhum produto encontrado";

  if (busca) mensagem += ` para "${busca}"`;
  if (fornecedor) mensagem += ` dentro do fornecedor ${fornecedor}`;
  if (codFornecedor) mensagem += ` com código de fornecedor ${codFornecedor}`;

  $("mensagemResultado").textContent = mensagem + ".";
}

function adicionarCotacao(codigo, botao) {
  const produto = produtos.find(p => p.codigo === codigo);
  if (!produto) return;

  const item = cotacao.find(p => p.codigo === codigo);

  if (item) item.quantidade += produto.qtdMaster;
  else cotacao.push({ ...produto, quantidade: produto.qtdMaster });

  botao.textContent = "Adicionado!";
  botao.classList.add("confirmado");

  const card = $(`card-${codigo}`);
  if (card) card.classList.add("destaque-adicionado");

  setTimeout(() => {
    botao.textContent = "Adicionar à Cotação";
    botao.classList.remove("confirmado");
    if (card) card.classList.remove("destaque-adicionado");
  }, 900);

  mostrarToast("Produto adicionado à cotação.");
  renderizarCotacao();
}

function renderizarCotacao() {
  $("contadorTopo").textContent = cotacao.length;
  $("contadorFlutuante").textContent = cotacao.length;

  if (cotacao.length === 0) {
    $("listaCotacao").innerHTML = "<p>Nenhum produto na cotação.</p>";
    return;
  }

  $("listaCotacao").innerHTML = cotacao.map(criarItemCotacao).join("");
}

function criarItemCotacao(item) {
  const restante = item.estoque - item.quantidade;
  const mastersRestantes = Math.floor(restante / item.qtdMaster);

  let aviso = "";

  if (item.quantidade > item.estoque) {
    aviso = `<div class="aviso-estoque erro">❌ Quantidade cotada superior ao estoque disponível.</div>`;
  } else if (mastersRestantes <= 3) {
    aviso = `<div class="aviso-estoque alerta">⚠️ Verifique a disponibilidade do estoque antes de concluir a cotação.</div>`;
  }

  return `
    <div class="item-cotacao">
      <strong>${item.descricao}</strong><br>
      <small>
        Código: ${item.codigo}<br>
        EAN: ${item.ean}<br>
        Embalagem: ${item.embalagem}<br>
        QTD Master: ${item.qtdMaster}
      </small>

      <div class="controle-quantidade">
        <button onclick="alterarQuantidade('${item.codigo}', -1)">−</button>
        <input type="number" value="${item.quantidade}" onchange="alterarQuantidadeManual('${item.codigo}', this.value)">
        <button onclick="alterarQuantidade('${item.codigo}', 1)">+</button>
      </div>

      <small>${item.quantidade / item.qtdMaster} master(s)</small>
      ${aviso}
      <button onclick="removerDaCotacao('${item.codigo}')">Remover</button>
    </div>
  `;
}

function alterarQuantidade(codigo, direcao) {
  const item = cotacao.find(p => p.codigo === codigo);
  if (!item) return;

  item.quantidade += item.qtdMaster * direcao;
  if (item.quantidade <= 0) return removerDaCotacao(codigo);

  renderizarCotacao();
}

function alterarQuantidadeManual(codigo, valor) {
  const item = cotacao.find(p => p.codigo === codigo);
  if (!item) return;

  let quantidade = numero(valor);
  if (quantidade < item.qtdMaster) quantidade = item.qtdMaster;

  item.quantidade = Math.ceil(quantidade / item.qtdMaster) * item.qtdMaster;
  renderizarCotacao();
}

function removerDaCotacao(codigo) {
  cotacao = cotacao.filter(p => p.codigo !== codigo);
  renderizarCotacao();
}

function limparCotacao() {
  if (!cotacao.length) return;
  if (confirm("Deseja limpar toda a lista de cotação?")) {
    cotacao = [];
    renderizarCotacao();
    mostrarToast("Lista de cotação limpa.");
  }
}

function textoCotacao() {
  if (!cotacao.length) return "Nenhum produto na cotação.";

  let textoFinal = "COTAÇÃO ELDORADO\n\n";

  cotacao.forEach((item, index) => {
    textoFinal += `${index + 1}. ${item.descricao}\n`;
    textoFinal += `Código: ${item.codigo}\n`;
    textoFinal += `EAN: ${item.ean}\n`;
    textoFinal += `Embalagem: ${item.embalagem}\n`;
    textoFinal += `QTD Master: ${item.qtdMaster}\n`;
    textoFinal += `Quantidade: ${item.quantidade}\n`;
    textoFinal += `Masters: ${item.quantidade / item.qtdMaster}\n`;
    textoFinal += `Fornecedor: ${item.fornecedor}\n\n`;
  });

  return textoFinal;
}

function copiarCotacao() {
  navigator.clipboard.writeText(textoCotacao());
  mostrarToast("Cotação copiada.");
}

function gerarPDF() {
  if (!cotacao.length) {
    mostrarToast("Nenhum produto na cotação.");
    return;
  }

  const { jsPDF } = window.jspdf;
  const pdf = new jsPDF();
  let y = 15;

  pdf.setFontSize(16);
  pdf.text("Cotação Eldorado", 10, y);
  y += 10;

  pdf.setFontSize(10);

  cotacao.forEach((item, index) => {
    if (y > 270) {
      pdf.addPage();
      y = 15;
    }

    pdf.text(`${index + 1}. ${item.descricao}`, 10, y);
    y += 6;
    pdf.text(`Código: ${item.codigo} | EAN: ${item.ean}`, 10, y);
    y += 6;
    pdf.text(`Embalagem: ${item.embalagem} | QTD Master: ${item.qtdMaster}`, 10, y);
    y += 6;
    pdf.text(`Quantidade: ${item.quantidade} | Masters: ${item.quantidade / item.qtdMaster}`, 10, y);
    y += 6;
    pdf.text(`Fornecedor: ${item.fornecedor}`, 10, y);
    y += 10;
  });

  pdf.save("cotacao-eldorado.pdf");
}

function mostrarToast(mensagem) {
  $("toast").textContent = mensagem;
  $("toast").classList.add("ativo");

  clearTimeout(window.toastTimer);
  window.toastTimer = setTimeout(() => {
    $("toast").classList.remove("ativo");
  }, 2500);
}

function mostrarSugestoesBusca() {
  const busca = texto($("buscaPrincipal").value);

  if (!busca || busca.length < 2) {
    $("sugestoesBusca").innerHTML = "";
    return;
  }

  const sugestoes = filtrarBuscaInteligente(produtosBase(), busca).slice(0, 6);

  $("sugestoesBusca").innerHTML = sugestoes.map(p => `
    <div onclick="selecionarBusca('${p.descricao.replace(/'/g, "\\'")}')">
      <strong>${p.codigo}</strong> - ${p.descricao}
    </div>
  `).join("");
}

function selecionarBusca(valor) {
  $("buscaPrincipal").value = valor;
  $("sugestoesBusca").innerHTML = "";
  salvarHistoricoBusca(valor);
  aplicarFiltros();
}

function mostrarSugestoesFornecedor() {
  const termo = normalizar($("buscaFornecedor").value);

  if (!termo) {
    $("sugestoesFornecedor").innerHTML = "";
    return;
  }

  const mapa = {};

  produtosBase().forEach(p => {
    if (!p.fornecedor) return;

    if (!mapa[p.fornecedor]) {
      mapa[p.fornecedor] = { nome: p.fornecedor, estoque: 0, pontos: 0 };
    }

    mapa[p.fornecedor].estoque += p.estoque;
    if (normalizar(p.fornecedor).includes(termo)) mapa[p.fornecedor].pontos += 10;
  });

  const sugestoes = Object.values(mapa)
    .filter(f => normalizar(f.nome).includes(termo))
    .sort((a, b) => (b.pontos + b.estoque / 1000) - (a.pontos + a.estoque / 1000))
    .slice(0, 5);

  $("sugestoesFornecedor").innerHTML = sugestoes.map(f => `
    <div onclick="selecionarFornecedor('${f.nome.replace(/'/g, "\\'")}')">${f.nome}</div>
  `).join("");
}

function selecionarFornecedor(nome) {
  $("buscaFornecedor").value = nome;
  $("sugestoesFornecedor").innerHTML = "";
  aplicarFiltros();
}

function salvarHistoricoBusca(valor) {
  const v = texto(valor);
  if (!v) return;

  ultimasBuscas = [v, ...ultimasBuscas.filter(x => x !== v)].slice(0, 5);
  localStorage.setItem("ultimasBuscasCatalogo", JSON.stringify(ultimasBuscas));
  renderizarHistoricoBusca();
}

function renderizarHistoricoBusca() {
  if (!ultimasBuscas.length) {
    $("historicoBusca").innerHTML = "";
    return;
  }

  $("historicoBusca").innerHTML = ultimasBuscas.map(b => `
    <button onclick="selecionarBusca('${b.replace(/'/g, "\\'")}')">🔎 ${b}</button>
  `).join("");
}

function trocarMenu(menu) {
  menuAtual = menu;
  document.body.classList.toggle("ternura", menu === "ternura");

  if (menu === "catalogo") {
    $("logoCatalogo").src = "Logos/Eldorado.png";
    $("tituloCatalogo").textContent = "Catálogo Eldorado";
    $("subtituloCatalogo").textContent = "Consulte produtos por código, descrição, EAN, fornecedor ou código do fornecedor.";
    $("seloModo").textContent = "Catálogo Comercial";
  }

  if (menu === "ternura") {
    $("logoCatalogo").src = "Logos/Produtos-ternura.png";
    $("tituloCatalogo").textContent = "Produtos Ternura";
    $("subtituloCatalogo").textContent = "Catálogo exclusivo da linha Produtos Ternura.";
    $("seloModo").textContent = "Linha Exclusiva";
  }

  if (menu === "novidades") {
    $("logoCatalogo").src = "Logos/Eldorado.png";
    $("tituloCatalogo").textContent = "Novidades Eldorado";
    $("subtituloCatalogo").textContent = "Os 150 maiores códigos de produtos, exceto 999999 e 116451.";
    $("seloModo").textContent = "Produtos em Destaque";
  }

  $("menuSuspenso").classList.remove("ativo");
  aplicarFiltros();
}

function limparFiltros() {
  $("buscaPrincipal").value = "";
  $("buscaCodFornecedor").value = "";
  $("buscaFornecedor").value = "";
  $("filtroEstoque").value = "com";
  $("ordenacao").value = "relevancia";
  $("sugestoesBusca").innerHTML = "";
  $("sugestoesFornecedor").innerHTML = "";
  aplicarFiltros();
}

function controlarBotaoTopo() {
  $("btnTopo").style.display = window.scrollY > 300 ? "block" : "none";
}

$("buscaPrincipal").addEventListener("input", () => {
  aplicarFiltros();
  mostrarSugestoesBusca();
});

$("buscaPrincipal").addEventListener("keydown", e => {
  if (e.key === "Enter") {
    salvarHistoricoBusca($("buscaPrincipal").value);
    $("sugestoesBusca").innerHTML = "";
  }
});

$("buscaCodFornecedor").addEventListener("input", aplicarFiltros);

$("buscaFornecedor").addEventListener("input", () => {
  aplicarFiltros();
  mostrarSugestoesFornecedor();
});

$("filtroEstoque").addEventListener("change", aplicarFiltros);
$("ordenacao").addEventListener("change", aplicarFiltros);
$("btnLimparFiltros").addEventListener("click", limparFiltros);

$("btnMenu").addEventListener("click", () => {
  $("menuSuspenso").classList.toggle("ativo");
});

document.querySelectorAll("#menuSuspenso button").forEach(botao => {
  botao.addEventListener("click", () => trocarMenu(botao.dataset.menu));
});

$("btnCarrinhoTopo").addEventListener("click", () => $("painelCotacao").classList.add("ativo"));
$("btnCarrinhoFlutuante").addEventListener("click", () => $("painelCotacao").classList.add("ativo"));
$("btnFecharCotacao").addEventListener("click", () => $("painelCotacao").classList.remove("ativo"));

$("btnCopiarCotacao").addEventListener("click", copiarCotacao);
$("btnGerarPdf").addEventListener("click", gerarPDF);
$("btnLimparCotacao").addEventListener("click", limparCotacao);

$("btnTopo").addEventListener("click", () => window.scrollTo({ top: 0, behavior: "smooth" }));

window.addEventListener("scroll", controlarBotaoTopo);

document.addEventListener("click", e => {
  if (!e.target.closest(".campo-busca")) $("sugestoesBusca").innerHTML = "";
  if (!e.target.closest(".campo-fornecedor")) $("sugestoesFornecedor").innerHTML = "";
});

controlarBotaoTopo();
renderizarCotacao();
renderizarHistoricoBusca();
