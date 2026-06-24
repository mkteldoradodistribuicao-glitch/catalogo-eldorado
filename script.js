let produtos = [];
let produtosFiltrados = [];
let cotacao = [];
let modoAtual = "eldorado";
let paginaAtual = 1;

const ITENS_POR_PAGINA_DESKTOP = 150;
const ITENS_POR_PAGINA_MOBILE = 80;
const CODIGOS_EXCLUIDOS_TERNURA = ["93217", "89817"];
const CODIGOS_EXCLUIDOS_NOVIDADES = ["999999", "116451"];
const AVISO_ESTOQUE = "⚠️ Verifique a disponibilidade de estoque antes de finalizar esta cotação.";

const CONFIG = {
  eldorado: {
    titulo: "Catálogo Eldorado",
    subtitulo: "Consulte produtos por código, descrição, EAN, fornecedor ou código do fornecedor.",
    logo: "Logos/Eldorado.png",
    tema: "tema-eldorado",
    filtroTernura: false,
    filtroNovidade: false
  },
  ternura: {
    titulo: "Produtos Ternura",
    subtitulo: "Catálogo exclusivo da linha Produtos Ternura.",
    logo: "Logos/Produtos-ternura.png",
    tema: "tema-ternura",
    filtroTernura: true,
    filtroNovidade: false
  },
  novidades: {
    titulo: "Novidades",
    subtitulo: "Conheça os lançamentos e as novidades que acabaram de chegar ao nosso mix.",
    logo: "Logos/Eldorado.png",
    tema: "tema-eldorado",
    filtroTernura: false,
    filtroNovidade: true
  }
};

async function carregarProdutos() {
  try {
    const resposta = await fetch("Produtos.xlsx");

    if (!resposta.ok) {
      throw new Error("Planilha Produtos.xlsx não encontrada na raiz do projeto.");
    }

    const arquivo = await resposta.arrayBuffer();
    const workbook = XLSX.read(arquivo, { type: "array" });
    const planilha = workbook.Sheets[workbook.SheetNames[0]];

    const dados = XLSX.utils.sheet_to_json(planilha, {
      header: 1,
      defval: ""
    });

    const linhas = detectarLinhasDeProdutos(dados);

    produtos = linhas
      .map(linha => criarProdutoAPartirDaLinha(linha))
      .filter(produto => produto.codigo || produto.descricao || produto.ean);

    marcarNovidades();
    carregarCotacaoSalva();
    aplicarFiltros();

  } catch (erro) {
    document.getElementById("contador").innerText = "";
    document.getElementById("catalogo").innerHTML =
      "<p>Não foi possível carregar a planilha de produtos.</p>";
    console.error(erro);
  }
}

function detectarLinhasDeProdutos(dados) {
  const indiceCabecalho = dados.findIndex(linha =>
    linha.some(celula => normalizar(celula).includes("codigo")) &&
    linha.some(celula => normalizar(celula).includes("descricao"))
  );

  if (indiceCabecalho >= 0) {
    return dados.slice(indiceCabecalho + 1);
  }

  return dados.slice(1);
}

function criarProdutoAPartirDaLinha(linha) {
  const primeiraColunaVazia = !String(linha[0] || "").trim();
  const inicio = primeiraColunaVazia ? 1 : 0;

  return {
    codigo: String(linha[inicio + 0] || "").trim(),
    ean: String(linha[inicio + 1] || "").trim(),
    descricao: String(linha[inicio + 2] || "").trim(),
    embalagem: String(linha[inicio + 3] || "").trim(),
    qtdMaster: String(linha[inicio + 4] || "").trim(),
    estoque: numeroSeguro(linha[inicio + 5]),
    imagem: String(linha[inicio + 6] || "").trim(),
    codigoFornecedor: String(linha[inicio + 7] || "").trim(),
    fornecedor: String(linha[inicio + 8] || "").trim(),
    novidade: false,
    relevancia: 999
  };
}

function numeroSeguro(valor) {
  if (typeof valor === "number") return valor;

  const texto = String(valor || "")
    .replace(/\./g, "")
    .replace(",", ".")
    .replace(/[^\d.-]/g, "");

  return Number(texto || 0);
}

function normalizar(texto) {
  return String(texto || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim();
}

function dividirTermos(texto) {
  return normalizar(texto)
    .replace(/^%+/, "")
    .split(/\s+|%+/)
    .map(t => t.trim())
    .filter(Boolean);
}

function codigoNumerico(produto) {
  const codigo = String(produto.codigo || "").replace(/\D/g, "");
  return Number(codigo || 0);
}

function qtdMasterNumerica(valor) {
  const numero = numeroSeguro(valor);
  return numero > 0 ? numero : 1;
}

function ajustarParaMultiplo(quantidade, qtdMaster) {
  const master = qtdMasterNumerica(qtdMaster);
  const qtd = Math.max(master, numeroSeguro(quantidade || master));
  return Math.ceil(qtd / master) * master;
}

function marcarNovidades() {
  const novidades = produtos
    .filter(produto =>
      codigoNumerico(produto) > 0 &&
      !CODIGOS_EXCLUIDOS_NOVIDADES.includes(String(produto.codigo).trim())
    )
    .sort((a, b) => codigoNumerico(b) - codigoNumerico(a))
    .slice(0, 150);

  const codigosNovidades = new Set(
    novidades.map(produto => String(produto.codigo).trim())
  );

  produtos.forEach(produto => {
    produto.novidade = codigosNovidades.has(String(produto.codigo).trim());
  });
}

function produtosDoModo() {
  const config = CONFIG[modoAtual];
  let lista = produtos;

  if (config.filtroTernura) {
    lista = lista.filter(produto =>
      normalizar(produto.descricao).includes("ternura") &&
      !CODIGOS_EXCLUIDOS_TERNURA.includes(String(produto.codigo).trim())
    );
  }

  if (config.filtroNovidade) {
    lista = lista.filter(produto => produto.novidade);
  }

  return lista;
}

function aplicarFiltroEstoque(lista) {
  const tipo = document.getElementById("filtroEstoque").value;

  if (tipo === "com-estoque") {
    return lista.filter(produto => numeroSeguro(produto.estoque) > 0);
  }

  if (tipo === "sem-estoque") {
    return lista.filter(produto => numeroSeguro(produto.estoque) <= 0);
  }

  return lista;
}

function compararCodigo(a, b) {
  const numeroA = codigoNumerico(a);
  const numeroB = codigoNumerico(b);

  if (numeroA !== numeroB) {
    return numeroA - numeroB;
  }

  return normalizar(a.codigo).localeCompare(normalizar(b.codigo), "pt-BR", { numeric: true });
}

function calcularRelevancia(produto, buscaPrincipal) {
  const busca = normalizar(buscaPrincipal);
  const termos = dividirTermos(buscaPrincipal);

  if (!termos.length) return 999;

  const codigo = normalizar(produto.codigo);
  const descricao = normalizar(produto.descricao);
  const ean = normalizar(produto.ean);
  const descricaoDigitada = termos.join(" ");
  const primeiraPalavra = termos[0];

  const todosNaDescricao = termos.every(termo => descricao.includes(termo));
  const todosNaOrdem = descricao.includes(descricaoDigitada);

  if (codigo === busca) return 1;
  if (codigo.startsWith(busca)) return 2;

  if (descricao === busca) return 3;
  if (descricao.startsWith(descricaoDigitada)) return 4;
  if (descricao.startsWith(primeiraPalavra) && todosNaDescricao) return 5;
  if (todosNaOrdem) return 6;
  if (todosNaDescricao) return 7;

  if (codigo.includes(busca)) return 8;

  if (ean.startsWith(busca)) return 9;
  if (ean.includes(busca)) return 10;

  return 999;
}

function produtoCombinaBuscaPrincipal(produto, buscaPrincipal) {
  const termos = dividirTermos(buscaPrincipal);

  if (!termos.length) {
    produto.relevancia = 999;
    return true;
  }

  produto.relevancia = calcularRelevancia(produto, buscaPrincipal);
  return produto.relevancia < 999;
}

function ordenarProdutos(lista) {
  const tipoOrdenacao = document.getElementById("ordenacao").value;
  const buscaPrincipal = document.getElementById("buscaPrincipal").value.trim();
  const existeBuscaPrincipal = dividirTermos(buscaPrincipal).length > 0;
  const listaOrdenada = [...lista];

  if (modoAtual === "novidades") {
    listaOrdenada.sort((a, b) => {
      if (existeBuscaPrincipal && a.relevancia !== b.relevancia) {
        return a.relevancia - b.relevancia;
      }

      return compararCodigo(b, a);
    });

    return listaOrdenada;
  }

  if (existeBuscaPrincipal) {
    listaOrdenada.sort((a, b) => {
      if (a.relevancia !== b.relevancia) return a.relevancia - b.relevancia;
      return aplicarOrdenacaoEscolhida(a, b, tipoOrdenacao);
    });

    return listaOrdenada;
  }

  listaOrdenada.sort((a, b) => aplicarOrdenacaoEscolhida(a, b, tipoOrdenacao));
  return listaOrdenada;
}

function aplicarOrdenacaoEscolhida(a, b, tipoOrdenacao) {
  switch (tipoOrdenacao) {
    case "codigo-crescente":
      return compararCodigo(a, b);

    case "codigo-decrescente":
      return compararCodigo(b, a);

    case "descricao-az":
      return normalizar(a.descricao).localeCompare(normalizar(b.descricao), "pt-BR");

    case "descricao-za":
      return normalizar(b.descricao).localeCompare(normalizar(a.descricao), "pt-BR");

    case "maior-estoque":
    default:
      return numeroSeguro(b.estoque) - numeroSeguro(a.estoque);
  }
}

function calcularDistancia(a, b) {
  a = normalizar(a);
  b = normalizar(b);

  const matriz = Array.from({ length: b.length + 1 }, (_, i) => [i]);

  for (let j = 0; j <= a.length; j++) {
    matriz[0][j] = j;
  }

  for (let i = 1; i <= b.length; i++) {
    for (let j = 1; j <= a.length; j++) {
      matriz[i][j] = b.charAt(i - 1) === a.charAt(j - 1)
        ? matriz[i - 1][j - 1]
        : Math.min(
            matriz[i - 1][j - 1] + 1,
            matriz[i][j - 1] + 1,
            matriz[i - 1][j] + 1
          );
    }
  }

  return matriz[b.length][a.length];
}

function pontuarFornecedor(nomeFornecedor, termoBusca, estoqueTotal) {
  const fornecedor = normalizar(nomeFornecedor);
  const busca = normalizar(termoBusca);

  if (!busca) return 0;

  let pontos = 0;

  if (fornecedor === busca) pontos += 1000;
  if (fornecedor.startsWith(busca)) pontos += 700;
  if (fornecedor.includes(busca)) pontos += 500;

  const termos = dividirTermos(termoBusca);

  termos.forEach(termo => {
    if (fornecedor.includes(termo)) pontos += 250;
    if (fornecedor.startsWith(termo)) pontos += 300;

    const palavras = fornecedor.split(/\s+/).filter(Boolean);

    palavras.forEach(palavra => {
      if (palavra.startsWith(termo)) pontos += 220;

      const distancia = calcularDistancia(palavra, termo);
      pontos += Math.max(0, 140 - distancia * 25);
    });
  });

  pontos += Math.min(numeroSeguro(estoqueTotal), 10000) * 0.01;

  return pontos;
}

function obterSugestoesFornecedor() {
  const termo = document.getElementById("buscaFornecedor").value;
  const termoNormalizado = normalizar(termo);

  if (!termoNormalizado) {
    return [];
  }

  const base = aplicarFiltroEstoque(produtosDoModo());
  const mapaFornecedores = new Map();

  base.forEach(produto => {
    if (!produto.fornecedor) return;

    const chave = normalizar(produto.fornecedor);

    if (!mapaFornecedores.has(chave)) {
      mapaFornecedores.set(chave, {
        nome: produto.fornecedor,
        estoqueTotal: 0,
        quantidadeProdutos: 0
      });
    }

    const fornecedor = mapaFornecedores.get(chave);
    fornecedor.estoqueTotal += numeroSeguro(produto.estoque);
    fornecedor.quantidadeProdutos += 1;
  });

  return [...mapaFornecedores.values()]
    .map(fornecedor => ({
      ...fornecedor,
      pontos: pontuarFornecedor(fornecedor.nome, termo, fornecedor.estoqueTotal)
    }))
    .filter(fornecedor => fornecedor.pontos > 0)
    .sort((a, b) => {
      if (b.pontos !== a.pontos) return b.pontos - a.pontos;
      if (b.estoqueTotal !== a.estoqueTotal) return b.estoqueTotal - a.estoqueTotal;
      return normalizar(a.nome).localeCompare(normalizar(b.nome), "pt-BR");
    })
    .slice(0, 5);
}

function mostrarSugestoesFornecedor() {
  const container = document.getElementById("sugestoesFornecedor");
  const sugestoes = obterSugestoesFornecedor();

  container.innerHTML = "";

  if (!sugestoes.length) {
    container.classList.remove("ativo");
    return;
  }

  sugestoes.forEach(sugestao => {
    const item = document.createElement("button");
    item.type = "button";
    item.className = "sugestao-item";
    item.innerHTML = `
      <strong>${sugestao.nome}</strong>
      <span>${sugestao.quantidadeProdutos} produtos</span>
    `;

    item.addEventListener("click", () => {
      document.getElementById("buscaFornecedor").value = sugestao.nome;
      container.classList.remove("ativo");
      paginaAtual = 1;
      aplicarFiltros();
    });

    container.appendChild(item);
  });

  container.classList.add("ativo");
}

function fornecedorCombina(produtoFornecedor, buscaFornecedor) {
  const termos = dividirTermos(buscaFornecedor);
  const fornecedor = normalizar(produtoFornecedor);

  if (!termos.length) return true;

  return termos.every(termo => {
    if (fornecedor.includes(termo)) return true;

    const palavrasFornecedor = fornecedor.split(/\s+/).filter(Boolean);

    return palavrasFornecedor.some(palavra => {
      const distancia = calcularDistancia(palavra, termo);
      return palavra.startsWith(termo) || distancia <= 2;
    });
  });
}

function gerarMensagemSemResultado() {
  const buscaPrincipal = document.getElementById("buscaPrincipal").value.trim();
  const buscaCodigoFornecedor = document.getElementById("buscaCodigoFornecedor").value.trim();
  const buscaFornecedor = document.getElementById("buscaFornecedor").value.trim();

  const partes = [];

  if (buscaPrincipal) partes.push(`"${buscaPrincipal}"`);
  if (buscaCodigoFornecedor) partes.push(`código de fornecedor "${buscaCodigoFornecedor}"`);
  if (buscaFornecedor) partes.push(`fornecedor "${buscaFornecedor}"`);

  if (!partes.length) return "Nenhum produto encontrado.";
  if (partes.length === 1) return `Nenhum produto encontrado para ${partes[0]}.`;
  if (partes.length === 2) return `Nenhum produto encontrado para ${partes[0]} dentro de ${partes[1]}.`;

  return `Nenhum produto encontrado para ${partes[0]} no ${partes[1]} e ${partes[2]}.`;
}

function aplicarFiltros() {
  const buscaPrincipal = document.getElementById("buscaPrincipal").value;
  const buscaCodigoFornecedor = document.getElementById("buscaCodigoFornecedor").value;
  const buscaFornecedor = document.getElementById("buscaFornecedor").value;

  const termosCodigoFornecedor = dividirTermos(buscaCodigoFornecedor);

  let resultado = produtosDoModo();

  resultado = aplicarFiltroEstoque(resultado);

  resultado = resultado.filter(produto => {
    const textoCodigoFornecedor = normalizar(produto.codigoFornecedor);

    const encontrouCodigoFornecedor = termosCodigoFornecedor.every(termo =>
      textoCodigoFornecedor.includes(termo)
    );

    const encontrouFornecedor = fornecedorCombina(produto.fornecedor, buscaFornecedor);

    return encontrouCodigoFornecedor && encontrouFornecedor;
  });

  resultado = resultado.filter(produto =>
    produtoCombinaBuscaPrincipal(produto, buscaPrincipal)
  );

  resultado = ordenarProdutos(resultado);

  produtosFiltrados = resultado;
  paginaAtual = 1;

  mostrarProdutos();
}

function obterItensPorPagina() {
  return window.matchMedia("(max-width: 700px)").matches
    ? ITENS_POR_PAGINA_MOBILE
    : ITENS_POR_PAGINA_DESKTOP;
}

function renderizarPaginacao(container, totalPaginas) {
  container.innerHTML = "";

  if (totalPaginas <= 1) return;

  const btnAnterior = document.createElement("button");
  btnAnterior.innerText = "Anterior";
  btnAnterior.disabled = paginaAtual === 1;
  btnAnterior.addEventListener("click", () => {
    paginaAtual--;
    mostrarProdutos();
    window.scrollTo({ top: 0, behavior: "smooth" });
  });

  const indicador = document.createElement("span");
  indicador.innerText = `Página ${paginaAtual} de ${totalPaginas}`;

  const btnProxima = document.createElement("button");
  btnProxima.innerText = "Próxima";
  btnProxima.disabled = paginaAtual === totalPaginas;
  btnProxima.addEventListener("click", () => {
    paginaAtual++;
    mostrarProdutos();
    window.scrollTo({ top: 0, behavior: "smooth" });
  });

  container.appendChild(btnAnterior);
  container.appendChild(indicador);
  container.appendChild(btnProxima);
}

function mostrarProdutos() {
  const catalogo = document.getElementById("catalogo");
  const contador = document.getElementById("contador");
  const paginacao = document.getElementById("paginacao");
  const paginacaoSuperior = document.getElementById("paginacaoSuperior");

  catalogo.innerHTML = "";
  paginacao.innerHTML = "";
  paginacaoSuperior.innerHTML = "";

  const totalProdutos = produtosFiltrados.length;
  const itensPorPagina = obterItensPorPagina();
  const totalPaginas = Math.max(1, Math.ceil(totalProdutos / itensPorPagina));

  if (paginaAtual > totalPaginas) {
    paginaAtual = totalPaginas;
  }

  if (totalProdutos === 0) {
    contador.innerText = "0 produtos encontrados";
    catalogo.innerHTML = `<p>${gerarMensagemSemResultado()}</p>`;
    return;
  }

  const inicio = (paginaAtual - 1) * itensPorPagina;
  const fim = inicio + itensPorPagina;
  const produtosDaPagina = produtosFiltrados.slice(inicio, fim);

  contador.innerText = `${totalProdutos} produtos encontrados • Página ${paginaAtual} de ${totalPaginas}`;

  produtosDaPagina.forEach(produto => {
    const card = document.createElement("div");
    card.className = "card";

    const caminhoImagem = produto.imagem ? `Imagens/${produto.imagem}` : "";

    card.innerHTML = `
      <div class="card-imagem">
        ${produto.novidade ? `<span class="selo-novidade">NOVIDADE</span>` : ""}
        ${
          produto.imagem
            ? `<img src="${caminhoImagem}" alt="${produto.descricao}" onerror="this.outerHTML='<div class=&quot;sem-imagem&quot;>Imagem não encontrada</div>'">`
            : `<div class="sem-imagem">Sem imagem</div>`
        }
      </div>

      <div class="codigo">Código: ${produto.codigo || "Não informado"}</div>
      <div class="descricao">${produto.descricao || "Descrição não informada"}</div>
      <div class="info">EAN: ${produto.ean || "Não informado"}</div>
      <div class="info">Embalagem: ${produto.embalagem || "Não informada"}</div>
      <div class="info">QTD Master: ${produto.qtdMaster || "Não informada"}</div>
      <div class="info">Código Fornecedor: ${produto.codigoFornecedor || "Não informado"}</div>
      <div class="fornecedor">${produto.fornecedor || "Fornecedor não informado"}</div>

      <button class="btn-adicionar-cotacao" type="button" title="Adicionar à cotação">🛒 Adicionar à cotação</button>
    `;

    card.querySelector(".btn-adicionar-cotacao").addEventListener("click", () => {
      adicionarNaCotacao(produto);
      aplicarEfeitoCardAdicionado(card);
    });

    catalogo.appendChild(card);
  });

  renderizarPaginacao(paginacaoSuperior, totalPaginas);
  renderizarPaginacao(paginacao, totalPaginas);
}

function aplicarEfeitoCardAdicionado(card) {
  card.classList.remove("produto-adicionado");
  void card.offsetWidth;
  card.classList.add("produto-adicionado");

  setTimeout(() => {
    card.classList.remove("produto-adicionado");
  }, 900);
}

function adicionarNaCotacao(produto) {
  const codigo = String(produto.codigo).trim();
  const itemExistente = cotacao.find(item => String(item.codigo).trim() === codigo);
  const master = qtdMasterNumerica(produto.qtdMaster);

  if (itemExistente) {
    itemExistente.quantidade = ajustarParaMultiplo(numeroSeguro(itemExistente.quantidade) + master, master);
  } else {
    cotacao.push({
      codigo: produto.codigo,
      descricao: produto.descricao,
      ean: produto.ean,
      quantidade: master,
      qtdMaster: master,
      estoque: numeroSeguro(produto.estoque)
    });
  }

  salvarCotacao();
  renderizarCotacao();
  mostrarToast("Produto adicionado à cotação.");
}

function removerDaCotacao(codigo) {
  cotacao = cotacao.filter(item => String(item.codigo).trim() !== String(codigo).trim());
  salvarCotacao();
  renderizarCotacao();
}

function atualizarQuantidade(codigo, quantidade) {
  const item = cotacao.find(item => String(item.codigo).trim() === String(codigo).trim());

  if (!item) return;

  item.quantidade = ajustarParaMultiplo(quantidade, item.qtdMaster);
  salvarCotacao();
  renderizarCotacao();
}

function salvarCotacao() {
  localStorage.setItem("cotacaoEldorado", JSON.stringify(cotacao));
}

function carregarCotacaoSalva() {
  try {
    cotacao = JSON.parse(localStorage.getItem("cotacaoEldorado")) || [];
  } catch {
    cotacao = [];
  }

  renderizarCotacao();
}

function deveAvisarEstoque(item) {
  const estoque = numeroSeguro(item.estoque);
  const master = qtdMasterNumerica(item.qtdMaster);
  const quantidade = numeroSeguro(item.quantidade);

  if (estoque <= 0 || master <= 0 || quantidade <= 0) return false;

  const limiteAviso = estoque - (master * 3);
  return quantidade >= limiteAviso;
}

function renderizarCotacao() {
  const contador = cotacao.length;

  document.getElementById("contadorCotacao").innerText = contador;
  document.getElementById("contadorCotacaoFlutuante").innerText = contador;

  const lista = document.getElementById("listaCotacao");
  lista.innerHTML = "";

  if (!cotacao.length) {
    lista.innerHTML = "<p>Nenhum produto adicionado à cotação.</p>";
    return;
  }

  cotacao.forEach(item => {
    const master = qtdMasterNumerica(item.qtdMaster);

    const div = document.createElement("div");
    div.className = "item-cotacao";

    div.innerHTML = `
      <strong>Código do Produto: ${item.codigo}</strong>
      <p><b>Descrição:</b> ${item.descricao || "Não informada"}</p>
      <p><b>EAN:</b> ${item.ean || "Não informado"}</p>

      <div class="item-cotacao-acoes">
        <label>
          Quantidade desejada
          <input type="number" min="${master}" step="${master}" value="${item.quantidade}" />
        </label>

        <button type="button">Remover</button>
      </div>

      <p><b>Múltiplo QTD Master:</b> ${master}</p>
      ${deveAvisarEstoque(item) ? `<div class="aviso-estoque">${AVISO_ESTOQUE}</div>` : ""}
    `;

    div.querySelector("input").addEventListener("change", event => {
      atualizarQuantidade(item.codigo, event.target.value);
    });

    div.querySelector("button").addEventListener("click", () => {
      removerDaCotacao(item.codigo);
    });

    lista.appendChild(div);
  });
}

function formatarCNPJ(valor) {
  return valor
    .replace(/\D/g, "")
    .replace(/^(\d{2})(\d)/, "$1.$2")
    .replace(/^(\d{2})\.(\d{3})(\d)/, "$1.$2.$3")
    .replace(/\.(\d{3})(\d)/, ".$1/$2")
    .replace(/(\d{4})(\d)/, "$1-$2")
    .slice(0, 18);
}

function gerarTextoCotacao() {
  const nomeInformado = document.getElementById("nomePdf").value.trim();
  const cnpjCliente = document.getElementById("cnpjCliente").value.trim();

  let texto = "SOLICITAÇÃO DE COTAÇÃO\n";
  texto += "Catálogo ilustrativo. Não caracteriza Pedido de Compra. Confirmar disponibilidade, preço e condições comerciais com o RCA.\n\n";

  if (nomeInformado) texto += `Razão Social/Comprador: ${nomeInformado}\n`;
  if (cnpjCliente) texto += `CNPJ: ${cnpjCliente}\n`;

  texto += "\n";

  cotacao.forEach((item, index) => {
    texto += `${index + 1}) Código do Produto: ${item.codigo}\n`;
    texto += `Descrição: ${item.descricao || "Não informada"}\n`;
    texto += `EAN: ${item.ean || "Não informado"}\n`;
    texto += `Quantidade desejada: ${item.quantidade}\n\n`;
  });

  return texto.trim();
}

function perguntarManterOuLimpar() {
  const limpar = confirm("Cotação concluída. Deseja limpar o carrinho?\n\nOK = Limpar carrinho\nCancelar = Manter carrinho");

  if (limpar) {
    cotacao = [];
    salvarCotacao();
    renderizarCotacao();
    mostrarToast("Carrinho limpo.");
  } else {
    mostrarToast("Carrinho mantido.");
  }
}

function copiarCotacao() {
  if (!cotacao.length) {
    alert("Nenhum produto foi adicionado à cotação.");
    return;
  }

  navigator.clipboard.writeText(gerarTextoCotacao())
    .then(() => {
      mostrarToast("Cotação copiada com sucesso.");
      perguntarManterOuLimpar();
    })
    .catch(() => alert("Não foi possível copiar a cotação."));
}

function limparCotacao() {
  if (!cotacao.length) return;

  const confirmar = confirm("Deseja limpar toda a lista de cotação?");

  if (!confirmar) return;

  cotacao = [];
  salvarCotacao();
  renderizarCotacao();
}

function limparNomeArquivo(texto) {
  return normalizar(texto)
    .replace(/[^a-z0-9\s-]/g, "")
    .replace(/\s+/g, "_")
    .toUpperCase();
}

async function gerarPdfCotacao() {
  if (!cotacao.length) {
    alert("Nenhum produto foi adicionado à cotação.");
    return;
  }

  const nomeInformado = document.getElementById("nomePdf").value.trim();
  const cnpjCliente = document.getElementById("cnpjCliente").value.trim();

  if (!nomeInformado) {
    alert("Informe a Razão Social da loja ou o Comprador responsável para nomear o PDF.");
    return;
  }

  if (!cnpjCliente) {
    alert("Informe o CNPJ do Cliente.");
    return;
  }

  const { jsPDF } = window.jspdf;
  const doc = new jsPDF("p", "mm", "a4");

  const dataAtual = new Date().toLocaleDateString("pt-BR");
  const totalItens = cotacao.reduce((soma, item) => soma + numeroSeguro(item.quantidade), 0);

  try {
    doc.addImage("Logos/Eldorado.png", "PNG", 14, 10, 42, 22);
  } catch {}

  doc.setFont("helvetica", "bold");
  doc.setFontSize(16);
  doc.text("SOLICITAÇÃO DE COTAÇÃO", 14, 42);

  doc.setFont("helvetica", "normal");
  doc.setFontSize(10);
  doc.text(`Razão Social/Comprador: ${nomeInformado}`, 14, 51);
  doc.text(`CNPJ: ${cnpjCliente}`, 14, 58);
  doc.text(`Data: ${dataAtual}`, 14, 65);
  doc.text(`Produtos distintos: ${cotacao.length}`, 14, 72);
  doc.text(`Total de itens solicitados: ${totalItens}`, 14, 79);

  const linhas = cotacao.map(item => [
    item.codigo || "",
    item.descricao || "",
    item.ean || "",
    String(item.quantidade || 1)
  ]);

  doc.autoTable({
    startY: 88,
    head: [["Código", "Descrição", "EAN", "Quantidade"]],
    body: linhas,
    styles: {
      fontSize: 8,
      cellPadding: 2
    },
    headStyles: {
      fillColor: [0, 150, 57],
      textColor: [255, 255, 255]
    },
    columnStyles: {
      0: { cellWidth: 25 },
      1: { cellWidth: 90 },
      2: { cellWidth: 45 },
      3: { cellWidth: 25, halign: "center" }
    }
  });

  const yFinal = doc.lastAutoTable.finalY + 10;

  doc.setFontSize(8);
  doc.text(
    "Catálogo ilustrativo. Não caracteriza Pedido de Compra. A disponibilidade, preços e condições comerciais devem ser confirmados com o RCA antes da finalização.",
    14,
    yFinal,
    { maxWidth: 180 }
  );

  const nomeArquivo = `COTACAO_${limparNomeArquivo(nomeInformado)}.pdf`;
  doc.save(nomeArquivo);

  mostrarToast("PDF gerado com sucesso.");
  perguntarManterOuLimpar();
}

function abrirCotacao() {
  document.getElementById("painelCotacao").classList.add("aberto");
  document.getElementById("overlayCotacao").classList.add("aberto");
  document.body.classList.add("cotacao-aberta");
}

function fecharCotacao() {
  document.getElementById("painelCotacao").classList.remove("aberto");
  document.getElementById("overlayCotacao").classList.remove("aberto");
  document.body.classList.remove("cotacao-aberta");
}

function limparFiltros() {
  document.getElementById("buscaPrincipal").value = "";
  document.getElementById("buscaCodigoFornecedor").value = "";
  document.getElementById("buscaFornecedor").value = "";
  document.getElementById("ordenacao").value = modoAtual === "novidades" ? "codigo-decrescente" : "maior-estoque";
  document.getElementById("filtroEstoque").value = "com-estoque";
  document.getElementById("sugestoesFornecedor").classList.remove("ativo");

  paginaAtual = 1;
  aplicarFiltros();
}

function trocarModo(novoModo) {
  modoAtual = novoModo;

  const config = CONFIG[modoAtual];

  document.body.className = config.tema;
  document.getElementById("logoCatalogo").src = config.logo;
  document.getElementById("tituloCatalogo").innerText = config.titulo;
  document.getElementById("subtituloCatalogo").innerText = config.subtitulo;

  limparFiltros();
  fecharMenu();
}

function abrirFecharMenu() {
  document.getElementById("menuLateral").classList.toggle("aberto");
}

function fecharMenu() {
  document.getElementById("menuLateral").classList.remove("aberto");
}

let toastTimer = null;

function mostrarToast(mensagem) {
  const toast = document.getElementById("toast");
  toast.innerText = mensagem;
  toast.classList.add("visivel");

  clearTimeout(toastTimer);

  toastTimer = setTimeout(() => {
    toast.classList.remove("visivel");
  }, 2600);
}

document.getElementById("buscaPrincipal").addEventListener("input", aplicarFiltros);
document.getElementById("buscaCodigoFornecedor").addEventListener("input", aplicarFiltros);

document.getElementById("buscaFornecedor").addEventListener("input", () => {
  mostrarSugestoesFornecedor();
  aplicarFiltros();
});

document.getElementById("buscaFornecedor").addEventListener("focus", mostrarSugestoesFornecedor);

document.getElementById("ordenacao").addEventListener("change", aplicarFiltros);

document.getElementById("filtroEstoque").addEventListener("change", () => {
  mostrarSugestoesFornecedor();
  aplicarFiltros();
});

document.getElementById("cnpjCliente").addEventListener("input", event => {
  event.target.value = formatarCNPJ(event.target.value);
});

document.getElementById("btnLimparFiltros").addEventListener("click", limparFiltros);
document.getElementById("btnCotacao").addEventListener("click", abrirCotacao);
document.getElementById("btnCarrinhoFlutuante").addEventListener("click", abrirCotacao);
document.getElementById("fecharCotacao").addEventListener("click", fecharCotacao);
document.getElementById("overlayCotacao").addEventListener("click", fecharCotacao);
document.getElementById("copiarCotacao").addEventListener("click", copiarCotacao);
document.getElementById("limparCotacao").addEventListener("click", limparCotacao);
document.getElementById("gerarPdf").addEventListener("click", gerarPdfCotacao);

document.getElementById("btnMenu").addEventListener("click", function(event) {
  event.stopPropagation();
  abrirFecharMenu();
});

document.getElementById("btnCatalogoEldorado").addEventListener("click", () => {
  trocarModo("eldorado");
});

document.getElementById("btnProdutosTernura").addEventListener("click", () => {
  trocarModo("ternura");
});

document.getElementById("btnNovidades").addEventListener("click", () => {
  trocarModo("novidades");
});

document.addEventListener("click", function(event) {
  const menu = document.getElementById("menuLateral");
  const botao = document.getElementById("btnMenu");
  const campoFornecedor = document.querySelector(".campo-fornecedor");

  if (!menu.contains(event.target) && !botao.contains(event.target)) {
    fecharMenu();
  }

  if (!campoFornecedor.contains(event.target)) {
    document.getElementById("sugestoesFornecedor").classList.remove("ativo");
  }
});

const btnTopo = document.getElementById("btnTopo");
const btnCarrinhoFlutuante = document.getElementById("btnCarrinhoFlutuante");

window.addEventListener("scroll", () => {
  if (window.scrollY > 300) {
    btnTopo.classList.add("visivel");
    btnCarrinhoFlutuante.classList.add("visivel");
  } else {
    btnTopo.classList.remove("visivel");
    btnCarrinhoFlutuante.classList.remove("visivel");
  }
});

btnTopo.addEventListener("click", () => {
  window.scrollTo({
    top: 0,
    behavior: "smooth"
  });
});

window.addEventListener("resize", () => {
  if (!produtosFiltrados.length) return;

  paginaAtual = Math.min(
    paginaAtual,
    Math.max(1, Math.ceil(produtosFiltrados.length / obterItensPorPagina()))
  );

  mostrarProdutos();
});

document.getElementById("ordenacao").value = "maior-estoque";
carregarProdutos();
