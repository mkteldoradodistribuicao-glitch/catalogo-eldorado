let produtos = [];
let produtosFiltrados = [];
let favoritos = [];
let scannerStream = null;
let scannerTrack = null;
let scannerTorchLigado = false;
let scannerHorizontalAtivo = false;
let scannerAtivo = false;
let scannerDetector = null;
let scannerAnimationFrame = null;
let scannerLeitorZXing = null;
let scannerControleZXing = null;
let ultimoCodigoScanner = "";
let ultimoCodigoScannerEm = 0;
let modoAtual = "eldorado";
let paginaAtual = 1;
let categoriaAtual = "todos";

const ITENS_POR_PAGINA_DESKTOP = 150;
const ITENS_POR_PAGINA_MOBILE = 80;
const CODIGOS_EXCLUIDOS_TERNURA = ["93217", "89817"];
const CODIGOS_EXCLUIDOS_NOVIDADES = ["999999", "116451"];
const CHAVE_FAVORITOS = "favoritosCatalogoEldorado";

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
    filtroNovidade: true,
    filtroFavoritos: false
  },
  favoritos: {
    titulo: "Favoritos",
    subtitulo: "Produtos marcados como favoritos para acesso rápido.",
    logo: "Logos/Eldorado.png",
    tema: "tema-eldorado",
    filtroTernura: false,
    filtroNovidade: false,
    filtroFavoritos: true
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

    const estrutura = detectarEstruturaDaPlanilha(dados);

    produtos = estrutura.linhas
      .map(linha => criarProdutoAPartirDaLinha(linha, estrutura.indices))
      .filter(produto => produto.codigo || produto.descricao || produto.ean);

    marcarNovidades();
    carregarFavoritos();
    aplicarFiltros();

  } catch (erro) {
    document.getElementById("contador").innerText = "";
    document.getElementById("catalogo").innerHTML =
      "<p>Não foi possível carregar a planilha de produtos.</p>";
    console.error(erro);
  }
}

function detectarEstruturaDaPlanilha(dados) {
  const indiceCabecalho = dados.findIndex(linha => {
    const cabecalhos = linha.map(celula => normalizar(celula));

    return cabecalhos.some(cabecalho =>
      cabecalho === "codigo" ||
      cabecalho.includes("codigo do produto") ||
      cabecalho.includes("cod produto")
    ) && cabecalhos.some(cabecalho => cabecalho.includes("descricao"));
  });

  if (indiceCabecalho < 0) {
    return {
      linhas: dados.slice(1),
      indices: null
    };
  }

  const cabecalho = dados[indiceCabecalho].map(celula => normalizar(celula));

  const localizar = (...nomes) => {
    for (const nome of nomes) {
      const nomeNormalizado = normalizar(nome);

      let indice = cabecalho.findIndex(valor => valor === nomeNormalizado);

      if (indice >= 0) return indice;

      indice = cabecalho.findIndex(valor =>
        valor && valor.includes(nomeNormalizado)
      );

      if (indice >= 0) return indice;
    }

    return -1;
  };

  return {
    linhas: dados.slice(indiceCabecalho + 1),
    indices: {
      codigo: localizar("codigo do produto", "codigo"),
      ean: localizar("ean"),
      descricao: localizar("descricao"),
      embalagem: localizar("embalagem"),
      qtdMaster: localizar("qtd master", "quantidade master"),
      estoque: localizar("estoque"),
      imagem: localizar("imagem"),
      codigoFornecedor: localizar("codigo fornecedor", "codigo do fornecedor"),
      fornecedor: localizar("fornecedor"),
      categoria: localizar("categoria")
    }
  };
}

function criarProdutoAPartirDaLinha(linha, indices = null) {
  if (indices) {
    const obter = indice => indice >= 0 ? linha[indice] : "";

    return {
      codigo: String(obter(indices.codigo) || "").trim(),
      ean: String(obter(indices.ean) || "").trim(),
      descricao: String(obter(indices.descricao) || "").trim(),
      embalagem: String(obter(indices.embalagem) || "").trim(),
      qtdMaster: String(obter(indices.qtdMaster) || "").trim(),
      estoque: numeroSeguro(obter(indices.estoque)),
      imagem: String(obter(indices.imagem) || "").trim(),
      codigoFornecedor: String(obter(indices.codigoFornecedor) || "").trim(),
      fornecedor: String(obter(indices.fornecedor) || "").trim(),
      categoria: String(obter(indices.categoria) || "Demais").trim() || "Demais",
      novidade: false,
      relevancia: 999
    };
  }

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
    categoria: String(linha[inicio + 9] || "Demais").trim() || "Demais",
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

  if (config.filtroFavoritos) {
    lista = lista.filter(produto => favoritos.includes(String(produto.codigo).trim()));
  }

  if (categoriaAtual !== "todos") {
    const categoriaNormalizada = normalizar(categoriaAtual);

    lista = lista.filter(produto =>
      normalizar(produto.categoria || "Demais") === categoriaNormalizada
    );
  }

  return lista;
}

function selecionarCategoria(categoria) {
  const categoriaSelecionada = String(categoria || "").trim();

  if (!categoriaSelecionada) return;

  if (normalizar(categoriaAtual) === normalizar(categoriaSelecionada)) {
    categoriaAtual = "todos";
  } else {
    categoriaAtual = categoriaSelecionada;
  }

  atualizarCategoriasAtivas();
  paginaAtual = 1;
  aplicarFiltros();
}

function atualizarCategoriasAtivas() {
  document.querySelectorAll(".categoria-btn").forEach(botao => {
    const ativa =
      categoriaAtual !== "todos" &&
      normalizar(botao.dataset.categoria) === normalizar(categoriaAtual);

    botao.classList.toggle("ativo", ativa);
    botao.setAttribute("aria-pressed", ativa ? "true" : "false");
  });
}

function limparCategoriaSelecionada() {
  categoriaAtual = "todos";
  atualizarCategoriasAtivas();
}

function configurarAnimacaoCategorias() {
  const nav = document.getElementById("categoriasNav");

  if (nav) {
    nav.addEventListener("pointermove", event => {
      const rect = nav.getBoundingClientRect();
      const x = ((event.clientX - rect.left) / rect.width) * 100;
      const y = ((event.clientY - rect.top) / rect.height) * 100;

      nav.style.setProperty("--categorias-mouse-x", `${x}%`);
      nav.style.setProperty("--categorias-mouse-y", `${y}%`);
    });
  }

  document.querySelectorAll(".categoria-btn").forEach(botao => {
    botao.addEventListener("pointermove", event => {
      const rect = botao.getBoundingClientRect();

      botao.style.setProperty(
        "--categoria-mouse-x",
        `${event.clientX - rect.left}px`
      );

      botao.style.setProperty(
        "--categoria-mouse-y",
        `${event.clientY - rect.top}px`
      );
    });

    botao.addEventListener("click", () => {
      selecionarCategoria(botao.dataset.categoria);
    });
  });
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
  if (codigo.includes(busca)) return 3;

  if (descricao === busca) return 4;
  if (descricao.startsWith(descricaoDigitada)) return 5;
  if (descricao.startsWith(primeiraPalavra) && todosNaDescricao) return 6;
  if (todosNaOrdem) return 7;
  if (todosNaDescricao) return 8;

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

function calcularRelevanciaCodigoFornecedor(produto, buscaCodigoFornecedor) {
  const termos = dividirTermos(buscaCodigoFornecedor);

  if (!termos.length) {
    produto.relevanciaCodigoFornecedor = 999;
    return 999;
  }

  const codigoFornecedor = normalizar(produto.codigoFornecedor);
  const buscaCompleta = termos.join("");

  if (!codigoFornecedor) return 999;
  if (codigoFornecedor === buscaCompleta) return 1;
  if (codigoFornecedor.startsWith(buscaCompleta)) return 2;
  if (termos.every(termo => codigoFornecedor.includes(termo))) return 3;

  return 999;
}

function ordenarProdutos(lista) {
  const tipoOrdenacao = document.getElementById("ordenacao").value;
  const buscaPrincipal = document.getElementById("buscaPrincipal").value.trim();
  const buscaCodigoFornecedor = document.getElementById("buscaCodigoFornecedor").value.trim();
  const existeBuscaPrincipal = dividirTermos(buscaPrincipal).length > 0;
  const existeBuscaCodigoFornecedor = dividirTermos(buscaCodigoFornecedor).length > 0;
  const listaOrdenada = [...lista];

  listaOrdenada.sort((a, b) => {
    if (existeBuscaCodigoFornecedor && a.relevanciaCodigoFornecedor !== b.relevanciaCodigoFornecedor) {
      return a.relevanciaCodigoFornecedor - b.relevanciaCodigoFornecedor;
    }

    if (existeBuscaPrincipal && a.relevancia !== b.relevancia) {
      return a.relevancia - b.relevancia;
    }

    if (modoAtual === "novidades") {
      return compararCodigo(b, a);
    }

    return aplicarOrdenacaoEscolhida(a, b, tipoOrdenacao);
  });

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

  let resultado = aplicarFiltroEstoque(produtosDoModo());

  resultado = resultado.filter(produto =>
    fornecedorCombina(produto.fornecedor, buscaFornecedor)
  );

  const termosCodigoFornecedor = dividirTermos(buscaCodigoFornecedor);
  const codigoFornecedorDigitado = normalizar(buscaCodigoFornecedor).replace(/\s+/g, "");

  if (termosCodigoFornecedor.length) {
    const exatos = resultado.filter(produto =>
      normalizar(produto.codigoFornecedor).replace(/\s+/g, "") === codigoFornecedorDigitado
    );

    if (exatos.length) {
      resultado = exatos;
      resultado.forEach(produto => {
        produto.relevanciaCodigoFornecedor = 1;
      });
    } else {
      const iniciando = resultado.filter(produto =>
        normalizar(produto.codigoFornecedor)
          .replace(/\s+/g, "")
          .startsWith(codigoFornecedorDigitado)
      );

      if (iniciando.length) {
        resultado = iniciando;
        resultado.forEach(produto => {
          produto.relevanciaCodigoFornecedor = 2;
        });
      } else {
        resultado = resultado.filter(produto => {
          produto.relevanciaCodigoFornecedor =
            calcularRelevanciaCodigoFornecedor(produto, buscaCodigoFornecedor);

          return produto.relevanciaCodigoFornecedor < 999;
        });
      }
    }
  } else {
    resultado.forEach(produto => {
      produto.relevanciaCodigoFornecedor = 999;
    });
  }

  resultado = resultado.filter(produto =>
    produtoCombinaBuscaPrincipal(produto, buscaPrincipal)
  );

  produtosFiltrados = ordenarProdutos(resultado);
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
      <button class="btn-favorito ${produtoEstaFavorito(produto) ? "ativo" : ""}" type="button" title="Favoritar produto" aria-label="Favoritar produto">
        ${produtoEstaFavorito(produto) ? "★" : "☆"}
      </button>

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
    `;

    card.querySelector(".btn-favorito").addEventListener("click", () => {
      alternarFavorito(produto);
    });

    catalogo.appendChild(card);
  });

  renderizarPaginacao(paginacaoSuperior, totalPaginas);
  renderizarPaginacao(paginacao, totalPaginas);
}

function carregarFavoritos() {
  try {
    favoritos = JSON.parse(localStorage.getItem(CHAVE_FAVORITOS)) || [];
  } catch {
    favoritos = [];
  }
}

function salvarFavoritos() {
  localStorage.setItem(CHAVE_FAVORITOS, JSON.stringify(favoritos));
}

function produtoEstaFavorito(produto) {
  return favoritos.includes(String(produto.codigo).trim());
}

function alternarFavorito(produto) {
  const codigo = String(produto.codigo).trim();

  if (favoritos.includes(codigo)) {
    favoritos = favoritos.filter(item => item !== codigo);
    mostrarToast("Produto removido dos favoritos.");
  } else {
    favoritos.push(codigo);
    mostrarToast("Produto adicionado aos favoritos.");
  }

  salvarFavoritos();

  if (modoAtual === "favoritos") {
    aplicarFiltros();
  } else {
    mostrarProdutos();
  }
}


function limparFiltros() {
  document.getElementById("buscaPrincipal").value = "";
  document.getElementById("buscaCodigoFornecedor").value = "";
  document.getElementById("buscaFornecedor").value = "";
  document.getElementById("ordenacao").value = modoAtual === "novidades" ? "codigo-decrescente" : "maior-estoque";
  document.getElementById("filtroEstoque").value = "com-estoque";
  document.getElementById("sugestoesFornecedor").classList.remove("ativo");
  limparCategoriaSelecionada();

  paginaAtual = 1;
  aplicarFiltros();
}

let timerTrocaCatalogo = null;

function trocarModo(novoModo) {
  fecharMenu();

  if (!CONFIG[novoModo]) return;

  clearTimeout(timerTrocaCatalogo);
  document.body.classList.add("trocando-catalogo");

  timerTrocaCatalogo = setTimeout(() => {
    modoAtual = novoModo;

    const config = CONFIG[modoAtual];

    document.body.className = `${config.tema} trocando-catalogo`;
    document.getElementById("logoCatalogo").src = config.logo;
    document.getElementById("tituloCatalogo").innerText = config.titulo;
    document.getElementById("subtituloCatalogo").innerText = config.subtitulo;

    limparFiltros();

    setTimeout(() => {
      document.body.classList.remove("trocando-catalogo");
    }, 220);
  }, 120);
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


async function aplicarMelhoriasCameraScanner(stream) {
  scannerTrack = stream.getVideoTracks()[0] || null;

  if (!scannerTrack) return;

  const capabilities = typeof scannerTrack.getCapabilities === "function"
    ? scannerTrack.getCapabilities()
    : {};

  const advanced = [];

  if (Array.isArray(capabilities.focusMode) && capabilities.focusMode.includes("continuous")) {
    advanced.push({ focusMode: "continuous" });
  }

  if (Array.isArray(capabilities.exposureMode) && capabilities.exposureMode.includes("continuous")) {
    advanced.push({ exposureMode: "continuous" });
  }

  if (Array.isArray(capabilities.whiteBalanceMode) && capabilities.whiteBalanceMode.includes("continuous")) {
    advanced.push({ whiteBalanceMode: "continuous" });
  }

  if (capabilities.zoom) {
    const zoomInicial = Math.min(
      Number(capabilities.zoom.max || 1),
      Math.max(Number(capabilities.zoom.min || 1), 1.35)
    );

    advanced.push({ zoom: zoomInicial });
  }

  if (advanced.length && typeof scannerTrack.applyConstraints === "function") {
    try {
      await scannerTrack.applyConstraints({ advanced });
    } catch (erro) {
      console.warn("Alguns controles avançados da câmera não foram aplicados.", erro);
    }
  }
}

function configurarControlesScanner() {
  const capabilities = scannerTrack?.getCapabilities?.() || {};
  const settings = scannerTrack?.getSettings?.() || {};
  const zoomArea = document.getElementById("scannerZoomArea");
  const zoomInput = document.getElementById("scannerZoom");
  const flashButton = document.getElementById("scannerFlash");

  if (capabilities.zoom) {
    const min = Number(capabilities.zoom.min || 1);
    const max = Number(capabilities.zoom.max || min);
    const step = Number(capabilities.zoom.step || 0.1);
    const atual = Number(settings.zoom || min);

    zoomInput.min = String(min);
    zoomInput.max = String(max);
    zoomInput.step = String(step);
    zoomInput.value = String(atual);

    zoomArea.classList.remove("oculto");
    atualizarTextoZoomScanner(atual);
  } else {
    zoomArea.classList.add("oculto");
  }

  flashButton.disabled = !capabilities.torch;
  atualizarBotaoFlashScanner();
}

async function alterarZoomScanner(valor) {
  if (!scannerTrack) return;

  const capabilities = scannerTrack.getCapabilities?.() || {};
  if (!capabilities.zoom) return;

  const min = Number(capabilities.zoom.min || 1);
  const max = Number(capabilities.zoom.max || min);
  const zoom = Math.min(max, Math.max(min, Number(valor)));

  try {
    await scannerTrack.applyConstraints({
      advanced: [{ zoom }]
    });

    document.getElementById("scannerZoom").value = String(zoom);
    atualizarTextoZoomScanner(zoom);
  } catch (erro) {
    console.warn("Não foi possível alterar o zoom.", erro);
  }
}

function atualizarTextoZoomScanner(valor) {
  document.getElementById("scannerZoomValor").innerText =
    `${Number(valor).toFixed(1)}×`;
}

async function alternarFlashScanner() {
  if (!scannerTrack || document.getElementById("scannerFlash").disabled) return;

  scannerTorchLigado = !scannerTorchLigado;

  try {
    await scannerTrack.applyConstraints({
      advanced: [{ torch: scannerTorchLigado }]
    });
  } catch (erro) {
    scannerTorchLigado = false;
    console.warn("O flash não está disponível neste dispositivo.", erro);
  }

  atualizarBotaoFlashScanner();
}

function atualizarBotaoFlashScanner() {
  document
    .getElementById("scannerFlash")
    .classList.toggle("ativo", scannerTorchLigado);
}

async function focarScanner(evento) {
  if (!scannerTrack) return;

  const capabilities = scannerTrack.getCapabilities?.() || {};
  const advanced = [];

  if (Array.isArray(capabilities.focusMode)) {
    if (capabilities.focusMode.includes("single-shot")) {
      advanced.push({ focusMode: "single-shot" });
    } else if (capabilities.focusMode.includes("continuous")) {
      advanced.push({ focusMode: "continuous" });
    }
  }

  if (evento && capabilities.pointsOfInterest) {
    const area = document.getElementById("scannerVideoArea");
    const rect = area.getBoundingClientRect();

    advanced.push({
      pointsOfInterest: [{
        x: Math.min(1, Math.max(0, (evento.clientX - rect.left) / rect.width)),
        y: Math.min(1, Math.max(0, (evento.clientY - rect.top) / rect.height))
      }]
    });
  }

  if (!advanced.length) {
    mostrarToast("Foco automático ativo.");
    return;
  }

  try {
    await scannerTrack.applyConstraints({ advanced });
    mostrarToast("Foco ajustado.");
  } catch (erro) {
    console.warn("Foco manual não suportado.", erro);
    mostrarToast("Foco automático ativo.");
  }
}

function alternarScannerHorizontal() {
  scannerHorizontalAtivo = !scannerHorizontalAtivo;

  document
    .getElementById("modalScanner")
    .classList.toggle("modo-horizontal", scannerHorizontalAtivo);

  document.querySelector("#scannerHorizontal small").innerText =
    scannerHorizontalAtivo ? "Padrão" : "Horizontal";
}

async function abrirScanner() {
  if (!navigator.mediaDevices?.getUserMedia) {
    alert("Este navegador não permite acessar a câmera.");
    return;
  }

  const modal = document.getElementById("modalScanner");
  const video = document.getElementById("videoScanner");
  const status = document.getElementById("scannerStatus");

  modal.classList.add("aberto");
  modal.setAttribute("aria-hidden", "false");
  document.body.classList.add("scanner-aberto");

  scannerAtivo = true;
  scannerTorchLigado = false;
  status.innerText = "Preparando câmera...";

  try {
    scannerStream = await navigator.mediaDevices.getUserMedia({
      audio: false,
      video: {
        facingMode: { ideal: "environment" },
        width: { ideal: 3840, min: 1280 },
        height: { ideal: 2160, min: 720 },
        frameRate: { ideal: 30, min: 20 },
        advanced: [
          { focusMode: "continuous" },
          { exposureMode: "continuous" },
          { whiteBalanceMode: "continuous" }
        ]
      }
    });

    await aplicarMelhoriasCameraScanner(scannerStream);

    video.srcObject = scannerStream;
    video.setAttribute("playsinline", "true");
    video.muted = true;
    await video.play();

    configurarControlesScanner();
    status.innerText = "Posicione o código dentro da moldura";
    iniciarLeituraScanner();
  } catch (erro) {
    console.error(erro);
    fecharScanner();
    alert("Não foi possível acessar a câmera. Verifique a permissão do navegador.");
  }
}

function iniciarLeituraScanner() {
  pararLeitoresScanner();

  if ("BarcodeDetector" in window) {
    iniciarDetectorNativoScanner();
    return;
  }

  iniciarLeitorZXingScanner();
}

async function iniciarDetectorNativoScanner() {
  try {
    const suportados = await BarcodeDetector.getSupportedFormats();
    const preferidos = [
      "ean_13",
      "ean_8",
      "upc_a",
      "upc_e",
      "code_128",
      "code_39",
      "itf"
    ].filter(formato => suportados.includes(formato));

    scannerDetector = new BarcodeDetector({
      formats: preferidos.length ? preferidos : suportados
    });

    const video = document.getElementById("videoScanner");

    const detectar = async () => {
      if (!scannerAtivo || !scannerDetector) return;

      try {
        if (video.readyState >= 2) {
          const codigos = await scannerDetector.detect(video);

          if (codigos.length) {
            processarCodigoScanner(codigos[0].rawValue);
            return;
          }
        }
      } catch (erro) {
        console.debug("Falha temporária na leitura do frame.", erro);
      }

      scannerAnimationFrame = requestAnimationFrame(detectar);
    };

    detectar();
  } catch (erro) {
    console.warn("Detector nativo indisponível. Tentando ZXing.", erro);
    iniciarLeitorZXingScanner();
  }
}

async function iniciarLeitorZXingScanner() {
  if (!window.ZXingBrowser?.BrowserMultiFormatReader) {
    document.getElementById("scannerStatus").innerText =
      "Leitor indisponível neste navegador";
    return;
  }

  try {
    scannerLeitorZXing = new ZXingBrowser.BrowserMultiFormatReader();

    scannerControleZXing = await scannerLeitorZXing.decodeFromVideoElement(
      document.getElementById("videoScanner"),
      (resultado, erro) => {
        if (!scannerAtivo) return;

        const codigo = resultado?.getText
          ? resultado.getText()
          : resultado?.text;

        if (codigo) {
          processarCodigoScanner(codigo);
        }

        if (erro && erro.name !== "NotFoundException") {
          console.debug("Leitura ZXing:", erro);
        }
      }
    );
  } catch (erro) {
    console.error(erro);
    document.getElementById("scannerStatus").innerText =
      "Não foi possível iniciar o leitor";
  }
}

function processarCodigoScanner(codigo) {
  const valor = String(codigo || "").trim();
  const agora = Date.now();

  if (!valor) return;

  if (
    valor === ultimoCodigoScanner &&
    agora - ultimoCodigoScannerEm < 1800
  ) {
    return;
  }

  ultimoCodigoScanner = valor;
  ultimoCodigoScannerEm = agora;

  if (navigator.vibrate) {
    navigator.vibrate(80);
  }

  document.getElementById("buscaPrincipal").value = valor;
  paginaAtual = 1;
  aplicarFiltros();
  mostrarToast(`Código lido: ${valor}`);
  fecharScanner();
}

function pararLeitoresScanner() {
  if (scannerAnimationFrame) {
    cancelAnimationFrame(scannerAnimationFrame);
    scannerAnimationFrame = null;
  }

  scannerDetector = null;

  if (scannerControleZXing?.stop) {
    try {
      scannerControleZXing.stop();
    } catch {}
  }

  scannerControleZXing = null;

  if (scannerLeitorZXing?.reset) {
    try {
      scannerLeitorZXing.reset();
    } catch {}
  }

  scannerLeitorZXing = null;
}

function fecharScanner() {
  scannerAtivo = false;
  pararLeitoresScanner();

  if (scannerTrack && scannerTorchLigado) {
    scannerTrack
      .applyConstraints({ advanced: [{ torch: false }] })
      .catch(() => {});
  }

  if (scannerStream) {
    scannerStream.getTracks().forEach(track => track.stop());
  }

  scannerStream = null;
  scannerTrack = null;
  scannerTorchLigado = false;

  const video = document.getElementById("videoScanner");
  video.pause();
  video.srcObject = null;

  const modal = document.getElementById("modalScanner");
  modal.classList.remove("aberto");
  modal.setAttribute("aria-hidden", "true");
  document.body.classList.remove("scanner-aberto");
}


function obterRotuloCategoriaAtual() {
  return categoriaAtual === "todos" ? "Todas as Categorias" : categoriaAtual;
}

function gerarPdfProdutosFiltrados() {
  if (!produtosFiltrados.length) {
    alert("Nenhum produto encontrado com os filtros atuais.");
    return;
  }

  if (!window.jspdf?.jsPDF) {
    alert("Não foi possível carregar o gerador de PDF.");
    return;
  }

  const { jsPDF } = window.jspdf;
  const doc = new jsPDF("p", "mm", "a4");

  const dataAtual = new Date().toLocaleDateString("pt-BR");
  const buscaPrincipal = document.getElementById("buscaPrincipal").value.trim();
  const codigoFornecedor = document.getElementById("buscaCodigoFornecedor").value.trim();
  const fornecedor = document.getElementById("buscaFornecedor").value.trim();
  const filtroEstoque = document.getElementById("filtroEstoque");
  const rotuloEstoque = filtroEstoque.options[filtroEstoque.selectedIndex]?.text || "";
  const categoria = obterRotuloCategoriaAtual();

  doc.setFont("helvetica", "bold");
  doc.setFontSize(18);
  doc.text("CATÁLOGO ELDORADO", 14, 18);

  doc.setFontSize(13);
  doc.text("Produtos filtrados", 14, 26);

  doc.setFont("helvetica", "normal");
  doc.setFontSize(9);

  let y = 34;

  doc.text(`Data: ${dataAtual}`, 14, y);
  y += 5;

  doc.text(`Categoria: ${categoria}`, 14, y);
  y += 5;

  if (buscaPrincipal) {
    doc.text(`Busca: ${buscaPrincipal}`, 14, y, { maxWidth: 180 });
    y += 5;
  }

  if (codigoFornecedor) {
    doc.text(`Código Fornecedor: ${codigoFornecedor}`, 14, y);
    y += 5;
  }

  if (fornecedor) {
    doc.text(`Fornecedor: ${fornecedor}`, 14, y, { maxWidth: 180 });
    y += 5;
  }

  doc.text(`Estoque: ${rotuloEstoque}`, 14, y, { maxWidth: 180 });
  y += 5;

  doc.text(`Total de produtos: ${produtosFiltrados.length}`, 14, y);
  y += 7;

  const linhas = produtosFiltrados.map(produto => [
    produto.codigo || "",
    produto.descricao || "",
    produto.ean || "",
    produto.embalagem || "",
    produto.qtdMaster || "",
    produto.fornecedor || ""
  ]);

  doc.autoTable({
    startY: y,
    head: [[
      "Código",
      "Descrição",
      "EAN",
      "Embalagem",
      "QTD Master",
      "Fornecedor"
    ]],
    body: linhas,
    styles: {
      fontSize: 7,
      cellPadding: 1.6,
      overflow: "linebreak",
      valign: "middle"
    },
    headStyles: {
      fillColor: [0, 150, 57],
      textColor: [255, 255, 255],
      fontStyle: "bold"
    },
    columnStyles: {
      0: { cellWidth: 18 },
      1: { cellWidth: 62 },
      2: { cellWidth: 31 },
      3: { cellWidth: 22 },
      4: { cellWidth: 18, halign: "center" },
      5: { cellWidth: 35 }
    },
    margin: {
      left: 10,
      right: 10,
      bottom: 14
    },
    didDrawPage: () => {
      const pageHeight = doc.internal.pageSize.height;
      doc.setFontSize(7);
      doc.setTextColor(90);
      doc.text(
        "Catálogo ilustrativo. Consulte disponibilidade, preço e condições comerciais com o RCA.",
        10,
        pageHeight - 7
      );
      doc.setTextColor(0);
    }
  });

  const nomeCategoria = normalizar(categoria)
    .replace(/[^a-z0-9]+/g, "_")
    .replace(/^_+|_+$/g, "")
    .toUpperCase();

  doc.save(`CATALOGO_ELDORADO_${nomeCategoria || "PRODUTOS"}.pdf`);
  mostrarToast("PDF gerado com os produtos filtrados.");
}


function voltarParaCatalogoPrincipal(){
  if(modoAtual!=="eldorado"){trocarModo("eldorado");}else{limparFiltros();}
  window.scrollTo({top:0,behavior:"smooth"});
}
function configurarAnimacaoCabecalho(){
  const topo=document.querySelector(".topo");
  if(!topo)return;
  topo.addEventListener("pointermove",e=>{
    const r=topo.getBoundingClientRect();
    topo.style.setProperty("--topo-x",`${((e.clientX-r.left)/r.width)*100}%`);
    topo.style.setProperty("--topo-y",`${((e.clientY-r.top)/r.height)*100}%`);
  });
  topo.addEventListener("pointerleave",()=>{
    topo.style.setProperty("--topo-x","24%");
    topo.style.setProperty("--topo-y","34%");
  });
}


function configurarRodapeCatalogo() {
  document.querySelectorAll("[data-rodape-modo]").forEach(botao => {
    botao.addEventListener("click", () => {
      const modo = botao.dataset.rodapeModo;
      trocarModo(modo);
      window.scrollTo({ top: 0, behavior: "smooth" });
    });
  });

  document.getElementById("rodapeVoltarTopo")?.addEventListener("click", () => {
    window.scrollTo({ top: 0, behavior: "smooth" });
  });
}

function atualizarInformacoesRodape() {
  const total = document.getElementById("rodapeTotalProdutos");
  if (!total) return;

  const quantidade = Array.isArray(produtos) ? produtos.length : 0;
  total.textContent = quantidade
    ? `${quantidade.toLocaleString("pt-BR")} produtos no catálogo`
    : "Catálogo digital";
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


document.getElementById("btnGerarPdfProdutos").addEventListener("click", gerarPdfProdutosFiltrados);
document.getElementById("btnPdfFlutuante").addEventListener("click", gerarPdfProdutosFiltrados);
document.getElementById("btnLogoCatalogo").addEventListener("click", voltarParaCatalogoPrincipal);
document.getElementById("btnLimparFiltros").addEventListener("click", limparFiltros);

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

document.getElementById("btnFavoritos").addEventListener("click", () => {
  trocarModo("favoritos");
});




document.getElementById("btnScanner").addEventListener("click", () => {
  abrirScanner();
  fecharMenu();
});
document.getElementById("fecharScanner").addEventListener("click", fecharScanner);
document.getElementById("scannerFlash").addEventListener("click", alternarFlashScanner);
document.getElementById("scannerFoco").addEventListener("click", () => focarScanner());
document.getElementById("scannerHorizontal").addEventListener("click", alternarScannerHorizontal);
document.getElementById("alternarScannerHorizontal").addEventListener("click", alternarScannerHorizontal);
document.getElementById("scannerVideoArea").addEventListener("click", focarScanner);

document.getElementById("scannerZoom").addEventListener("input", event => {
  alterarZoomScanner(event.target.value);
});

document.getElementById("scannerZoomMenos").addEventListener("click", () => {
  const input = document.getElementById("scannerZoom");
  alterarZoomScanner(Number(input.value) - Number(input.step || 0.1));
});

document.getElementById("scannerZoomMais").addEventListener("click", () => {
  const input = document.getElementById("scannerZoom");
  alterarZoomScanner(Number(input.value) + Number(input.step || 0.1));
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


document.addEventListener("visibilitychange", () => {
  if (document.hidden && scannerAtivo) {
    fecharScanner();
  }
});

window.addEventListener("beforeunload", () => {
  if (scannerAtivo) {
    fecharScanner();
  }
});

const btnTopo = document.getElementById("btnTopo");
const btnPdfFlutuante = document.getElementById("btnPdfFlutuante");

window.addEventListener("scroll", () => {
  if (window.scrollY > 300) {
    btnTopo.classList.add("visivel");
    btnPdfFlutuante.classList.add("visivel");
  } else {
    btnTopo.classList.remove("visivel");
    btnPdfFlutuante.classList.remove("visivel");
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
configurarAnimacaoCabecalho();
configurarRodapeCatalogo();
configurarAnimacaoCategorias();
atualizarCategoriasAtivas();
carregarProdutos();
