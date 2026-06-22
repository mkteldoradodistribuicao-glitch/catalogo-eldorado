let produtos = [];
let produtosFiltrados = [];
let modoAtual = "eldorado";
let paginaAtual = 1;

const ITENS_POR_PAGINA = 150;
const CODIGOS_EXCLUIDOS_TERNURA = ["93217", "89817"];
const CODIGOS_EXCLUIDOS_NOVIDADES = ["999999", "116451"];

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
    logo: "Logos/Produtos-Ternura.png",
    tema: "tema-ternura",
    filtroTernura: true,
    filtroNovidade: false
  },
  novidades: {
    titulo: "Novidades",
    subtitulo: "Os 150 maiores códigos de produtos, destacados com selo de novidade.",
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
      throw new Error("Planilha não encontrada.");
    }

    const arquivo = await resposta.arrayBuffer();
    const workbook = XLSX.read(arquivo, { type: "array" });
    const planilha = workbook.Sheets[workbook.SheetNames[0]];

    const dados = XLSX.utils.sheet_to_json(planilha, {
      header: 1,
      defval: ""
    });

    const linhas = dados.slice(1);

    produtos = linhas
      .map(linha => ({
        codigo: String(linha[0] || "").trim(),
        ean: String(linha[1] || "").trim(),
        descricao: String(linha[2] || "").trim(),
        embalagem: String(linha[3] || "").trim(),
        qtdMaster: String(linha[4] || "").trim(),
        estoque: Number(linha[5] || 0),
        imagem: String(linha[6] || "").trim(),
        codigoFornecedor: String(linha[7] || "").trim(),
        fornecedor: String(linha[8] || "").trim(),
        novidade: false
      }))
      .filter(produto =>
        produto.codigo ||
        produto.descricao ||
        produto.ean
      );

    marcarNovidades();
    aplicarFiltros();

  } catch (erro) {
    document.getElementById("contador").innerText = "";
    document.getElementById("catalogo").innerHTML =
      "<p>Não foi possível carregar a planilha de produtos.</p>";
    console.error(erro);
  }
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

  return lista;
}

function aplicarFiltroEstoque(lista) {
  const tipo = document.getElementById("filtroEstoque").value;

  if (tipo === "com-estoque") {
    return lista.filter(produto => Number(produto.estoque || 0) > 0);
  }

  if (tipo === "sem-estoque") {
    return lista.filter(produto => Number(produto.estoque || 0) <= 0);
  }

  return lista;
}

function compararCodigo(a, b) {
  const numeroA = codigoNumerico(a);
  const numeroB = codigoNumerico(b);

  if (!isNaN(numeroA) && !isNaN(numeroB)) {
    return numeroA - numeroB;
  }

  return normalizar(a.codigo).localeCompare(normalizar(b.codigo), "pt-BR", { numeric: true });
}

function ordenarProdutos(lista) {
  const tipoOrdenacao = document.getElementById("ordenacao").value;
  const listaOrdenada = [...lista];

  switch (tipoOrdenacao) {
    case "codigo-crescente":
      listaOrdenada.sort((a, b) => compararCodigo(a, b));
      break;

    case "codigo-decrescente":
      listaOrdenada.sort((a, b) => compararCodigo(b, a));
      break;

    case "descricao-az":
      listaOrdenada.sort((a, b) =>
        normalizar(a.descricao).localeCompare(normalizar(b.descricao), "pt-BR")
      );
      break;

    case "descricao-za":
      listaOrdenada.sort((a, b) =>
        normalizar(b.descricao).localeCompare(normalizar(a.descricao), "pt-BR")
      );
      break;

    case "maior-estoque":
      listaOrdenada.sort((a, b) => b.estoque - a.estoque);
      break;
  }

  return listaOrdenada;
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

  pontos += Math.min(estoqueTotal, 10000) * 0.01;

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
    fornecedor.estoqueTotal += Number(produto.estoque || 0);
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

function produtoCombinaPrincipal(produto, buscaPrincipal, permitirBuscaLivre) {
  const busca = normalizar(buscaPrincipal);
  const termos = dividirTermos(buscaPrincipal);

  if (!termos.length) return true;

  const codigo = normalizar(produto.codigo);
  const ean = normalizar(produto.ean);
  const descricao = normalizar(produto.descricao);
  const primeiraPalavra = termos[0];

  const codigoOuEanCombina = termos.every(termo =>
    codigo.includes(termo) || ean.includes(termo)
  );

  if (codigoOuEanCombina) return true;

  const descricaoContemTodos = termos.every(termo =>
    descricao.includes(termo)
  );

  if (!descricaoContemTodos) return false;

  if (permitirBuscaLivre || busca.startsWith("%")) {
    return true;
  }

  return descricao.startsWith(primeiraPalavra);
}

function aplicarBuscaPrincipalComPrioridade(lista, buscaPrincipal) {
  const termos = dividirTermos(buscaPrincipal);

  if (!termos.length) return lista;

  const buscaLivre = normalizar(buscaPrincipal).startsWith("%");

  const resultadoPrioritario = lista.filter(produto =>
    produtoCombinaPrincipal(produto, buscaPrincipal, buscaLivre)
  );

  if (resultadoPrioritario.length > 0) {
    return resultadoPrioritario;
  }

  return lista.filter(produto =>
    produtoCombinaPrincipal(produto, buscaPrincipal, true)
  );
}

function gerarMensagemSemResultado() {
  const buscaPrincipal = document.getElementById("buscaPrincipal").value.trim();
  const buscaCodigoFornecedor = document.getElementById("buscaCodigoFornecedor").value.trim();
  const buscaFornecedor = document.getElementById("buscaFornecedor").value.trim();

  const partes = [];

  if (buscaPrincipal) {
    partes.push(`"${buscaPrincipal}"`);
  }

  if (buscaCodigoFornecedor) {
    partes.push(`código de fornecedor "${buscaCodigoFornecedor}"`);
  }

  if (buscaFornecedor) {
    partes.push(`fornecedor "${buscaFornecedor}"`);
  }

  if (!partes.length) {
    return "Nenhum produto encontrado.";
  }

  if (partes.length === 1) {
    return `Nenhum produto encontrado para ${partes[0]}.`;
  }

  if (partes.length === 2) {
    return `Nenhum produto encontrado para ${partes[0]} dentro de ${partes[1]}.`;
  }

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

  resultado = aplicarBuscaPrincipalComPrioridade(resultado, buscaPrincipal);
  resultado = ordenarProdutos(resultado);

  produtosFiltrados = resultado;
  paginaAtual = 1;

  mostrarProdutos();
}

function mostrarProdutos() {
  const catalogo = document.getElementById("catalogo");
  const contador = document.getElementById("contador");
  const paginacao = document.getElementById("paginacao");

  catalogo.innerHTML = "";
  paginacao.innerHTML = "";

  const totalProdutos = produtosFiltrados.length;
  const totalPaginas = Math.max(1, Math.ceil(totalProdutos / ITENS_POR_PAGINA));

  if (paginaAtual > totalPaginas) {
    paginaAtual = totalPaginas;
  }

  if (totalProdutos === 0) {
    contador.innerText = "0 produtos encontrados";
    catalogo.innerHTML = `<p>${gerarMensagemSemResultado()}</p>`;
    return;
  }

  const inicio = (paginaAtual - 1) * ITENS_POR_PAGINA;
  const fim = inicio + ITENS_POR_PAGINA;
  const produtosDaPagina = produtosFiltrados.slice(inicio, fim);

  contador.innerText = `${totalProdutos} produtos encontrados • Página ${paginaAtual} de ${totalPaginas}`;

  produtosDaPagina.forEach(produto => {
    const card = document.createElement("div");
    card.className = "card";

    const caminhoImagem = produto.imagem
      ? `Imagens/${produto.imagem}`
      : "";

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
    `;

    catalogo.appendChild(card);
  });

  if (totalPaginas > 1) {
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

    paginacao.appendChild(btnAnterior);
    paginacao.appendChild(indicador);
    paginacao.appendChild(btnProxima);
  }
}

function limparFiltros() {
  document.getElementById("buscaPrincipal").value = "";
  document.getElementById("buscaCodigoFornecedor").value = "";
  document.getElementById("buscaFornecedor").value = "";
  document.getElementById("ordenacao").value = "descricao-az";
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

carregarProdutos();