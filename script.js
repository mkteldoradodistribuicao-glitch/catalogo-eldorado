let produtos = [];
let modoAtual = "eldorado";

const CONFIG = {
  eldorado: {
    titulo: "Catálogo Eldorado",
    subtitulo: "Consulte produtos por código, descrição, EAN, fornecedor ou código do fornecedor.",
    logo: "Logos/Eldorado.png",
    tema: "tema-eldorado",
    filtroTernura: false
  },
  ternura: {
    titulo: "Produtos Ternura",
    subtitulo: "Catálogo exclusivo da linha Produtos Ternura.",
    logo: "Logos/Produtos-ternura.png",
    tema: "tema-ternura",
    filtroTernura: true
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
        estoque: Number(linha[4] || 0),
        imagem: String(linha[5] || "").trim(),
        codigoFornecedor: String(linha[6] || "").trim(),
        fornecedor: String(linha[7] || "").trim()
      }))
      .filter(produto =>
        produto.codigo ||
        produto.descricao ||
        produto.ean
      );

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
    .split(/\s+|%+/)
    .map(t => t.trim())
    .filter(Boolean);
}

function produtosDoModo() {
  const config = CONFIG[modoAtual];

  if (!config.filtroTernura) {
    return produtos;
  }

  return produtos.filter(produto =>
    normalizar(produto.descricao).includes("ternura")
  );
}

function compararCodigo(a, b) {
  const codigoA = normalizar(a.codigo);
  const codigoB = normalizar(b.codigo);

  const numeroA = Number(codigoA);
  const numeroB = Number(codigoB);

  if (!isNaN(numeroA) && !isNaN(numeroB)) {
    return numeroA - numeroB;
  }

  return codigoA.localeCompare(codigoB, "pt-BR", { numeric: true });
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

    default:
      listaOrdenada.sort((a, b) =>
        normalizar(a.descricao).localeCompare(normalizar(b.descricao), "pt-BR")
      );
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

    const distancia = calcularDistancia(fornecedor, termo);
    pontos += Math.max(0, 180 - distancia * 20);
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

  const base = produtosDoModo();
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
      <span>${sugestao.quantidadeProdutos} produtos • Estoque total: ${sugestao.estoqueTotal}</span>
    `;

    item.addEventListener("click", () => {
      document.getElementById("buscaFornecedor").value = sugestao.nome;
      container.classList.remove("ativo");
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
      return distancia <= 2 || palavra.startsWith(termo);
    });
  });
}

function aplicarFiltros() {
  const buscaPrincipal = document.getElementById("buscaPrincipal").value;
  const buscaCodigoFornecedor = document.getElementById("buscaCodigoFornecedor").value;
  const buscaFornecedor = document.getElementById("buscaFornecedor").value;

  const termosPrincipal = dividirTermos(buscaPrincipal);
  const termosCodigoFornecedor = dividirTermos(buscaCodigoFornecedor);
  const termosFornecedor = dividirTermos(buscaFornecedor);

  const base = produtosDoModo();

  const temBusca =
    termosPrincipal.length > 0 ||
    termosCodigoFornecedor.length > 0 ||
    termosFornecedor.length > 0;

  let resultado;

  if (!temBusca) {
    resultado = ordenarProdutos(base).slice(0, 100);
  } else {
    resultado = base.filter(produto => {
      const textoPrincipal = normalizar(`
        ${produto.codigo}
        ${produto.ean}
        ${produto.descricao}
      `);

      const textoCodigoFornecedor = normalizar(produto.codigoFornecedor);

      const encontrouPrincipal = termosPrincipal.every(termo =>
        textoPrincipal.includes(termo)
      );

      const encontrouCodigoFornecedor = termosCodigoFornecedor.every(termo =>
        textoCodigoFornecedor.includes(termo)
      );

      const encontrouFornecedor = fornecedorCombina(produto.fornecedor, buscaFornecedor);

      return (
        encontrouPrincipal &&
        encontrouCodigoFornecedor &&
        encontrouFornecedor
      );
    });

    resultado = ordenarProdutos(resultado).slice(0, 300);
  }

  mostrarProdutos(resultado, !temBusca);
}

function mostrarProdutos(lista, telaInicial = false) {
  const catalogo = document.getElementById("catalogo");
  const contador = document.getElementById("contador");

  catalogo.innerHTML = "";

  contador.innerText = telaInicial
    ? `${lista.length} produtos em destaque`
    : `${lista.length} produtos encontrados`;

  if (lista.length === 0) {
    catalogo.innerHTML = "<p>Nenhum produto encontrado.</p>";
    return;
  }

  lista.forEach(produto => {
    const card = document.createElement("div");
    card.className = "card";

    const caminhoImagem = produto.imagem
      ? `Imagens/${produto.imagem}`
      : "";

    card.innerHTML = `
      ${
        produto.imagem
          ? `<img src="${caminhoImagem}" alt="${produto.descricao}" onerror="this.outerHTML='<div class=&quot;sem-imagem&quot;>Imagem não encontrada</div>'">`
          : `<div class="sem-imagem">Sem imagem</div>`
      }

      <div class="codigo">Código: ${produto.codigo || "Não informado"}</div>
      <div class="descricao">${produto.descricao || "Descrição não informada"}</div>
      <div class="info">EAN: ${produto.ean || "Não informado"}</div>
      <div class="info">Embalagem: ${produto.embalagem || "Não informada"}</div>
      <div class="info">Estoque: ${produto.estoque || 0}</div>
      <div class="info">Código Fornecedor: ${produto.codigoFornecedor || "Não informado"}</div>
      <div class="fornecedor">${produto.fornecedor || "Fornecedor não informado"}</div>
    `;

    catalogo.appendChild(card);
  });
}

function trocarModo(novoModo) {
  modoAtual = novoModo;

  const config = CONFIG[modoAtual];

  document.body.className = config.tema;
  document.getElementById("logoCatalogo").src = config.logo;
  document.getElementById("tituloCatalogo").innerText = config.titulo;
  document.getElementById("subtituloCatalogo").innerText = config.subtitulo;

  document.getElementById("buscaPrincipal").value = "";
  document.getElementById("buscaCodigoFornecedor").value = "";
  document.getElementById("buscaFornecedor").value = "";

  document.getElementById("sugestoesFornecedor").classList.remove("ativo");

  fecharMenu();
  aplicarFiltros();
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