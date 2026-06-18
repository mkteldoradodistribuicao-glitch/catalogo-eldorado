let produtos = [];
let modoAtual = "eldorado";

const CONFIG = {
  eldorado: {
    titulo: "Catálogo Eldorado",
    subtitulo: "Consulte produtos por código, descrição ou EAN.",
    logo: "logos/eldorado.png",
    tema: "tema-eldorado",
    filtroFornecedor: null
  },
  ternura: {
    titulo: "Produtos Ternura",
    subtitulo: "Catálogo exclusivo da linha Produtos Ternura.",
    logo: "logos/produtos-ternura.png",
    tema: "tema-ternura",
    filtroFornecedor: "ternura"
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

  if (!config.filtroFornecedor) {
    return produtos;
  }

  return produtos.filter(produto =>
    normalizar(produto.fornecedor).includes(config.filtroFornecedor)
  );
}

function calcularRelevancia(produto, termos, buscaOriginal) {
  const codigo = normalizar(produto.codigo);
  const ean = normalizar(produto.ean);
  const descricao = normalizar(produto.descricao);
  const embalagem = normalizar(produto.embalagem);

  const textoCompleto = normalizar(`
    ${produto.codigo}
    ${produto.ean}
    ${produto.descricao}
    ${produto.embalagem}
  `);

  let pontos = 0;

  if (!termos.length) {
    return produto.estoque;
  }

  if (codigo === buscaOriginal) pontos += 1000;
  if (ean === buscaOriginal) pontos += 1000;

  if (codigo.includes(buscaOriginal)) pontos += 500;
  if (ean.includes(buscaOriginal)) pontos += 500;

  if (descricao === buscaOriginal) pontos += 400;
  if (descricao.startsWith(buscaOriginal)) pontos += 300;
  if (descricao.includes(buscaOriginal)) pontos += 200;

  termos.forEach(termo => {
    if (codigo.includes(termo)) pontos += 120;
    if (ean.includes(termo)) pontos += 120;
    if (descricao.includes(termo)) pontos += 100;
    if (embalagem.includes(termo)) pontos += 40;
  });

  const todosTermosEncontrados = termos.every(termo =>
    textoCompleto.includes(termo)
  );

  if (todosTermosEncontrados) pontos += 200;

  pontos += produto.estoque * 0.01;

  return pontos;
}

function aplicarFiltros() {
  const campoBusca = document.getElementById("busca");
  const termoDigitado = campoBusca.value;
  const buscaNormalizada = normalizar(termoDigitado);
  const termos = dividirTermos(termoDigitado);

  const base = produtosDoModo();

  let resultado;

  if (!termos.length) {
    resultado = [...base]
      .sort((a, b) => b.estoque - a.estoque)
      .slice(0, 100);
  } else {
    resultado = base
      .map(produto => {
        const textoCompleto = normalizar(`
          ${produto.codigo}
          ${produto.ean}
          ${produto.descricao}
          ${produto.embalagem}
        `);

        const encontrado = termos.every(termo =>
          textoCompleto.includes(termo)
        );

        return {
          ...produto,
          relevancia: encontrado
            ? calcularRelevancia(produto, termos, buscaNormalizada)
            : 0
        };
      })
      .filter(produto => produto.relevancia > 0)
      .sort((a, b) => b.relevancia - a.relevancia)
      .slice(0, 300);
  }

  mostrarProdutos(resultado, termos.length === 0);
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

      <div class="codigo">Código: ${produto.codigo}</div>
      <div class="descricao">${produto.descricao}</div>
      <div class="info">EAN: ${produto.ean || "Não informado"}</div>
      <div class="info">Embalagem: ${produto.embalagem || "Não informada"}</div>
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
  document.getElementById("busca").value = "";

  fecharMenu();
  aplicarFiltros();
}

function abrirFecharMenu() {
  document.getElementById("menuLateral").classList.toggle("aberto");
}

function fecharMenu() {
  document.getElementById("menuLateral").classList.remove("aberto");
}

document.getElementById("busca").addEventListener("input", aplicarFiltros);

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

  if (!menu.contains(event.target) && !botao.contains(event.target)) {
    fecharMenu();
  }
});

carregarProdutos();