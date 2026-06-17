let produtos = [];

async function carregarProdutos() {
  const resposta = await fetch("Produtos.xlsx");
  const arquivo = await resposta.arrayBuffer();

  const workbook = XLSX.read(arquivo, { type: "array" });
  const planilha = workbook.Sheets[workbook.SheetNames[0]];

  const dados = XLSX.utils.sheet_to_json(planilha, {
    header: 1,
    defval: ""
  });

  const linhas = dados.slice(1);

  produtos = linhas.map(linha => ({
    codigo: String(linha[0]).trim(),
    ean: String(linha[1]).trim(),
    descricao: String(linha[2]).trim(),
    embalagem: String(linha[3]).trim(),
    estoque: String(linha[4]).trim(),
    imagem: String(linha[5]).trim(),
    fornecedor: String(linha[6]).trim()
  })).filter(produto => produto.codigo || produto.descricao || produto.ean);

  mostrarProdutos(produtos);
}

function mostrarProdutos(lista) {
  const catalogo = document.getElementById("catalogo");
  const contador = document.getElementById("contador");

  catalogo.innerHTML = "";
  contador.innerText = `${lista.length} produtos encontrados`;

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
      <div class="info">EAN: ${produto.ean}</div>
      <div class="info">Embalagem: ${produto.embalagem}</div>
      <div class="fornecedor">${produto.fornecedor || "Fornecedor não informado"}</div>
    `;

    catalogo.appendChild(card);
  });
}

document.getElementById("busca").addEventListener("input", function () {
  const termo = this.value.toLowerCase().trim();

  const filtrados = produtos.filter(produto =>
    produto.codigo.toLowerCase().includes(termo) ||
    produto.descricao.toLowerCase().includes(termo) ||
    produto.ean.toLowerCase().includes(termo) ||
    produto.fornecedor.toLowerCase().includes(termo)
  );

  mostrarProdutos(filtrados);
});

carregarProdutos();