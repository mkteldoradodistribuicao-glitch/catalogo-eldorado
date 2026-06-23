from pathlib import Path

base = Path("/mnt/data/catalogo_eldorado_codigos_completos")
base.mkdir(exist_ok=True)

index_html = r'''<!DOCTYPE html>
<html lang="pt-BR">
<head>
  <meta charset="UTF-8" />
  <meta name="viewport" content="width=device-width, initial-scale=1.0" />

  <title>Catálogo Eldorado</title>

  <link rel="icon" type="image/png" href="Logos/Favicon.png" />
  <link rel="stylesheet" href="style.css" />
</head>
<body class="tema-eldorado">

  <header class="topo">
    <div class="topo-conteudo">
      <img id="logoCatalogo" src="Logos/Eldorado.png" alt="Logo do catálogo" class="logo" />

      <div class="titulo-area">
        <h1 id="tituloCatalogo">Catálogo Eldorado</h1>
        <p id="subtituloCatalogo">Consulte produtos por código, descrição, EAN, fornecedor ou código do fornecedor.</p>

        <div class="aviso-catalogo">
          Catálogo ilustrativo. Não caracteriza Pedido de Compra. A disponibilidade, preços e condições comerciais devem ser confirmados com o RCA antes da finalização.
        </div>
      </div>

      <button id="btnCotacao" class="cotacao-btn" type="button" aria-label="Abrir cotação">
        <span class="icone-carrinho">🛒</span>
        <span id="contadorCotacao">0</span>
      </button>

      <button id="btnMenu" class="menu-btn" type="button" aria-label="Abrir menu">☰</button>
    </div>

    <nav id="menuLateral" class="menu-lateral">
      <button id="btnCatalogoEldorado" type="button">Catálogo Eldorado</button>
      <button id="btnProdutosTernura" type="button">Produtos Ternura</button>
      <button id="btnNovidades" type="button">Novidades</button>
    </nav>
  </header>

  <main class="container">
    <section class="filtros-area">
      <div class="busca-area">
        <input type="text" id="buscaPrincipal" placeholder="Buscar por código do produto, descrição ou EAN..." autocomplete="off" />
        <input type="text" id="buscaCodigoFornecedor" placeholder="Código Fornecedor" autocomplete="off" />

        <div class="campo-fornecedor">
          <input type="text" id="buscaFornecedor" placeholder="Fornecedor" autocomplete="off" />
          <div id="sugestoesFornecedor" class="sugestoes-fornecedor"></div>
        </div>
      </div>

      <div class="controles-area">
        <div class="ordenacao-area">
          <label for="ordenacao">Ordenar por:</label>
          <select id="ordenacao">
            <option value="descricao-az">Descrição A-Z</option>
            <option value="descricao-za">Descrição Z-A</option>
            <option value="codigo-crescente">Código do Produto Crescente</option>
            <option value="codigo-decrescente">Código do Produto Decrescente</option>
            <option value="maior-estoque" selected>Maior Estoque</option>
          </select>
        </div>

        <div class="estoque-area">
          <label for="filtroEstoque">Estoque:</label>
          <select id="filtroEstoque">
            <option value="com-estoque" selected>Apenas Produtos com Estoque</option>
            <option value="sem-estoque">Apenas Produtos sem Estoque</option>
            <option value="todos">Todos os Produtos</option>
          </select>
        </div>

        <button id="btnLimparFiltros" class="btn-limpar" type="button">Limpar Filtros</button>
      </div>

      <div id="paginacaoSuperior" class="paginacao paginacao-superior"></div>
    </section>

    <div id="contador"></div>

    <section id="catalogo"></section>

    <div id="paginacao" class="paginacao"></div>
  </main>

  <aside id="painelCotacao" class="painel-cotacao">
    <div class="painel-cabecalho">
      <h2>Lista de Cotação</h2>
      <button id="fecharCotacao" type="button">×</button>
    </div>

    <div class="dados-pdf">
      <label for="nomePdf">Razão Social da loja ou Comprador responsável:</label>
      <input type="text" id="nomePdf" placeholder="Ex.: Supermercado São João" autocomplete="off" />

      <label for="cnpjCliente">CNPJ do Cliente:</label>
      <input type="text" id="cnpjCliente" placeholder="00.000.000/0000-00" maxlength="18" autocomplete="off" />
    </div>

    <div id="listaCotacao" class="lista-cotacao"></div>

    <div class="acoes-cotacao">
      <button id="limparCotacao" type="button" class="botao-perigo">Limpar carrinho</button>
      <button id="gerarPdf" type="button">Gerar Cotação em PDF</button>
      <button id="copiarCotacao" type="button">Copiar cotação</button>
    </div>
  </aside>

  <div id="overlayCotacao" class="overlay-cotacao"></div>

  <div id="toast" class="toast"></div>

  <div class="botoes-flutuantes">
    <button id="btnCarrinhoFlutuante" class="btn-flutuante btn-carrinho-flutuante" type="button" aria-label="Abrir carrinho">
      🛒 <span id="contadorCotacaoFlutuante">0</span>
    </button>

    <button id="btnTopo" class="btn-flutuante btn-topo" type="button" aria-label="Voltar ao topo">
      ↑
    </button>
  </div>

  <script src="https://cdn.jsdelivr.net/npm/xlsx/dist/xlsx.full.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/jspdf@2.5.1/dist/jspdf.umd.min.js"></script>
  <script src="https://cdn.jsdelivr.net/npm/jspdf-autotable@3.8.2/dist/jspdf.plugin.autotable.min.js"></script>
  <script src="script.js"></script>
</body>
</html>
'''

style_css = r'''* {
  box-sizing: border-box;
}

:root {
  --cor-principal: #009639;
  --cor-secundaria: #ff6b1a;
  --cor-fundo: #f4f6f4;
  --cor-card: #ffffff;
  --cor-texto: #222222;
  --cor-suave: #e9f7ee;
}

body {
  margin: 0;
  font-family: Arial, sans-serif;
  background:
    radial-gradient(circle at top left, rgba(0, 150, 57, 0.10), transparent 32%),
    radial-gradient(circle at bottom right, rgba(255, 107, 26, 0.10), transparent 30%),
    linear-gradient(180deg, var(--cor-fundo), #ffffff);
  color: var(--cor-texto);
  min-height: 100vh;
}

body.cotacao-aberta {
  overflow: hidden;
}

body.tema-eldorado {
  --cor-principal: #009639;
  --cor-secundaria: #ff6b1a;
  --cor-fundo: #f4f6f4;
  --cor-suave: #e9f7ee;
}

body.tema-ternura {
  --cor-principal: #002f5f;
  --cor-secundaria: #0b74bd;
  --cor-fundo: #f3f8fc;
  --cor-suave: #e8f3fb;

  background:
    radial-gradient(circle at top left, rgba(0, 47, 95, 0.12), transparent 32%),
    radial-gradient(circle at bottom right, rgba(11, 116, 189, 0.10), transparent 30%),
    linear-gradient(180deg, var(--cor-fundo), #ffffff);
}

.topo {
  background:
    radial-gradient(circle at top left, rgba(255,255,255,0.18), transparent 30%),
    linear-gradient(135deg, var(--cor-principal), var(--cor-secundaria));
  color: white;
  padding: 30px 16px 34px;
  position: relative;
  box-shadow: 0 8px 24px rgba(0,0,0,0.16);
}

.topo-conteudo {
  max-width: 1200px;
  margin: 0 auto;
  display: grid;
  grid-template-columns: 180px 1fr 58px 58px;
  align-items: center;
  gap: 18px;
}

.logo {
  height: 120px;
  max-width: 180px;
  object-fit: contain;
  background: rgba(255,255,255,0.96);
  border-radius: 18px;
  padding: 12px;
  box-shadow: 0 6px 18px rgba(0,0,0,0.16);
}

.titulo-area {
  text-align: center;
}

.titulo-area h1 {
  margin: 0;
  font-size: 34px;
  font-weight: 800;
  letter-spacing: 0.4px;
}

.titulo-area p {
  margin: 8px auto 0;
  font-size: 15px;
  opacity: 0.96;
  max-width: 680px;
  line-height: 1.4;
}

.aviso-catalogo {
  margin: 14px auto 0;
  max-width: 820px;
  background: rgba(255,255,255,0.16);
  border: 1px solid rgba(255,255,255,0.28);
  border-radius: 12px;
  padding: 10px 14px;
  font-size: 12px;
  line-height: 1.4;
}

.cotacao-btn,
.menu-btn,
.btn-flutuante {
  background: rgba(255,255,255,0.96);
  color: var(--cor-principal);
  border: none;
  width: 52px;
  height: 52px;
  border-radius: 14px;
  cursor: pointer;
  font-weight: bold;
  box-shadow: 0 6px 16px rgba(0,0,0,0.16);
  transition: transform 0.2s ease, box-shadow 0.2s ease;
}

.cotacao-btn,
.btn-carrinho-flutuante {
  font-size: 20px;
  position: relative;
}

.icone-carrinho {
  display: block;
  line-height: 1;
}

#contadorCotacao,
#contadorCotacaoFlutuante {
  position: absolute;
  top: -7px;
  right: -7px;
  min-width: 22px;
  height: 22px;
  padding: 0 6px;
  background: #d92525;
  color: white;
  border-radius: 999px;
  font-size: 12px;
  display: flex;
  align-items: center;
  justify-content: center;
}

.menu-btn {
  font-size: 28px;
}

.cotacao-btn:hover,
.menu-btn:hover,
.btn-flutuante:hover {
  transform: translateY(-2px);
  box-shadow: 0 10px 22px rgba(0,0,0,0.22);
}

.menu-lateral {
  display: none;
  position: absolute;
  top: 128px;
  right: 20px;
  width: 230px;
  background: white;
  border-radius: 14px;
  box-shadow: 0 8px 24px rgba(0,0,0,0.18);
  overflow: hidden;
  z-index: 20;
}

.menu-lateral.aberto {
  display: block;
}

.menu-lateral button {
  width: 100%;
  padding: 16px;
  background: white;
  border: none;
  text-align: left;
  cursor: pointer;
  font-size: 15px;
  color: #222;
  border-bottom: 1px solid #eee;
}

.menu-lateral button:hover {
  background: var(--cor-suave);
  color: var(--cor-principal);
  font-weight: bold;
}

.container {
  max-width: 1200px;
  margin: 26px auto;
  padding: 0 16px;
}

.filtros-area {
  margin-bottom: 14px;
  padding: 18px;
  border-radius: 18px;
  background: rgba(255,255,255,0.78);
  box-shadow: 0 6px 18px rgba(0,0,0,0.06);
  backdrop-filter: blur(4px);
}

.busca-area {
  display: grid;
  grid-template-columns: 2fr 0.7fr 1fr;
  gap: 12px;
  margin-bottom: 12px;
  align-items: start;
}

.busca-area input {
  width: 100%;
  padding: 15px 18px;
  font-size: 16px;
  border: 1px solid #cfcfcf;
  border-radius: 12px;
  outline: none;
  background: white;
}

.busca-area input:focus {
  border-color: var(--cor-secundaria);
  box-shadow: 0 0 0 3px rgba(0,0,0,0.06);
}

.campo-fornecedor {
  position: relative;
}

.sugestoes-fornecedor {
  display: none;
  position: absolute;
  top: 54px;
  left: 0;
  right: 0;
  background: white;
  border-radius: 12px;
  box-shadow: 0 8px 24px rgba(0,0,0,0.18);
  overflow: hidden;
  z-index: 30;
}

.sugestoes-fornecedor.ativo {
  display: block;
}

.sugestao-item {
  width: 100%;
  padding: 12px 14px;
  background: white;
  border: none;
  border-bottom: 1px solid #eeeeee;
  text-align: left;
  cursor: pointer;
}

.sugestao-item:hover {
  background: var(--cor-suave);
}

.sugestao-item strong {
  display: block;
  color: var(--cor-principal);
  font-size: 14px;
}

.sugestao-item span {
  display: block;
  color: #666;
  font-size: 12px;
  margin-top: 3px;
}

.controles-area {
  display: flex;
  justify-content: flex-end;
  gap: 14px;
  flex-wrap: wrap;
}

.ordenacao-area,
.estoque-area {
  display: flex;
  align-items: center;
  gap: 10px;
}

.ordenacao-area label,
.estoque-area label {
  font-size: 14px;
  font-weight: bold;
  color: var(--cor-principal);
}

.ordenacao-area select,
.estoque-area select {
  padding: 12px 14px;
  border: 1px solid #cfcfcf;
  border-radius: 10px;
  background: white;
  color: #222;
  font-size: 14px;
  outline: none;
  cursor: pointer;
}

.btn-limpar {
  padding: 12px 18px;
  border: none;
  border-radius: 10px;
  background: var(--cor-principal);
  color: white;
  font-size: 14px;
  font-weight: bold;
  cursor: pointer;
}

#contador {
  margin-bottom: 18px;
  font-size: 14px;
  color: #555;
}

#catalogo {
  display: grid;
  grid-template-columns: repeat(auto-fill, minmax(220px, 1fr));
  gap: 18px;
}

.card {
  background: rgba(255, 255, 255, 0.96);
  border-radius: 18px;
  padding: 16px;
  box-shadow: 0 6px 18px rgba(0,0,0,0.08);
  transition: transform 0.2s ease, box-shadow 0.2s ease;
  border-top: 4px solid var(--cor-principal);
  backdrop-filter: blur(4px);
  position: relative;
}

.card:hover {
  transform: translateY(-3px);
  box-shadow: 0 10px 24px rgba(0,0,0,0.12);
}

.card.produto-adicionado {
  animation: efeitoAdicionado 0.75s ease;
  border-top-color: var(--cor-secundaria);
}

.card.produto-adicionado::after {
  content: "✓ Adicionado";
  position: absolute;
  top: 12px;
  left: 12px;
  background: var(--cor-principal);
  color: white;
  padding: 7px 11px;
  border-radius: 999px;
  font-size: 12px;
  font-weight: bold;
  box-shadow: 0 5px 14px rgba(0,0,0,0.18);
  z-index: 5;
}

@keyframes efeitoAdicionado {
  0% {
    transform: scale(1);
    box-shadow: 0 6px 18px rgba(0,0,0,0.08);
  }

  35% {
    transform: scale(1.025);
    box-shadow: 0 0 0 4px rgba(0, 150, 57, 0.16), 0 12px 28px rgba(0,0,0,0.16);
  }

  100% {
    transform: scale(1);
    box-shadow: 0 6px 18px rgba(0,0,0,0.08);
  }
}

.card-imagem {
  position: relative;
  background: #ffffff;
  border-radius: 12px;
  padding: 10px;
  margin-bottom: 12px;
  height: 200px;
  display: flex;
  align-items: center;
  justify-content: center;
  overflow: hidden;
}

.card-imagem img {
  width: 100%;
  height: 180px;
  object-fit: contain;
}

.selo-novidade {
  position: absolute;
  top: 10px;
  right: 10px;
  background: var(--cor-secundaria);
  color: white;
  padding: 6px 10px;
  border-radius: 999px;
  font-size: 11px;
  font-weight: bold;
  letter-spacing: 0.5px;
  box-shadow: 0 4px 10px rgba(0,0,0,0.18);
  z-index: 2;
}

.codigo {
  font-weight: bold;
  color: var(--cor-principal);
  font-size: 15px;
}

.descricao {
  font-size: 15px;
  font-weight: bold;
  margin: 8px 0;
  min-height: 44px;
  line-height: 1.3;
}

.info {
  font-size: 13px;
  color: #555;
  margin: 4px 0;
}

.fornecedor {
  margin-top: 10px;
  display: inline-block;
  padding: 6px 10px;
  background: var(--cor-suave);
  color: var(--cor-principal);
  border-radius: 20px;
  font-size: 12px;
  font-weight: bold;
}

.sem-imagem {
  height: 180px;
  width: 100%;
  display: flex;
  align-items: center;
  justify-content: center;
  background: #eef1f4;
  color: #777;
  border-radius: 10px;
  font-size: 13px;
}

.btn-adicionar-cotacao {
  width: 100%;
  margin-top: 14px;
  padding: 11px 12px;
  border: none;
  border-radius: 12px;
  background: linear-gradient(135deg, var(--cor-principal), var(--cor-secundaria));
  color: white;
  font-size: 14px;
  font-weight: bold;
  cursor: pointer;
  box-shadow: 0 5px 14px rgba(0,0,0,0.14);
  transition: transform 0.2s ease, box-shadow 0.2s ease;
}

.btn-adicionar-cotacao:hover {
  transform: translateY(-2px);
  box-shadow: 0 8px 20px rgba(0,0,0,0.18);
}

.paginacao {
  margin: 28px 0 10px;
  display: flex;
  align-items: center;
  justify-content: center;
  gap: 14px;
  flex-wrap: wrap;
}

.paginacao-superior {
  margin: 16px 0 0;
}

.paginacao button {
  padding: 11px 18px;
  border: none;
  border-radius: 10px;
  background: var(--cor-principal);
  color: white;
  font-weight: bold;
  cursor: pointer;
}

.paginacao button:disabled {
  background: #cccccc;
  cursor: not-allowed;
}

.paginacao span {
  font-size: 14px;
  font-weight: bold;
  color: var(--cor-principal);
}

.painel-cotacao {
  position: fixed;
  top: 0;
  right: -430px;
  width: 410px;
  max-width: 94vw;
  height: 100vh;
  background: white;
  z-index: 100;
  padding: 18px;
  box-shadow: -8px 0 24px rgba(0,0,0,0.22);
  transition: right 0.25s ease;
  display: flex;
  flex-direction: column;
  overflow: hidden;
}

.painel-cotacao.aberto {
  right: 0;
}

.painel-cabecalho {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 16px;
  border-bottom: 1px solid #eee;
  padding-bottom: 12px;
}

.painel-cabecalho h2 {
  margin: 0;
  color: var(--cor-principal);
  font-size: 22px;
}

.painel-cabecalho button {
  border: none;
  background: var(--cor-suave);
  color: var(--cor-principal);
  width: 38px;
  height: 38px;
  border-radius: 10px;
  font-size: 24px;
  cursor: pointer;
}

.dados-pdf {
  margin: 16px 0 12px;
}

.dados-pdf label {
  display: block;
  font-size: 13px;
  font-weight: bold;
  color: var(--cor-principal);
  margin: 10px 0 6px;
}

.dados-pdf input {
  width: 100%;
  padding: 13px 14px;
  border: 1px solid #cfcfcf;
  border-radius: 12px;
  outline: none;
}

.lista-cotacao {
  overflow-y: auto;
  overscroll-behavior: contain;
  touch-action: pan-y;
  flex: 1;
  padding-right: 4px;
}

.item-cotacao {
  border: 1px solid #eeeeee;
  border-left: 4px solid var(--cor-principal);
  border-radius: 14px;
  padding: 12px;
  margin-bottom: 12px;
  background: #fff;
  box-shadow: 0 4px 12px rgba(0,0,0,0.05);
}

.item-cotacao strong {
  color: var(--cor-principal);
  font-size: 14px;
}

.item-cotacao p {
  margin: 6px 0;
  font-size: 13px;
  color: #444;
}

.item-cotacao-acoes {
  display: flex;
  align-items: end;
  justify-content: space-between;
  gap: 10px;
  margin-top: 10px;
}

.item-cotacao-acoes label {
  font-size: 12px;
  font-weight: bold;
  color: #555;
}

.item-cotacao-acoes input {
  display: block;
  width: 120px;
  margin-top: 5px;
  padding: 9px;
  border: 1px solid #cfcfcf;
  border-radius: 10px;
}

.item-cotacao-acoes button {
  border: none;
  background: #f3e4e4;
  color: #9c1b1b;
  padding: 10px 12px;
  border-radius: 10px;
  cursor: pointer;
  font-weight: bold;
}

.aviso-estoque {
  margin-top: 8px;
  padding: 8px 10px;
  background: #fff7df;
  border: 1px solid #f2d28a;
  color: #7a5200;
  border-radius: 10px;
  font-size: 12px;
}

.acoes-cotacao {
  display: grid;
  gap: 10px;
  border-top: 1px solid #eee;
  padding-top: 14px;
  margin-top: 12px;
}

.acoes-cotacao button {
  padding: 13px 14px;
  border: none;
  border-radius: 12px;
  background: var(--cor-principal);
  color: white;
  font-weight: bold;
  cursor: pointer;
}

.acoes-cotacao .botao-perigo {
  background: #b82323;
}

.overlay-cotacao {
  display: none;
  position: fixed;
  inset: 0;
  background: rgba(0,0,0,0.35);
  z-index: 90;
}

.overlay-cotacao.aberto {
  display: block;
}

.toast {
  position: fixed;
  right: 22px;
  bottom: 96px;
  max-width: 330px;
  background: #1f2933;
  color: white;
  padding: 13px 16px;
  border-radius: 14px;
  box-shadow: 0 8px 22px rgba(0,0,0,0.22);
  font-size: 14px;
  opacity: 0;
  transform: translateY(12px);
  pointer-events: none;
  transition: opacity 0.25s ease, transform 0.25s ease;
  z-index: 150;
}

.toast.visivel {
  opacity: 1;
  transform: translateY(0);
}

.botoes-flutuantes {
  position: fixed;
  right: 18px;
  bottom: 18px;
  display: flex;
  gap: 10px;
  z-index: 70;
}

.btn-flutuante {
  display: none;
}

.btn-flutuante.visivel {
  display: block;
}

.btn-carrinho-flutuante {
  background: var(--cor-principal);
  color: white;
}

.btn-topo {
  background: white;
  color: var(--cor-principal);
  font-size: 24px;
}

@media (max-width: 700px) {
  .topo {
    padding: 22px 14px 26px;
  }

  .topo-conteudo {
    grid-template-columns: 1fr 48px 48px;
    gap: 10px;
  }

  .logo {
    grid-column: 1 / 4;
    justify-self: center;
    height: 90px;
    max-width: 210px;
  }

  .titulo-area {
    grid-column: 1 / 4;
    text-align: center;
  }

  .titulo-area h1 {
    font-size: 24px;
  }

  .titulo-area p,
  .aviso-catalogo {
    font-size: 12px;
  }

  .cotacao-btn,
  .menu-btn,
  .btn-flutuante {
    width: 48px;
    height: 48px;
  }

  .menu-btn {
    font-size: 24px;
  }

  .menu-lateral {
    top: 205px;
    right: 12px;
  }

  .busca-area {
    grid-template-columns: 1fr;
  }

  .controles-area {
    flex-direction: column;
  }

  .ordenacao-area,
  .estoque-area {
    align-items: stretch;
    flex-direction: column;
  }

  .ordenacao-area select,
  .estoque-area select,
  .btn-limpar {
    width: 100%;
  }

  #catalogo {
    grid-template-columns: repeat(auto-fill, minmax(160px, 1fr));
    gap: 14px;
  }

  .card-imagem {
    height: 160px;
  }

  .card-imagem img,
  .sem-imagem {
    height: 140px;
  }

  .paginacao {
    flex-direction: column;
  }

  .paginacao button {
    width: 100%;
  }

  .painel-cotacao {
    width: 94vw;
  }

  .toast {
    left: 14px;
    right: 14px;
    bottom: 88px;
  }
}
'''

script_js = r'''let produtos = [];
let produtosFiltrados = [];
let cotacao = [];
let modoAtual = "eldorado";
let paginaAtual = 1;

const ITENS_POR_PAGINA = 150;
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
        novidade: false,
        relevancia: 999
      }))
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
  const numero = Number(String(valor || "").replace(",", ".").replace(/[^\d.]/g, ""));
  return numero > 0 ? numero : 1;
}

function ajustarParaMultiplo(quantidade, qtdMaster) {
  const master = qtdMasterNumerica(qtdMaster);
  const qtd = Math.max(master, Number(quantidade || master));
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
          return b.estoque - a.estoque;
        default:
          return compararCodigo(a, b);
      }
    });

    return listaOrdenada;
  }

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
    default:
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
    itemExistente.quantidade = ajustarParaMultiplo(Number(itemExistente.quantidade || master) + master, master);
  } else {
    cotacao.push({
      codigo: produto.codigo,
      descricao: produto.descricao,
      ean: produto.ean,
      quantidade: master,
      qtdMaster: master,
      estoque: Number(produto.estoque || 0)
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
  const estoque = Number(item.estoque || 0);
  const master = qtdMasterNumerica(item.qtdMaster);
  const quantidade = Number(item.quantidade || 0);

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
  const totalItens = cotacao.reduce((soma, item) => soma + Number(item.quantidade || 0), 0);

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
  document.getElementById("ordenacao").value = modoAtual === "novidades" ? "codigo-decrescente" : "maior-estoque";

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

document.getElementById("ordenacao").value = "maior-estoque";
carregarProdutos();
'''

(index_path := base / "index.html").write_text(index_html, encoding="utf-8")
(style_path := base / "style.css").write_text(style_css, encoding="utf-8")
(script_path := base / "script.js").write_text(script_js, encoding="utf-8")

print("Arquivos criados:")
print(index_path)
print(style_path)
print(script_path)
