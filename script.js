const $ = id => document.getElementById(id);
const porPagina = 150;
const ocultosTernura = ["93217", "89817"];

let produtos = [], filtrados = [], cotacao = [], pagina = 1, catalogo = "eldorado";

const col = {
  codigo:0, ean:1, descricao:2, embalagem:3, master:4,
  estoque:5, imagem:6, codFornecedor:7, fornecedor:8
};

const norm = v => String(v ?? "").normalize("NFD").replace(/[\u0300-\u036f]/g,"").toLowerCase().trim();
const num = v => Number(String(v ?? "0").replace(",", ".")) || 0;
const txt = v => String(v ?? "").trim();

fetch("Produtos.xlsx")
  .then(r => r.arrayBuffer())
  .then(b => {
    const wb = XLSX.read(b);
    const ws = wb.Sheets[wb.SheetNames[0]];
    produtos = XLSX.utils.sheet_to_json(ws, { header:1 }).slice(1).map(l => ({
      codigo:txt(l[col.codigo]),
      ean:txt(l[col.ean]),
      descricao:txt(l[col.descricao]),
      embalagem:txt(l[col.embalagem]),
      master:num(l[col.master]) || 1,
      estoque:num(l[col.estoque]),
      imagem:txt(l[col.imagem]),
      codFornecedor:txt(l[col.codFornecedor]),
      fornecedor:txt(l[col.fornecedor])
    }));
    aplicarFiltros();
  });

function baseCatalogo() {
  return produtos.filter(p => {
    if (catalogo === "eldorado") return true;
    return norm(p.descricao).includes("ternura") && !ocultosTernura.includes(p.codigo);
  });
}

function buscaLivre(p, termo) {
  const campos = [p.codigo, p.ean, p.descricao].map(norm).join(" ");
  return norm(termo).split("%").every(t => campos.includes(t));
}

function buscaDescricaoPrioritaria(lista, termo) {
  const n = norm(termo);
  if (!n) return lista;
  if (n.includes("%")) return lista.filter(p => buscaLivre(p, n));

  const primeira = n.split(/\s+/)[0];
  const campos = p => [p.codigo, p.ean, p.descricao].map(norm).join(" ");
  const inicio = lista.filter(p => norm(p.descricao).startsWith(primeira) && campos(p).includes(n));
  return inicio.length ? inicio : lista.filter(p => campos(p).includes(n));
}

function aplicarFiltros() {
  let lista = baseCatalogo();
  const busca = $("busca").value;
  const cf = norm($("codFornecedor").value);
  const forn = norm($("fornecedor").value);
  const est = $("estoqueFiltro").value;

  lista = buscaDescricaoPrioritaria(lista, busca);
  if (cf) lista = lista.filter(p => norm(p.codFornecedor).includes(cf));
  if (forn) lista = lista.filter(p => norm(p.fornecedor).includes(forn));
  if (est === "com") lista = lista.filter(p => p.estoque > 0);
  if (est === "sem") lista = lista.filter(p => p.estoque <= 0);

  ordenar(lista);
  filtrados = lista;
  pagina = 1;
  render();
}

function ordenar(lista) {
  const ord = $("ordenacao").value;
  const porCod = (a,b) => num(a.codigo) - num(b.codigo);
  const porDesc = (a,b) => a.descricao.localeCompare(b.descricao);

  ({
    codigoAsc:() => lista.sort(porCod),
    codigoDesc:() => lista.sort((a,b) => porCod(b,a)),
    descAsc:() => lista.sort(porDesc),
    descDesc:() => lista.sort((a,b) => porDesc(b,a)),
    estoqueDesc:() => lista.sort((a,b) => b.estoque - a.estoque)
  }[ord])();
}

function render() {
  const totalPag = Math.max(1, Math.ceil(filtrados.length / porPagina));
  pagina = Math.min(pagina, totalPag);
  const ini = (pagina - 1) * porPagina;
  const lista = filtrados.slice(ini, ini + porPagina);

  $("produtos").innerHTML = lista.map(card).join("");
  $("paginaInfo").textContent = `Página ${pagina} de ${totalPag}`;
  $("anterior").disabled = pagina === 1;
  $("proxima").disabled = pagina === totalPag;
  mensagem();
}

function card(p) {
  const img = p.imagem ? `Imagens/${p.imagem}` : "";
  return `
    <article class="card" id="card-${p.codigo}">
      <div class="imgBox">
        ${img ? `<img src="${img}" alt="${p.descricao}" onerror="this.style.display='none'">` : ""}
      </div>
      <h3>${p.descricao}</h3>
      <div class="info">
        <strong>Código:</strong> ${p.codigo}<br>
        <strong>EAN:</strong> ${p.ean}<br>
        <strong>Embalagem:</strong> ${p.embalagem}<br>
        <strong>QTD Master:</strong> ${p.master}<br>
        <strong>Cód. Fornecedor:</strong> ${p.codFornecedor}<br>
        <strong>Fornecedor:</strong> ${p.fornecedor}
      </div>
      <button class="addBtn" onclick="addCotacao('${p.codigo}', this)">Adicionar à Cotação</button>
    </article>`;
}

function mensagem() {
  const b = txt($("busca").value), f = txt($("fornecedor").value), c = txt($("codFornecedor").value);
  if (filtrados.length) return $("mensagem").textContent = "";

  let partes = [];
  if (b) partes.push(`"${b}"`);
  if (f) partes.push(`fornecedor ${f}`);
  if (c) partes.push(`código fornecedor ${c}`);

  $("mensagem").textContent = partes.length
    ? `Nenhum produto encontrado para ${partes.join(" dentro de ")}.`
    : "Nenhum produto encontrado.";
}

function addCotacao(codigo, btn) {
  const p = produtos.find(x => x.codigo === codigo);
  const item = cotacao.find(x => x.codigo === codigo);

  item ? item.qtd += p.master : cotacao.push({...p, qtd:p.master});

  btn.classList.add("confirmado");
  btn.textContent = "Adicionado!";
  setTimeout(() => {
    btn.classList.remove("confirmado");
    btn.textContent = "Adicionar à Cotação";
  }, 900);

  const c = $(`card-${codigo}`);
  c.classList.add("added");
  setTimeout(() => c.classList.remove("added"), 900);

  toast(`Produto adicionado à cotação. ${cotacao.length} item(ns) na cotação.`);
  renderCotacao();
}

function renderCotacao() {
  $("cotacaoQtd").textContent = cotacao.length;
  $("cotacaoLista").innerHTML = cotacao.length ? cotacao.map(itemCotacao).join("") : "<p>Nenhum produto adicionado.</p>";
}

function itemCotacao(p) {
  const restantes = p.estoque - p.qtd;
  const mastersRestantes = Math.floor(restantes / p.master);
  let aviso = "";

  if (p.qtd > p.estoque) {
    aviso = `<div class="aviso erro">❌ Quantidade cotada superior ao estoque disponível.</div>`;
  } else if (mastersRestantes <= 3) {
    aviso = `<div class="aviso alerta">⚠️ Verifique a disponibilidade do estoque antes de concluir a cotação.</div>`;
  }

  return `
    <div class="itemCotacao">
      <h4>${p.descricao}</h4>
      <small>Cód.: ${p.codigo} | Master: ${p.master}</small>
      <div class="qtdBox">
        <button onclick="alterarQtd('${p.codigo}', -1)">−</button>
        <input value="${p.qtd}" onchange="qtdManual('${p.codigo}', this.value)">
        <button onclick="alterarQtd('${p.codigo}', 1)">+</button>
      </div>
      <small>${p.qtd / p.master} master(s)</small>
      ${aviso}
      <button onclick="removerCotacao('${p.codigo}')">Remover</button>
    </div>`;
}

function alterarQtd(codigo, mult) {
  const item = cotacao.find(p => p.codigo === codigo);
  item.qtd += item.master * mult;
  if (item.qtd <= 0) removerCotacao(codigo);
  renderCotacao();
}

function qtdManual(codigo, valor) {
  const item = cotacao.find(p => p.codigo === codigo);
  let qtd = Math.max(item.master, num(valor));
  qtd = Math.ceil(qtd / item.master) * item.master;
  item.qtd = qtd;
  renderCotacao();
}

function removerCotacao(codigo) {
  cotacao = cotacao.filter(p => p.codigo !== codigo);
  renderCotacao();
}

function toast(msg) {
  $("toast").textContent = msg;
  $("toast").classList.add("ativo");
  clearTimeout(window.toastTimer);
  window.toastTimer = setTimeout(() => $("toast").classList.remove("ativo"), 2600);
}

function sugestoesFornecedor() {
  const termo = norm($("fornecedor").value);
  if (!termo) return $("sugestoes").innerHTML = "";

  const mapa = {};
  baseCatalogo().forEach(p => {
    const f = p.fornecedor;
    if (!f) return;
    if (!mapa[f]) mapa[f] = { nome:f, pontos:0, estoque:0 };
    mapa[f].estoque += p.estoque;
    mapa[f].pontos += norm(f).includes(termo) ? 10 : 0;
  });

  const lista = Object.values(mapa)
    .filter(x => norm(x.nome).includes(termo))
    .sort((a,b) => (b.pontos + b.estoque/1000) - (a.pontos + a.estoque/1000))
    .slice(0,5);

  $("sugestoes").innerHTML = lista.map(x => `<div class="sugestao" onclick="selecionarFornecedor('${x.nome.replace(/'/g,"\\'")}')">${x.nome}</div>`).join("");
}

function selecionarFornecedor(nome) {
  $("fornecedor").value = nome;
  $("sugestoes").innerHTML = "";
  aplicarFiltros();
}

function trocarCatalogo(tipo) {
  catalogo = tipo;
  document.body.classList.toggle("ternura", tipo === "ternura");

  $("logo").src = tipo === "ternura" ? "Logos/Produtos-ternura.png" : "Logos/Eldorado.png";
  $("titulo").textContent = tipo === "ternura" ? "Produtos Ternura" : "Catálogo Eldorado";
  $("subtitulo").textContent = tipo === "ternura"
    ? "Catálogo exclusivo da linha Produtos Ternura."
    : "Consulte produtos por código, descrição, EAN, fornecedor ou código do fornecedor.";

  $("menu").classList.remove("ativo");
  aplicarFiltros();
}

["busca","codFornecedor","fornecedor","estoqueFiltro","ordenacao"].forEach(id => {
  $(id).addEventListener("input", aplicarFiltros);
  $(id).addEventListener("change", aplicarFiltros);
});

$("fornecedor").addEventListener("input", sugestoesFornecedor);
$("limpar").onclick = () => {
  ["busca","codFornecedor","fornecedor"].forEach(id => $(id).value = "");
  $("estoqueFiltro").value = "com";
  $("ordenacao").value = "codigoAsc";
  aplicarFiltros();
};

$("anterior").onclick = () => { pagina--; render(); };
$("proxima").onclick = () => { pagina++; render(); };
$("menuBtn").onclick = () => $("menu").classList.toggle("ativo");
document.querySelectorAll("#menu button").forEach(b => b.onclick = () => trocarCatalogo(b.dataset.catalogo));
$("cotacaoFloat").onclick = () => $("cotacaoPainel").classList.add("ativo");
$("fecharCotacao").onclick = () => $("cotacaoPainel").classList.remove("ativo");