// ====== PRODUTOS: [código, nome, validade, estoque mínimo na prateleira] ======
// Para mudar o estoque mínimo ou incluir produtos, edite esta lista.
const PRODUTOS = [
  ["8035","Trufas sortidas (6 UNIDADES)","7 dias",10],
  ["8033","Brigadeiro de colher","5 dias",10],
  ["8124","Marshmallow","7 dias",15],
  ["8034","Torta banoffee","3 dias",10],
  ["7983","Bolo de coco 1Kg","5 dias",8],
  ["7982","Bolo de pistache 1kg","5 dias",8],
  ["7984","Bolo Prestígio 1kg","5 dias",8],
  ["8114","Brigadeiro de pistache (UNIDADE)","7 dias",10],
  ["8115","Doce de leite","7 dias",10],
  ["8116","Energyball (UNIDADE)","7 dias",10],
  ["8126","Prestígio (UNIDADE)","7 dias",10],
  ["8138","Trufa de brigadeiro (UNIDADE)","7 dias",10],
  ["7985","Cookie proteico","7 dias",30],
  ["8306","Cookie proteico de chocolate","7 dias",30],
  ["8304","Bolo de milho","5 dias",0],
  ["8303","Bolo de cenoura","7 dias",12],
  ["8445","Bolo de laranja","3 dias",0],
  ["8135","Pão de mel (unidade)","7 dias",30],
  ["8305","Brownie","6 dias",0],
  ["8456","Crostini Parmesão Tats","7 dias",12],
  ["8457","Cookie Proteico de Coco Tats","7 dias",30],
  ["8458","TATS Coin Preto 100g","2 meses",5],
  ["8459","TATS Coin Branco 100g","2 meses",5],
  ["8460","TATS Coin Preto 200g","2 meses",5],
  ["8462","TATS Coin Branco 200g","2 meses",5],
  ["8527","ALFAJOR frutas vermelhas","7 dias",30],
  ["8528","Tiramisu Vegano","7 dias",10],
  ["8530","Pudim","15 dias",15],
  ["8531","Panna Cotta Vegana","15 dias",15],
  ["8537","Tapioca Cremosa","15 dias",15],
  ["8536","Musse de Maracujá","15 dias",15],
  ["8535","Bolo de Chocolate c Brigadeiro pq","15 dias",10],
  ["8534","Tartelette de Limão","15 dias",20],
  ["8533","Bolo de Churros pq","15 dias",10],
  ["8532","Bolo Matilda kg","15 dias",8]
];
const POR_CODIGO = Object.fromEntries(PRODUTOS.map(p => [p[0], p]));
const KEY = "tats_reposicoes_v1";

// ====== DADOS ======
let db = {};
try { db = JSON.parse(localStorage.getItem(KEY)) || {}; } catch (e) { db = {}; }
const salvar = () => {
  try { localStorage.setItem(KEY, JSON.stringify(db)); }
  catch (e) { aviso("Não foi possível salvar neste navegador.", true); }
};

const $ = id => document.getElementById(id);
const hoje = () => new Date().toLocaleDateString("en-CA"); // aaaa-mm-dd, horário local
const fmtData = d => d.split("-").reverse().join("/");
const num = v => Math.max(0, parseInt(v, 10) || 0);

function aviso(t, erro) { const m = $("msg"); m.textContent = t; m.className = erro ? "err" : ""; }

// Vendido = (estoque anterior + reposição anterior) − desperdício − estoque atual
function vendido(codigo, data) {
  const anteriores = Object.keys(db).filter(d => d < data && db[d][codigo]).sort();
  if (!anteriores.length) return null;
  const a = db[anteriores.pop()][codigo], r = db[data][codigo];
  return Math.max(0, a.estoque + a.repo - r.desp - r.estoque);
}

// ====== FORMULÁRIO ======
function iniciar() {
  $("data").value = hoje();
  $("de").value = $("ate").value = hoje();
  $("produto").innerHTML = PRODUTOS.map(p => `<option value="${p[0]}">${p[1]}</option>`).join("");
  ["data", "produto"].forEach(id => $(id).addEventListener("change", aoMudar));
  ["de", "ate"].forEach(id => $(id).addEventListener("change", relatorio));
  $("add").onclick = adicionar;
  $("xlsx").onclick = exportarExcel;
  $("pdf").onclick = gerarPDF;
  $("exp").onclick = baixarBackup;
  $("imp").onchange = restaurarBackup;
  aoMudar();
}

function aoMudar() {
  const cod = $("produto").value, data = $("data").value;
  $("codigo").value = cod;
  $("minimo").value = POR_CODIGO[cod][3];
  const r = db[data] && db[data][cod];
  $("estoque").value = r ? r.estoque : "";
  $("reposicao").value = r ? r.repo : "";
  $("desperdicio").value = r ? r.desp : "";
  aviso(r ? "Já existe registro deste produto nesta data. Salvar vai substituir." : "");
  tabelaDia(); producao(); relatorio();
}

function adicionar() {
  const data = $("data").value, cod = $("produto").value;
  if (!data) return aviso("Escolha uma data.", true);
  if ($("estoque").value === "" || $("reposicao").value === "")
    return aviso("Preencha o estoque atual e a reposição (use 0 se não houver).", true);
  (db[data] = db[data] || {})[cod] = {
    estoque: num($("estoque").value), repo: num($("reposicao").value), desp: num($("desperdicio").value)
  };
  salvar();
  aviso(`Reposição de "${POR_CODIGO[cod][1]}" salva.`);
  $("estoque").value = $("reposicao").value = $("desperdicio").value = "";
  const i = $("produto").selectedIndex;
  if (i < PRODUTOS.length - 1) { $("produto").selectedIndex = i + 1; $("codigo").value = $("produto").value; $("minimo").value = POR_CODIGO[$("produto").value][3]; }
  tabelaDia(); producao(); relatorio();
}

function remover(data, cod) {
  delete db[data][cod];
  if (!Object.keys(db[data]).length) delete db[data];
  salvar(); aoMudar();
}

// ====== TABELAS ======
const tabela = (el, cab, linhas, vazio) => {
  el.innerHTML = linhas.length
    ? `<thead><tr>${cab.map((c, i) => `<th class="${c[1] ? "n" : ""}">${c[0]}</th>`).join("")}</tr></thead><tbody>${linhas.join("")}</tbody>`
    : `<tbody><tr><td class="empty">${vazio}</td></tr></tbody>`;
};

function tabelaDia() {
  const data = $("data").value, reg = db[data] || {};
  const linhas = PRODUTOS.filter(p => reg[p[0]]).map(p => {
    const r = reg[p[0]], v = vendido(p[0], data);
    return `<tr><td>${p[0]}</td><td>${p[1]}</td><td class="n">${r.estoque}</td><td class="n">${r.repo}</td><td class="n">${r.desp}</td><td class="n">${v === null ? "–" : v}</td><td><button class="del" data-c="${p[0]}">Excluir</button></td></tr>`;
  });
  tabela($("tabDia"), [["Cód."], ["Produto"], ["Estoque", 1], ["Reposição", 1], ["Desperdício", 1], ["Vendido", 1], [""]], linhas, "Nenhuma reposição registrada nesta data.");
  document.querySelectorAll(".del").forEach(b => b.onclick = () => remover(data, b.dataset.c));
}

function listaProducao() {
  const data = $("data").value, reg = db[data] || {};
  return PRODUTOS.filter(p => reg[p[0]]).map(p => {
    const r = reg[p[0]];
    return { cod: p[0], nome: p[1], estoque: r.estoque, repo: r.repo, min: p[3], produzir: Math.max(0, p[3] - (r.estoque + r.repo)) };
  });
}

function producao() {
  const l = listaProducao();
  tabela($("tabProd"), [["Cód."], ["Produto"], ["Na prateleira", 1], ["Mínimo", 1], ["Produzir", 1]],
    l.map(x => `<tr><td>${x.cod}</td><td>${x.nome}</td><td class="n">${x.estoque + x.repo}</td><td class="n">${x.min}</td><td class="n ${x.produzir ? "prod" : ""}">${x.produzir}</td></tr>`),
    "Registre reposições para ver o que produzir.");
}

// Dados do período
function dadosPeriodo() {
  const de = $("de").value, ate = $("ate").value;
  const datas = Object.keys(db).filter(d => d >= de && d <= ate).sort();
  const resumo = {}, detalhe = [];
  datas.forEach(d => PRODUTOS.forEach(p => {
    const r = db[d][p[0]]; if (!r) return;
    const v = vendido(p[0], d);
    const s = resumo[p[0]] = resumo[p[0]] || { repo: 0, desp: 0, vend: 0 };
    s.repo += r.repo; s.desp += r.desp; s.vend += v || 0;
    detalhe.push([fmtData(d), p[0], p[1], r.estoque, r.repo, r.desp, v === null ? "" : v]);
  }));
  return { resumo, detalhe };
}

function relatorio() {
  const { resumo } = dadosPeriodo();
  const linhas = PRODUTOS.filter(p => resumo[p[0]]).map(p => {
    const s = resumo[p[0]];
    return `<tr><td>${p[0]}</td><td>${p[1]}</td><td class="n">${s.repo}</td><td class="n">${s.desp}</td><td class="n">${s.vend}</td></tr>`;
  });
  tabela($("tabRel"), [["Cód."], ["Produto"], ["Reposto", 1], ["Desperdício", 1], ["Vendido", 1]], linhas, "Sem reposições neste período.");
}

// ====== EXPORTAÇÕES ======
function exportarExcel() {
  const { resumo, detalhe } = dadosPeriodo();
  if (!detalhe.length) return alert("Não há reposições no período escolhido.");
  const r = [["Código", "Produto", "Reposto", "Desperdício", "Vendido"]];
  PRODUTOS.filter(p => resumo[p[0]]).forEach(p => r.push([Number(p[0]), p[1], resumo[p[0]].repo, resumo[p[0]].desp, resumo[p[0]].vend]));
  const d = [["Data", "Código", "Produto", "Estoque atual", "Reposição", "Desperdício", "Vendido"], ...detalhe.map(x => [x[0], Number(x[1]), ...x.slice(2)])];
  const wb = XLSX.utils.book_new();
  const s1 = XLSX.utils.aoa_to_sheet(r), s2 = XLSX.utils.aoa_to_sheet(d);
  s1["!cols"] = [{ wch: 9 }, { wch: 36 }, { wch: 10 }, { wch: 12 }, { wch: 10 }];
  s2["!cols"] = [{ wch: 12 }, { wch: 9 }, { wch: 36 }, { wch: 13 }, { wch: 11 }, { wch: 12 }, { wch: 10 }];
  XLSX.utils.book_append_sheet(wb, s1, "Resumo do período");
  XLSX.utils.book_append_sheet(wb, s2, "Detalhe por dia");
  XLSX.writeFile(wb, `reposicao_${$("de").value}_a_${$("ate").value}.xlsx`);
}

function gerarPDF() {
  const l = listaProducao(), data = $("data").value;
  if (!l.length) return alert("Registre reposições nesta data antes de gerar o PDF.");
  const { jsPDF } = window.jspdf, doc = new jsPDF();
  let y = 20;
  doc.setFontSize(16); doc.text("TATS Confeitaria - O que produzir", 14, y);
  doc.setFontSize(10); doc.text("Data da reposicao: " + fmtData(data), 14, y += 7);
  const cab = () => {
    y += 9; doc.setFont(undefined, "bold");
    doc.text("Cod.", 14, y); doc.text("Produto", 30, y); doc.text("Prateleira", 118, y); doc.text("Minimo", 148, y); doc.text("Produzir", 175, y);
    doc.setFont(undefined, "normal"); doc.line(14, y + 2, 196, y + 2);
  };
  cab();
  const prod = l.filter(x => x.produzir > 0);
  prod.forEach(x => {
    if (y > 275) { doc.addPage(); y = 10; cab(); }
    y += 7;
    doc.text(x.cod, 14, y); doc.text(x.nome.slice(0, 48), 30, y);
    doc.text(String(x.estoque + x.repo), 125, y); doc.text(String(x.min), 153, y);
    doc.setFont(undefined, "bold"); doc.text(String(x.produzir), 181, y); doc.setFont(undefined, "normal");
  });
  y += 10;
  if (!prod.length) doc.text("Nada a produzir: todos os produtos estao no estoque minimo.", 14, y);
  else doc.text("Total a produzir: " + prod.reduce((a, x) => a + x.produzir, 0) + " unidades", 14, y);
  doc.save(`producao_${data}.pdf`);
}

// ====== BACKUP ======
function baixarBackup() {
  const a = document.createElement("a");
  a.href = URL.createObjectURL(new Blob([JSON.stringify(db)], { type: "application/json" }));
  a.download = `backup_tats_${hoje()}.json`; a.click();
}
function restaurarBackup(e) {
  const f = e.target.files[0]; if (!f) return;
  const rd = new FileReader();
  rd.onload = () => {
    try {
      const novo = JSON.parse(rd.result);
      if (typeof novo !== "object" || Array.isArray(novo)) throw 0;
      if (!confirm("Isso substitui todos os dados atuais. Continuar?")) return;
      db = novo; salvar(); aoMudar();
    } catch (x) { alert("Arquivo de backup inválido."); }
  };
  rd.readAsText(f);
}

iniciar();
