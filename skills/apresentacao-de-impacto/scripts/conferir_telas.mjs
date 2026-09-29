#!/usr/bin/env node
/*
  conferir_telas.mjs — confere se uma apresentação de palco fixo cabe na tela.

  Uso:
    node conferir_telas.mjs caminho/para/apresentacao.html [--fotos pasta] [--telas 1,5,9]

  O que faz, por tela:
    1. abre num navegador sem interface, no palco de 1600x900;
    2. mede cada elemento da tela ativa e aponta o que passa da área útil
       (112px no topo, 100px no rodapé) ou vaza pelas laterais;
    3. coleta erros de JavaScript;
    4. salva uma foto para você olhar depois.

  Sai com código 1 se achar estouro ou erro — dá para usar em automação.
  Requer Node 18+ e Chrome ou Edge instalados. Defina NAVEGADOR=/caminho/do/chrome
  se o script não encontrar sozinho.
*/

import { execFileSync } from "node:child_process";
import { existsSync, mkdirSync, readFileSync, writeFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve, basename } from "node:path";
import { pathToFileURL } from "node:url";

/* ---------- argumentos ---------- */
const args = process.argv.slice(2);
const alvo = args.find(a => !a.startsWith("--"));
if (!alvo) {
  console.error("Uso: node conferir_telas.mjs apresentacao.html [--fotos pasta] [--telas 1,5]");
  process.exit(2);
}
const arquivo = resolve(alvo);
if (!existsSync(arquivo)) { console.error("Arquivo não encontrado: " + arquivo); process.exit(2); }

const opt = (nome, padrao) => {
  const k = args.indexOf("--" + nome);
  return k >= 0 && args[k + 1] ? args[k + 1] : padrao;
};
const pastaFotos = resolve(opt("fotos", join(tmpdir(), "conferir-telas", basename(arquivo, ".html"))));
const telasPedidas = opt("telas", null);

/* ---------- navegador ---------- */
function acharNavegador() {
  if (process.env.NAVEGADOR) return process.env.NAVEGADOR;
  const candidatos = [
    "C:/Program Files/Google/Chrome/Application/chrome.exe",
    "C:/Program Files (x86)/Google/Chrome/Application/chrome.exe",
    "C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe",
    "C:/Program Files/Microsoft/Edge/Application/msedge.exe",
    "/Applications/Google Chrome.app/Contents/MacOS/Google Chrome",
    "/Applications/Microsoft Edge.app/Contents/MacOS/Microsoft Edge",
    "/usr/bin/google-chrome", "/usr/bin/chromium", "/usr/bin/chromium-browser",
    "/snap/bin/chromium"
  ];
  for (const c of candidatos) if (existsSync(c)) return c;
  return null;
}
const navegador = acharNavegador();
if (!navegador) {
  console.error("Não achei Chrome nem Edge. Defina NAVEGADOR=/caminho/do/chrome e rode de novo.");
  process.exit(2);
}

/* ---------- sonda injetada na cópia ----------
   O Chrome não executa código nosso pela linha de comando, então a sonda
   se auto-dispara depois do carregamento e publica o resultado num atributo
   do <html>, que lemos no --dump-dom.                                      */
const SONDA = `
<script>
window.__erros = [];
window.addEventListener("error", function(e){
  __erros.push(String(e.message || e.error) + " (linha " + (e.lineno || "?") + ")");
});
/* Medir com a animacao de entrada no ar acusaria os 18px do translateY como
   estouro. Antes de medir, levamos tudo ao estado final.                    */
function __congelar(){
  var s = document.createElement("style");
  s.textContent = ".anim{opacity:1!important;transform:none!important;transition:none!important}"
                + "*,*::before,*::after{animation:none!important;transition:none!important}";
  document.head.appendChild(s);
  void document.body.offsetHeight;   /* forca o recalculo antes da medicao */
}
function __auditar(){
  __congelar();
  var palco = document.getElementById("stage") || document.getElementById("palco")
           || document.querySelector("[data-palco]");
  if(!palco) return { erro: "palco nao encontrado (esperava #stage)" };
  var pr = palco.getBoundingClientRect();
  var esc = pr.width / 1600 || 1;
  var tela = document.querySelector(".slide.is-active, .tela.on, section.is-active, section.on");
  /* Tolerancias generosas de proposito: diferenca de 1 a 3px e arredondamento de
     linha, nao defeito. Estouro que importa passa disso com folga, e um
     verificador que grita a toa deixa de ser lido.                            */
  var TOPO = 112, BASE = 800, DIR = 1512, ESQ = 88, TOL = 4, TOL_DENTRO = 8;

  /* invisivel ou recortado por um ancestral (ex.: lista recolhida do acordeao
     de fases) nao e estouro — o cliente nunca ve aquilo                        */
  function __escondido(n){
    var a = n;
    while (a && a !== document.body) {
      var cs = getComputedStyle(a);
      if (cs.visibility === "hidden" || cs.display === "none" || parseFloat(cs.opacity) === 0) return true;
      if (a !== n && cs.overflowY !== "visible") {
        var ar = a.getBoundingClientRect(), nr = n.getBoundingClientRect();
        if (nr.top >= ar.bottom - 1 || nr.bottom <= ar.top + 1) return true;
      }
      a = a.parentElement;
    }
    return false;
  }

  /* nome curto de um elemento, para o relatorio */
  function __curto(n){
    var s = n.tagName.toLowerCase();
    if (n.id) s += "#" + n.id;
    if (n.className && typeof n.className === "string")
      s += "." + n.className.trim().split(/\\s+/).slice(0,2).join(".");
    return s;
  }

  /* Texto que passa da borda do proprio SVG fica cortado sem aviso nenhum:
     o rotulo de ponta de uma barra e o caso mais comum. Comparamos a caixa
     de cada <text> com o viewBox, nas coordenadas do desenho.              */
  function __textoCortado(tela){
    var achados = [], svgs = tela.querySelectorAll("svg");
    for (var a = 0; a < svgs.length; a++) {
      var svg = svgs[a];
      if (!svg.viewBox || !svg.viewBox.baseVal || !svg.viewBox.baseVal.width) continue;
      if (svg.classList.contains("ico")) continue;
      var vb = svg.viewBox.baseVal, textos = svg.querySelectorAll("text"), TX = 2;
      for (var b = 0; b < textos.length; b++) {
        var cx; try { cx = textos[b].getBBox(); } catch (e) { continue; }
        if (!cx || cx.width < 1) continue;
        var sai = [];
        if (cx.x + cx.width > vb.x + vb.width + TX) sai.push("direita");
        if (cx.x < vb.x - TX) sai.push("esquerda");
        if (cx.y + cx.height > vb.y + vb.height + TX) sai.push("baixo");
        if (cx.y < vb.y - TX) sai.push("cima");
        if (sai.length) achados.push({
          onde: (svg.id ? "#" + svg.id : "svg") + " \\u201c" + (textos[b].textContent || "").slice(0, 30) + "\\u201d",
          o_que: "rotulo cortado pela borda do grafico (" + sai.join(" e ") + ") — aumente a margem desse lado"
        });
      }
    }
    return achados;
  }

  var achados = [], nos = [];
  if (tela) {
    var todos = tela.querySelectorAll("*");
    for (var k = 0; k < todos.length; k++) {
      var n = todos[k];
      if (n.ownerSVGElement || n.tagName.toLowerCase() === "svg") continue;  /* dentro de SVG nao se aplica */
      var r = n.getBoundingClientRect();
      if (r.width < 1 || r.height < 1) continue;
      if (__escondido(n)) continue;
      var cima = (r.top - pr.top) / esc, baixo = (r.bottom - pr.top) / esc;
      var esquerda = (r.left - pr.left) / esc, direita = (r.right - pr.left) / esc;
      var falhas = [];
      /* 1. o elemento vaza para fora da area util da tela */
      if (baixo > BASE + TOL)   falhas.push("passa " + Math.round(baixo - BASE) + "px do rodape");
      if (cima < TOPO - TOL)    falhas.push("passa " + Math.round(TOPO - cima) + "px do topo");
      if (direita > DIR + TOL)  falhas.push("passa " + Math.round(direita - DIR) + "px da direita");
      if (esquerda < ESQ - TOL) falhas.push("passa " + Math.round(ESQ - esquerda) + "px da esquerda");
      /* 2. o conteudo vaza para fora do proprio quadro — acontece quando o flex
            encolhe um cartao e o texto escapa por baixo, sem sair da tela      */
      var cs = getComputedStyle(n);
      if (cs.overflowY === "visible" && cs.position !== "absolute" && cs.position !== "fixed") {
        /* so vale para conteudo em fluxo: trilho, agulha e escala do simulador
           sao filhos absolutos que saem do container de proposito             */
        var temFluxo = n.children.length === 0;
        for (var j = 0; j < n.children.length && !temFluxo; j++) {
          var pf = getComputedStyle(n.children[j]).position;
          if (pf !== "absolute" && pf !== "fixed") temFluxo = true;
        }
        var sobra = n.scrollHeight - n.clientHeight;
        if (temFluxo && n.clientHeight > 0 && sobra > TOL_DENTRO) {
          /* Nao basta apontar o filho mais alto: numa grade, TODOS os filhos sao
             esticados ate a altura da linha, entao o mais alto empata com os
             outros e o culpado se esconde. O que denuncia e a altura do
             CONTEUDO de cada filho — quem tem mais conteudo e quem manda.   */
          var medidas = [];
          for (var c = 0; c < n.children.length; c++) {
            var f = n.children[c], fs2 = getComputedStyle(f);
            var dentro = (parseFloat(fs2.paddingTop) || 0) + (parseFloat(fs2.paddingBottom) || 0)
                       + Math.max(0, f.children.length - 1) * (parseFloat(fs2.rowGap) || 0);
            for (var g = 0; g < f.children.length; g++) {
              var nf = f.children[g], ns = getComputedStyle(nf);
              if (ns.position === "absolute" || ns.position === "fixed") continue;
              dentro += nf.getBoundingClientRect().height
                      + (parseFloat(ns.marginTop) || 0) + (parseFloat(ns.marginBottom) || 0);
            }
            medidas.push({ no: f, dentro: Math.round(dentro / esc) });
          }
          medidas.sort(function(a, b){ return b.dentro - a.dentro; });
          var quem = medidas.slice(0, 3).map(function(m){ return __curto(m.no) + " ~" + m.dentro + "px"; });
          falhas.push("o conteudo passa " + Math.round(sobra) + "px do proprio quadro"
            + (quem.length ? " (conteudo dos filhos: " + quem.join(", ") + " — o primeiro e quem manda)" : ""));
        }
      }
      if (falhas.length) { nos.push(n); achados.push({ no: n, falhas: falhas }); }
    }
  }
  /* so o culpado mais externo: se o pai ja estourou, o filho e consequencia */
  var externos = achados.filter(function(a){
    for (var k = 0; k < nos.length; k++) if (nos[k] !== a.no && nos[k].contains(a.no)) return false;
    return true;
  });
  var descrever = function(n){
    var t = (n.textContent || "").replace(/\\s+/g, " ").trim().slice(0, 44);
    return __curto(n) + (t ? " \\u201c" + t + "\\u201d" : "");
  };
  var lista = externos.map(function(a){ return { onde: descrever(a.no), o_que: a.falhas.join("; ") }; });
  if (tela) lista = lista.concat(__textoCortado(tela));
  return {
    rotulo: tela ? (tela.getAttribute("data-label") || tela.getAttribute("data-rot") || "") : "(sem tela ativa)",
    total: document.querySelectorAll(".slide, .tela").length,
    estouros: lista,
    erros: window.__erros.slice(0, 5)
  };
}
window.addEventListener("load", function(){
  setTimeout(function(){
    var r;
    try { r = __auditar(); } catch(e) { r = { erro: String(e), estouros: [], erros: [String(e)] }; }
    document.documentElement.setAttribute("data-auditoria", encodeURIComponent(JSON.stringify(r)));
  }, 3000);
});
</script>`;

/* ---------- sonda 2: simula um iframe protegido ----------
   No preview do claude.ai, no Notion e em qualquer <iframe sandbox>, chamadas
   como history.replaceState e requestFullscreen LANCAM erro. Se uma delas
   estiver no meio da funcao que troca de tela, tudo que vem depois para de
   rodar -- inclusive o desenho dos graficos -- e a tela aparece em branco sem
   aviso nenhum. Aqui forcamos essas chamadas a falhar e navegamos algumas
   telas, para ver se a apresentacao aguenta.                                */
const SONDA_IFRAME = `
<script>
window.__erros = [];
window.addEventListener("error", function(e){
  __erros.push(String(e.message || e.error) + " (linha " + (e.lineno || "?") + ")");
});
(function(){
  var negar = function(nome){ return function(){ throw new Error("bloqueado em iframe protegido: " + nome); }; };
  try { history.replaceState = negar("history.replaceState"); } catch(e){}
  try { history.pushState    = negar("history.pushState"); } catch(e){}
  try { Element.prototype.requestFullscreen = negar("requestFullscreen"); } catch(e){}
  try { Document.prototype.exitFullscreen   = negar("exitFullscreen"); } catch(e){}
})();
window.addEventListener("load", function(){
  setTimeout(function(){
    var proximo = document.getElementById("next");
    for (var k = 0; k < 6; k++) {
      try {
        if (proximo) proximo.click();
        else document.dispatchEvent(new KeyboardEvent("keydown", { key: "ArrowRight" }));
      } catch(e) { __erros.push("ao trocar de tela: " + e.message); }
    }
    document.documentElement.setAttribute("data-iframe",
      encodeURIComponent(JSON.stringify({ erros: window.__erros.slice(0, 5) })));
  }, 2600);
});
</script>`;

/* cópia instrumentada, para não tocar no arquivo do usuário */
const tmpDir = join(tmpdir(), "conferir-telas-" + process.pid);
mkdirSync(tmpDir, { recursive: true });
const copia = join(tmpDir, basename(arquivo));
const html = readFileSync(arquivo, "utf8");
if (!/<head[^>]*>/i.test(html)) {
  console.error("Isto nao parece um HTML completo (nao achei <head>).");
  process.exit(2);
}
writeFileSync(copia, html.replace(/<head([^>]*)>/i, "<head$1>" + SONDA));
const url = pathToFileURL(copia).href;

/* segunda copia, com as chamadas de iframe forcadas a falhar */
const copiaIframe = join(tmpDir, "iframe-" + basename(arquivo));
writeFileSync(copiaIframe, html.replace(/<head([^>]*)>/i, "<head$1>" + SONDA_IFRAME));
const urlIframe = pathToFileURL(copiaIframe).href;

function auditarIframe() {
  let dom = "";
  try {
    dom = execFileSync(navegador, [
      "--headless=new", "--disable-gpu", "--no-sandbox", "--hide-scrollbars",
      "--window-size=1600,900", "--virtual-time-budget=9000",
      "--dump-dom", urlIframe + "#1"
    ], { encoding: "utf8", maxBuffer: 1 << 28, stdio: ["ignore", "pipe", "ignore"] });
  } catch { return { erros: ["o navegador nao abriu a pagina"] }; }
  const m = dom.match(/data-iframe="([^"]*)"/);
  if (!m) return { erros: [] };
  try { return JSON.parse(decodeURIComponent(m[1])); } catch { return { erros: [] }; }
}

/* ---------- roda uma tela ---------- */
function auditar(n) {
  const flags = [
    "--headless=new", "--disable-gpu", "--no-sandbox", "--hide-scrollbars",
    "--window-size=1600,900", "--virtual-time-budget=8000",
    "--screenshot=" + join(pastaFotos, "tela-" + String(n).padStart(2, "0") + ".png"),
    "--dump-dom", url + "#" + n
  ];
  let dom = "";
  try {
    dom = execFileSync(navegador, flags, {
      encoding: "utf8", maxBuffer: 1 << 28, stdio: ["ignore", "pipe", "ignore"]
    });
  } catch {
    return { erros: ["o navegador nao abriu a pagina"], estouros: [] };
  }
  const m = dom.match(/data-auditoria="([^"]*)"/);
  if (!m) return { erros: ["a sonda nao respondeu (pagina travou ou demorou demais)"], estouros: [] };
  try { return JSON.parse(decodeURIComponent(m[1])); }
  catch { return { erros: ["resposta da sonda ilegivel"], estouros: [] }; }
}

/* ---------- laço principal ---------- */
mkdirSync(pastaFotos, { recursive: true });
console.log("\nConferindo " + basename(arquivo));
console.log("Navegador: " + navegador);

const primeira = auditar(1);
const total = primeira.total || 1;
const lista = telasPedidas
  ? telasPedidas.split(",").map(s => parseInt(s.trim(), 10)).filter(Boolean)
  : Array.from({ length: total }, (_, k) => k + 1);

const relatorio = [];
for (const n of lista) {
  const r = (n === 1 && !telasPedidas) ? primeira : auditar(n);
  relatorio.push({ n, ...r });
  process.stdout.write(".");
}
console.log("\n");

/* ---------- relatório ---------- */
/* a prova de fogo do iframe protegido roda uma vez, no caminho comum */
const iframe = telasPedidas ? { erros: [] } : auditarIframe();

let problemas = 0;
for (const r of relatorio) {
  const nome = "Tela " + String(r.n).padStart(2, "0") + (r.rotulo ? " · " + r.rotulo : "");
  const estouros = r.estouros || [], erros = r.erros || [];
  if (!estouros.length && !erros.length) { console.log("  ok    " + nome); continue; }
  problemas++;
  console.log("  FALHA " + nome);
  for (const e of estouros) console.log("        - " + e.o_que + " -> " + e.onde);
  for (const e of erros)    console.log("        - erro de JavaScript: " + e);
}

if ((iframe.erros || []).length) {
  problemas++;
  console.log("  FALHA Dentro de um iframe protegido (preview do claude.ai, Notion)");
  for (const e of iframe.erros) console.log("        - " + e);
  console.log("        Isole a chamada num try/catch e deixe-a por ultimo na funcao:");
  console.log("        um erro ali aborta o desenho dos graficos sem aviso nenhum.");
} else if (!telasPedidas) {
  console.log("  ok    Sobrevive dentro de um iframe protegido");
}

console.log("\nFotos em: " + pastaFotos);
if (problemas) {
  console.log("\n" + problemas + " tela(s) com problema. Conserte antes de entregar.");
  console.log("E olhe as fotos: o script pega estouro e colisao, mas nao enxerga");
  console.log("buraco no meio da tela nem titulo quebrado em palavra feia.");
} else {
  console.log("\nNenhum estouro e nenhum erro de JavaScript. Agora olhe as fotos.");
}
rmSync(tmpDir, { recursive: true, force: true });
process.exit(problemas ? 1 : 0);
