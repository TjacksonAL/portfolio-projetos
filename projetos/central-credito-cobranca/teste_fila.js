// Testa a funcao real da pagina (extraida do central.html) contra um DOM de mentira.
const fs = require('fs');
const html = fs.readFileSync('central.html', 'utf8');
const js = html.match(/<script>([\s\S]*?)<\/script>/)[1];

// --- DOM de mentira, so com o que a funcao usa ---
class Classes {
  constructor(el) { this.el = el; this.set = new Set(el._classe.split(' ').filter(Boolean)); }
  toggle(c, on) { if (on) this.set.add(c); else this.set.delete(c); this.el._classe = [...this.set].join(' '); }
  contains(c) { return this.set.has(c); }
}
class El {
  constructor(tag, classe = '', dataset = {}) {
    this.tag = tag; this._classe = classe; this.dataset = dataset;
    this.filhos = []; this._texto = ''; this.innerHTML = '';
  }
  // o DOM de verdade converte para texto ao atribuir; o stub precisa fazer igual
  get textContent() { return this._texto; }
  set textContent(v) { this._texto = String(v); }
  get className() { return this._classe; }
  get classList() { return new Classes(this); }
  querySelector(sel) {
    const alvo = sel.replace('.', '');
    return this.filhos.find(f => f._classe.split(' ').includes(alvo)) || null;
  }
  querySelectorAll(sel) {
    const alvo = sel.split(' ').pop().replace('.', '');
    return this.filhos.filter(f => f._classe.split(' ').includes(alvo));
  }
  appendChild(f) { const i = this.filhos.indexOf(f); if (i >= 0) this.filhos.splice(i, 1); this.filhos.push(f); return f; }
  insertBefore(novo, ref) { const i = this.filhos.indexOf(ref); this.filhos.splice(i < 0 ? 0 : i, 0, novo); return novo; }
  remove() {}
}

// monta a fila de Hoje igual a que o Python gera: 7 linhas, ranks crescentes
// a data vem da propria pagina: com a base do mes seguinte, uma data fixa aqui
// deixaria os registros de teste velhos demais e o teste reprovaria sem defeito
const REF = html.match(/REF_ISO = '(\d{4}-\d{2}-\d{2})'/)[1];
function diasAntes(iso, n) {
  const d = new Date(iso + 'T12:00:00Z');
  d.setUTCDate(d.getUTCDate() - n);
  return d.toISOString().substring(0, 10);
}
const linhas = [
  { cod: 1, rank: 2.0000, recente: '0' },
  { cod: 5, rank: 2.6667, recente: '0' },
  { cod: 3, rank: 3.0000, recente: '0' },
  { cod: 11, rank: 3.6667, recente: '0' },
  { cod: 4, rank: 4.6667, recente: '0' },
  { cod: 6, rank: 5.0000, recente: '0' },
  { cod: 12, rank: 5.3333, recente: '1' },  // ja veio cobrado da planilha
];
const fila = new El('div', 'fila');
for (const l of linhas) {
  const linha = new El('div', 'linha', { cod: String(l.cod), rank: String(l.rank), recente: l.recente });
  const pos = new El('div', 'posicao');
  const exp = new El('p', 'explicacao');
  linha.filhos.push(pos, exp);
  fila.filhos.push(linha);
}
const telaHoje = new El('div', 'tela'); telaHoje.filhos.push(fila);

let armazem = {};
global.localStorage = {
  getItem: k => (k in armazem ? armazem[k] : null),
  setItem: (k, v) => { armazem[k] = v; },
};
global.alert = () => {};
global.document = {
  getElementById: id => (id === 'dados-regua' ? { textContent: '{}' }
                  : id === 'registros-arquivo' ? { textContent: '[]' } : null),
  querySelector: sel => (sel === '#tela-hoje .fila' ? fila : null),
  createElement: t => new El(t),
  addEventListener: () => {},
  body: new El('body'),
};
global.window = {};

// carrega o codigo real, tirando so as linhas que leem os dados embutidos na pagina
eval(js.replace(/var TEXTOS = .*?;\n/, 'var TEXTOS = {};\n')
       .replace(/var REG_ARQUIVO = .*?;\n/, 'var REG_ARQUIVO = [];\n'));

function ordemAtual() { return fila.filhos.map(f => Number(f.dataset.cod)); }
function posicoesAtuais() { return fila.filhos.map(f => f.querySelector('.posicao').textContent); }

let falhas = 0;
function conferir(nome, real, esperado) {
  const ok = JSON.stringify(real) === JSON.stringify(esperado);
  if (!ok) falhas++;
  console.log(`${ok ? '  [OK]  ' : '  [FALHOU]'} ${nome}`);
  if (!ok) console.log(`          esperado ${JSON.stringify(esperado)}\n          obtido   ${JSON.stringify(real)}`);
}

console.log('TESTE DA FILA DE HOJE (funcao real da pagina)\n');

// 1. sem nenhum registro: quem ja veio cobrado da planilha fica no fim
aplicarRegistrosNaFila();
conferir('sem registros, ordem preservada e cobrado da planilha no fim',
  ordemAtual(), [1, 5, 3, 11, 4, 6, 12]);
conferir('posicoes renumeradas de 1 a 7', posicoesAtuais(), ['1','2','3','4','5','6','7']);

// 2. registrando o primeiro da fila, ele tem que cair para o fim
armazem['aurora_cobrancas_registradas'] = JSON.stringify([{
  cod_cliente: 1, nome: 'Atacado Sao Jorge Ltda',
  dt_hora_contato: REF + 'T09:00:00', canal: 'WhatsApp',
  tom: 'Primeiro aviso cordial', valor_cobrado_centavos: 6013000, titulos_na_cobranca: 1,
}]);
aplicarRegistrosNaFila();
conferir('registrei o 1o da fila -> ele desce para o fim',
  ordemAtual(), [5, 3, 11, 4, 6, 1, 12]);

// 3. registro antigo (fora da janela de 48h) nao pode derrubar ninguem
armazem['aurora_cobrancas_registradas'] = JSON.stringify([{
  cod_cliente: 5, nome: 'Mercearia Dona Zica',
  dt_hora_contato: diasAntes(REF, 11) + 'T09:00:00', canal: 'Telefone',
  tom: 'Reaproximacao comercial', valor_cobrado_centavos: 1217000, titulos_na_cobranca: 2,
}]);
aplicarRegistrosNaFila();
conferir('registro de 11 dias atras nao derruba ninguem',
  ordemAtual(), [1, 5, 3, 11, 4, 6, 12]);

// 4. dois registros na janela: ambos descem, mantendo a ordem de prioridade entre si
armazem['aurora_cobrancas_registradas'] = JSON.stringify([
  { cod_cliente: 3, nome: 'Vila Nova', dt_hora_contato: REF + 'T09:00:00', canal: 'E-mail', tom: 'Conversa', valor_cobrado_centavos: 12678000, titulos_na_cobranca: 5 },
  { cod_cliente: 1, nome: 'Sao Jorge', dt_hora_contato: REF + 'T09:00:00', canal: 'WhatsApp', tom: 'Cordial', valor_cobrado_centavos: 6013000, titulos_na_cobranca: 1 },
]);
aplicarRegistrosNaFila();
conferir('dois registrados descem e mantem prioridade entre si',
  ordemAtual(), [5, 11, 4, 6, 1, 3, 12]);

// 5. o download junta arquivo e navegador: quem baixa de outra maquina nao apaga o que ja existia
const A = { cod_cliente: 3, nome: 'Vila Nova', dt_hora_contato: diasAntes(REF, 30) + 'T09:00:00',
            canal: 'WhatsApp', tom: 'Conversa', valor_cobrado_centavos: 12678000, titulos_na_cobranca: 5 };
const B = { cod_cliente: 5, nome: 'Dona Zica', dt_hora_contato: REF + 'T09:00:00',
            canal: 'Telefone', tom: 'Reaproximacao', valor_cobrado_centavos: 1217000, titulos_na_cobranca: 2 };
REG_ARQUIVO = [A];                                            // o que ja estava no CSV
armazem['aurora_cobrancas_registradas'] = JSON.stringify([B]); // outra maquina: so o que ela registrou
conferir('navegador de outra pessoa + 1 registro: o download leva os 2, nao apaga o do arquivo',
  todosRegistros().map(r => r.cod_cliente), [3, 5]);
armazem['aurora_cobrancas_registradas'] = JSON.stringify([A, B]);
conferir('registro que esta no arquivo e no navegador sai uma vez so',
  todosRegistros().map(r => r.cod_cliente), [3, 5]);
conferir('o que ja esta no arquivo nao e reaplicado na fila pelo navegador',
  registrosSoNoNavegador().map(r => r.cod_cliente), [5]);

// 6. reordenar a fila por outra coisa (aqui, valor vencido) nao tira do fim quem ja
//    foi cobrado, e o numero de cada cartao continua sendo a prioridade da Central
const valores = { 1: 6013000, 5: 1217000, 3: 12678000, 11: 2128000, 4: 3080000, 6: 3100000, 12: 2052000 };
fila.filhos.forEach(f => { f.dataset.valor = String(valores[f.dataset.cod]); });
REG_ARQUIVO = [];
armazem['aurora_cobrancas_registradas'] = '[]';
ordemHoje = 'valor';
aplicarRegistrosNaFila();
conferir('ordenar por valor: maior valor primeiro, quem ja foi cobrado continua no fim',
  ordemAtual(), [3, 1, 6, 4, 11, 5, 12]);
conferir('ordenar por valor: cada cartao mostra a posicao na prioridade da Central',
  posicoesAtuais(), ['3', '1', '6', '5', '4', '2', '7']);
ordemHoje = 'rank';
aplicarRegistrosNaFila();
conferir('voltar para a prioridade da Central devolve a ordem original',
  ordemAtual(), [1, 5, 3, 11, 4, 6, 12]);

console.log(`\n${falhas === 0 ? 'TODOS OS TESTES PASSARAM' : falhas + ' TESTE(S) FALHARAM'}`);
process.exit(falhas === 0 ? 0 : 1);
