/**
 * Simula SpreadsheetApp, Session, PropertiesService etc. no navegador para
 * rodar o Code.gs de verdade na prévia local. Os dados ficam no localStorage.
 * NÃO vai para o Apps Script.
 */
(function () {
  'use strict';

  const CHAVE = 'mock-compras-v2';
  const USUARIOS_DEMO = [
    ['aprovador@amais.com.br', 'Ana (aprovadora)', 'APROVADOR', 'SIM'],
    ['comprador1@amais.com.br', 'Bruno', 'COMPRADOR', 'SIM'],
    ['comprador2@amais.com.br', 'Carla', 'COMPRADOR', 'SIM'],
  ];
  const ISO = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(\.\d+)?Z$/;

  let db = carregar();
  let emailAtual = localStorage.getItem(CHAVE + '-email') || USUARIOS_DEMO[0][0];

  function carregar() {
    try {
      const bruto = JSON.parse(localStorage.getItem(CHAVE));
      if (!bruto) return { abas: {}, ordem: [], props: {}, userProps: {} };
      Object.keys(bruto.abas).forEach(n => {
        bruto.abas[n] = bruto.abas[n].map(l => l.map(v => (typeof v === 'string' && ISO.test(v) ? new Date(v) : v)));
      });
      return bruto;
    } catch (e) {
      return { abas: {}, ordem: [], props: {}, userProps: {} };
    }
  }
  function salvar() { localStorage.setItem(CHAVE, JSON.stringify(db)); }

  const noop = function () { return this; };
  const celula = v => (typeof v === 'string' && v.charAt(0) === "'" ? v.slice(1) : v);

  function Aba(nome) { this.nome = nome; }
  Aba.prototype = {
    get dados() { return db.abas[this.nome]; },
    getName() { return this.nome; },
    getLastRow() {
      const d = this.dados;
      for (let i = d.length - 1; i >= 0; i--) if (d[i] && d[i].some(v => v !== '' && v !== null && v !== undefined)) return i + 1;
      return 0;
    },
    getLastColumn() {
      return this.dados.reduce((m, l) => {
        let u = 0;
        (l || []).forEach((v, j) => { if (v !== '' && v !== null && v !== undefined) u = j + 1; });
        return Math.max(m, u);
      }, 0);
    },
    getRange(a, b, c, d) {
      if (typeof a === 'string') return { setDataValidation: noop, setFontWeight: noop, setBackground: noop };
      return new Intervalo(this, a, b, c || 1, d || 1);
    },
    getDataRange() { return new Intervalo(this, 1, 1, Math.max(this.getLastRow(), 1), Math.max(this.getLastColumn(), 1)); },
    appendRow(valores) { this.dados[this.getLastRow()] = valores.map(celula); },
    deleteRows(inicio, quantidade) { this.dados.splice(inicio - 1, quantidade); },
    setFrozenRows: noop,
  };

  function Intervalo(aba, r, c, nr, nc) { Object.assign(this, { aba, r, c, nr, nc }); }
  Intervalo.prototype = {
    getValues() {
      const out = [];
      for (let i = 0; i < this.nr; i++) {
        const l = this.aba.dados[this.r - 1 + i] || [];
        const linha = [];
        for (let j = 0; j < this.nc; j++) {
          const v = l[this.c - 1 + j];
          linha.push(v === undefined || v === null ? '' : (v instanceof Date ? new Date(v) : v));
        }
        out.push(linha);
      }
      return out;
    },
    setValues(vals) {
      vals.forEach((l, i) => {
        const linha = this.aba.dados[this.r - 1 + i] = this.aba.dados[this.r - 1 + i] || [];
        l.forEach((v, j) => { linha[this.c - 1 + j] = celula(v); });
      });
      return this;
    },
    setFontWeight: noop,
    setBackground: noop,
    setDataValidation: noop,
  };

  const planilha = {
    getId: () => 'planilha-local',
    getSheetByName: n => (db.abas[n] ? new Aba(n) : null),
    insertSheet(n) { db.abas[n] = []; db.ordem.push(n); return new Aba(n); },
    getSheets: () => db.ordem.map(n => new Aba(n)),
    deleteSheet(a) { delete db.abas[a.nome]; db.ordem = db.ordem.filter(n => n !== a.nome); },
  };

  const validacao = { requireValueInList: noop, requireCheckbox: noop, build: () => ({}) };

  window.SpreadsheetApp = {
    getActive: () => planilha,
    openById: () => planilha,
    flush() {},
    newDataValidation: () => validacao,
    getUi: () => ({ alert() {}, createMenu() { return { addItem() { return this; }, addToUi() {} }; } }),
  };
  window.Session = { getActiveUser: () => ({ getEmail: () => emailAtual }) };

  function props(obj) {
    return { getProperty: k => (k in obj ? obj[k] : null), setProperty(k, v) { obj[k] = String(v); return this; } };
  }
  window.PropertiesService = {
    getScriptProperties: () => props(db.props),
    getUserProperties: () => props(db.userProps[emailAtual] = db.userProps[emailAtual] || {}),
  };
  window.LockService = { getScriptLock: () => ({ tryLock: () => true, releaseLock() {} }) };
  window.Utilities = { getUuid: () => crypto.randomUUID() };

  // ---- Dados de exemplo (só na primeira vez) ----
  function comoUsuario(email, fn) {
    const antes = emailAtual;
    emailAtual = email;
    try { return fn(); } finally { emailAtual = antes; }
  }

  function semear() {
    if (db.ordem.length) return;
    comoUsuario(USUARIOS_DEMO[0][0], () => window.configurarPlanilha());
    const usuarios = planilha.getSheetByName('Usuarios');
    usuarios.getRange(2, 1, USUARIOS_DEMO.length, 4).setValues(USUARIOS_DEMO);

    const [ana, bruno, carla] = USUARIOS_DEMO.map(u => u[0]);
    const cot = (fornecedor, canal, valorUnitario, frete, formaPagamento, condicaoPagamento, prazoEntrega) =>
      ({ fornecedor, canal, valorUnitario, frete, formaPagamento, condicaoPagamento, prazoEntrega });
    comoUsuario(bruno, () => {
      const dc = 'Datacenter Hospital Santa Clara';
      window.criarProjeto(dc, 'Hospital Santa Clara');
      window.criarProjeto('Cabeamento Escritório Centro', 'Construtora Alfa');
      window.criarProjeto('Interno / Manutenção', '');
      const [cabo, switchPoe, patch, nobreak] = window.criarItens(dc, [
        { descricao: 'Cabo de rede Cat6 azul (caixa 305 m)', quantidade: 4, unidade: 'cx', urgencia: 'ALTA', observacao: 'Certificado Anatel.' },
        { descricao: 'Switch PoE 24 portas gerenciável', quantidade: 2, unidade: 'un' },
        { descricao: 'Patch panel 24p Cat6', quantidade: 3, unidade: 'un' },
        { descricao: 'Nobreak 3 kVA senoidal', quantidade: 1, unidade: 'un' },
      ]).retorno;
      window.salvarCotacao(cabo, cot('Furukawa Distribuidor SP', 'Representante / vendedor', 1189.9, 0, 'Boleto', '28 dias', '3 dias úteis'));
      window.salvarCotacao(cabo, cot('Mercado Livre — Loja Oficial', 'Site / loja online', 1099, 85, 'Cartão de crédito', '3x sem juros', '7 dias'));
      window.salvarCotacao(switchPoe, cot('Intelbras Revenda', 'WhatsApp', 2349.9, 45, 'PIX', 'à vista', '2 dias'));
      window.salvarCotacao(switchPoe, cot('Amazon', 'Site / loja online', 2199, 0, 'Cartão de crédito', '10x', '5 dias'));
      window.salvarCotacao(patch, cot('Kalunga', 'Loja física', 189.9, 0, 'PIX', 'à vista', 'Retirada'));
      window.enviarParaAprovacao([cabo, switchPoe, patch]);
      window.comentar(cabo, 'Ana, o do ML é mais barato mas demora 7 dias. O pessoal precisa até sexta.');
      window.salvarCotacao(nobreak, cot('SMS Tecnologia', 'E-mail', 4890, 120, 'Faturado', '30/60 dias', '10 dias'));
    });
    comoUsuario(carla, () => {
      const [rack, alicate] = window.criarItens('Cabeamento Escritório Centro', [
        { descricao: 'Rack 19" 12U de parede', quantidade: 2, unidade: 'un' },
        { descricao: 'Alicate de crimpagem RJ45', quantidade: 3, unidade: 'un' },
      ]).retorno;
      window.salvarCotacao(rack, cot('Kalunga', 'Loja física', 689.5, 0, 'PIX', 'À vista (5% desc.)', 'Retirada imediata'));
      window.salvarCotacao(alicate, cot('Leroy Merlin', 'Site / loja online', 79.9, 19.9, 'Cartão de crédito', 'à vista', '4 dias'));
      window.enviarParaAprovacao([rack]);
      comoUsuario(ana, () => window.aprovar([{ itemId: rack, cotacaoId: window.getDados().cotacoes.find(c => c.ItemID === rack).ID }], 'Pode comprar, retira hoje.'));
      window.marcarComprado(rack, 'NF 4521');
    });
    salvar();
  }

  // ---- google.script.run ----
  function criarRunner(ok, falha) {
    return new Proxy({}, {
      get(_, nome) {
        if (nome === 'withSuccessHandler') return fn => criarRunner(fn, falha);
        if (nome === 'withFailureHandler') return fn => criarRunner(ok, fn);
        return function () {
          const args = JSON.parse(JSON.stringify(Array.prototype.slice.call(arguments)));
          setTimeout(() => {
            try {
              semear();
              if (typeof window[nome] !== 'function') throw new Error('Função ' + nome + ' não existe no Code.gs');
              const r = window[nome].apply(null, args);
              salvar();
              if (ok) ok(r === undefined ? r : JSON.parse(JSON.stringify(r)));
            } catch (e) {
              console.error('[mock]', nome, e);
              if (falha) falha(e);
            }
          }, 300);
        };
      },
    });
  }
  window.google = { script: { run: criarRunner(null, null) } };

  // ---- Barra para trocar de usuário ----
  document.addEventListener('DOMContentLoaded', () => {
    const barra = document.createElement('div');
    barra.style.cssText = 'position:fixed;left:12px;bottom:12px;z-index:99;background:#fff8d6;border:1px solid #e6cf6b;border-radius:8px;padding:6px 10px;font:12px system-ui;box-shadow:0 2px 8px rgba(0,0,0,.15);display:flex;gap:6px;align-items:center';
    barra.innerHTML = '<b>Prévia local</b> · entrar como <select id="mock-user">' +
      USUARIOS_DEMO.concat([['intruso@gmail.com', 'Sem acesso']]).map(u => '<option value="' + u[0] + '"' + (u[0] === emailAtual ? ' selected' : '') + '>' + u[1] + '</option>').join('') +
      '</select><button id="mock-reset" type="button">Resetar dados</button>';
    document.body.appendChild(barra);
    document.getElementById('mock-user').addEventListener('change', e => {
      localStorage.setItem(CHAVE + '-email', e.target.value);
      location.reload();
    });
    document.getElementById('mock-reset').addEventListener('click', () => {
      localStorage.removeItem(CHAVE);
      location.reload();
    });
  });
})();
