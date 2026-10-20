/**
 * Sistema de Compras A+.com — backend (Google Apps Script).
 *
 * O banco de dados é a própria planilha onde este script está instalado
 * (Extensões > Apps Script). Rode o menu "Sistema de Compras > Configurar
 * planilha" uma vez para criar as abas.
 */

const PAPEL = { APROVADOR: 'APROVADOR', COMPRADOR: 'COMPRADOR' };

const STATUS = {
  EM_COTACAO: 'EM_COTACAO',
  AGUARDANDO: 'AGUARDANDO_APROVACAO',
  APROVADO: 'APROVADO',
  REPROVADO: 'REPROVADO',
  COMPRADO: 'COMPRADO',
  ENTREGUE: 'ENTREGUE',
  CANCELADO: 'CANCELADO',
};

const ROTULO_STATUS = {
  EM_COTACAO: 'Em cotação',
  AGUARDANDO_APROVACAO: 'Aguardando aprovação',
  APROVADO: 'Aprovado',
  REPROVADO: 'Reprovado',
  COMPRADO: 'Comprado',
  ENTREGUE: 'Entregue',
  CANCELADO: 'Cancelado',
};

const ABAS = {
  USUARIOS: { nome: 'Usuarios', colunas: ['Email', 'Nome', 'Papel', 'Ativo'] },
  PROJETOS: { nome: 'Projetos', colunas: ['ID', 'Nome', 'Cliente', 'Ativo', 'CriadoPor', 'CriadoEm'] },
  ITENS: {
    nome: 'Itens',
    colunas: [
      'ID', 'CriadoEm', 'CriadoPor', 'Projeto', 'Descricao', 'Quantidade', 'Unidade', 'Urgencia',
      'NecessarioAte', 'Observacao', 'Status', 'CotacaoAprovadaID', 'ValorAprovado', 'DecididoPor',
      'DecididoEm', 'MotivoDecisao', 'CompradoPor', 'CompradoEm', 'NumeroPedido', 'EntregueEm',
      'RecebidoPor', 'AtualizadoEm',
    ],
  },
  COTACOES: {
    nome: 'Cotacoes',
    colunas: [
      'ID', 'ItemID', 'Fornecedor', 'Canal', 'Link', 'ValorUnitario', 'Frete', 'ValorTotal',
      'FormaPagamento', 'CondicaoPagamento', 'PrazoEntrega', 'Observacao', 'CriadoPor', 'CriadoEm', 'Ativa',
    ],
  },
  COMENTARIOS: { nome: 'Comentarios', colunas: ['ID', 'ItemID', 'Autor', 'Texto', 'CriadoEm'] },
  HISTORICO: { nome: 'Historico', colunas: ['ItemID', 'Data', 'Usuario', 'Acao', 'Detalhe'] },
};

// ---------------------------------------------------------------------------
// Web app
// ---------------------------------------------------------------------------

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('Compras · A+.com')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(nome) {
  return HtmlService.createHtmlOutputFromFile(nome).getContent();
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Sistema de Compras')
    .addItem('Configurar planilha', 'configurarPlanilha')
    .addToUi();
}

/** Cria as abas que faltarem e cadastra quem rodou como primeiro aprovador. */
function configurarPlanilha() {
  const ss = SpreadsheetApp.getActive();
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', ss.getId());

  Object.keys(ABAS).forEach(chave => {
    const def = ABAS[chave];
    const sh = ss.getSheetByName(def.nome) || ss.insertSheet(def.nome);
    const ultimaColuna = Math.max(sh.getLastColumn(), 1);
    const cabecalho = sh.getRange(1, 1, 1, ultimaColuna).getValues()[0].map(String);
    const faltando = def.colunas.filter(c => cabecalho.indexOf(c) < 0);
    if (faltando.length) {
      const inicio = cabecalho.join('') === '' ? 1 : cabecalho.length + 1;
      sh.getRange(1, inicio, 1, faltando.length).setValues([faltando]);
    }
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, sh.getLastColumn()).setFontWeight('bold').setBackground('#e8f2d4');
  });

  const usuarios = ss.getSheetByName(ABAS.USUARIOS.nome);
  const validacaoPapel = SpreadsheetApp.newDataValidation()
    .requireValueInList([PAPEL.APROVADOR, PAPEL.COMPRADOR], true).build();
  const validacaoAtivo = SpreadsheetApp.newDataValidation().requireValueInList(['SIM', 'NAO'], true).build();
  usuarios.getRange('C2:C').setDataValidation(validacaoPapel);
  usuarios.getRange('D2:D').setDataValidation(validacaoAtivo);
  if (usuarios.getLastRow() < 2) {
    const email = Session.getActiveUser().getEmail();
    if (email) usuarios.appendRow([email, email.split('@')[0], PAPEL.APROVADOR, 'SIM']);
  }

  // Remove a aba padrão vazia que vem em toda planilha nova.
  ['Página1', 'Planilha1', 'Sheet1'].forEach(nome => {
    const sh = ss.getSheetByName(nome);
    if (sh && sh.getLastRow() === 0 && ss.getSheets().length > 1) ss.deleteSheet(sh);
  });

  try {
    SpreadsheetApp.getUi().alert(
      'Planilha configurada.\n\nCadastre as pessoas na aba "Usuarios" (Papel: APROVADOR ou COMPRADOR) ' +
      'e depois publique em Implantar > Nova implantação > App da Web.'
    );
  } catch (e) {
    // Rodando pelo editor: não há interface da planilha para o alerta.
  }
}

// ---------------------------------------------------------------------------
// Leitura
// ---------------------------------------------------------------------------

/** Primeira chamada do site: diz quem é o usuário ou que ele não tem acesso. */
function getInicial() {
  const email = emailAtual_();
  const usuario = buscarUsuario_(email);
  if (!usuario) return { semAcesso: true, email: email };
  return getDados();
}

function getDados() {
  const usuario = usuarioAtual_();
  const lidos = JSON.parse(PropertiesService.getUserProperties().getProperty('lidos') || '{}');
  return {
    usuario: usuario,
    itens: lerTabela_(ABAS.ITENS).map(limpar_),
    cotacoes: lerTabela_(ABAS.COTACOES).filter(c => ativo_(c.Ativa)).map(limpar_),
    comentarios: lerTabela_(ABAS.COMENTARIOS).map(limpar_),
    historico: lerTabela_(ABAS.HISTORICO).map(limpar_),
    projetos: lerTabela_(ABAS.PROJETOS).map(p => Object.assign(limpar_(p), { Ativo: ativo_(p.Ativo) })),
    usuarios: lerTabela_(ABAS.USUARIOS).map(u => ({
      email: normalizarEmail_(u.Email),
      nome: String(u.Nome || u.Email),
      papel: papel_(u.Papel),
    })),
    lidos: lidos,
    agora: new Date().toISOString(),
  };
}

/** Guarda até quando o usuário já leu a conversa de um item (por usuário). */
function marcarLido(itemId) {
  usuarioAtual_();
  const props = PropertiesService.getUserProperties();
  const lidos = JSON.parse(props.getProperty('lidos') || '{}');
  lidos[itemId] = new Date().toISOString();
  props.setProperty('lidos', JSON.stringify(lidos));
  return lidos;
}

// ---------------------------------------------------------------------------
// Solicitações (itens)
// ---------------------------------------------------------------------------

/** Lança vários itens de uma vez num projeto (todos entram "Em cotação"). */
function criarItens(projeto, itens) {
  const ids = comLock_(() => {
    const u = usuarioAtual_();
    if (!itens || !itens.length) throw new Error('Preencha pelo menos um item.');
    const validados = itens.map((d, k) => {
      try {
        return validarItem_(Object.assign({}, d, { projeto: projeto }));
      } catch (e) {
        throw new Error('Linha ' + (k + 1) + ': ' + e.message);
      }
    });
    let proximo = proximoNumeroItem_();
    const agora = new Date();
    return validados.map(campos => {
      const id = 'SC-' + String(proximo++).padStart(4, '0');
      inserir_(ABAS.ITENS, Object.assign(campos, {
        ID: id, CriadoEm: agora, CriadoPor: u.email, Status: STATUS.EM_COTACAO, AtualizadoEm: agora,
      }));
      registrar_(id, u, 'Criou a solicitação', '');
      return id;
    });
  });
  return Object.assign(getDados(), { retorno: ids });
}

function editarItem(itemId, dados) {
  comItem_(itemId, [STATUS.EM_COTACAO, STATUS.AGUARDANDO], (item, u, atualizar) => {
    exigirDonoOuAprovador_(item.CriadoPor, u);
    const campos = validarItem_(dados);
    atualizar(campos);
    if (Number(campos.Quantidade) !== Number(item.Quantidade)) recalcularTotais_(itemId, campos.Quantidade);
    registrar_(itemId, u, 'Editou a solicitação', '');
  });
  return getDados();
}

function enviarParaAprovacao(itemIds) {
  emLote_(itemIds, id => {
    comItem_(id, [STATUS.EM_COTACAO], (item, u, atualizar) => {
      if (!cotacoesDoItem_(id).length) {
        throw new Error('Adicione pelo menos uma cotação antes de enviar para aprovação.');
      }
      atualizar({ Status: STATUS.AGUARDANDO });
      registrar_(id, u, 'Enviou para aprovação', '');
    });
  });
  return getDados();
}

/** escolhas: [{ itemId, cotacaoId }] — um ou vários itens. */
function aprovar(escolhas, observacao) {
  emLote_(escolhas, e => {
    comItem_(e.itemId, [STATUS.EM_COTACAO, STATUS.AGUARDANDO], (item, u, atualizar) => {
      exigirAprovador_(u);
      const cotacao = cotacoesDoItem_(e.itemId).find(c => String(c.ID) === String(e.cotacaoId));
      if (!cotacao) throw new Error('Escolha uma das cotações deste item para aprovar.');
      const obs = texto_(observacao);
      atualizar({
        Status: STATUS.APROVADO,
        CotacaoAprovadaID: cotacao.ID,
        ValorAprovado: Number(cotacao.ValorTotal) || 0,
        DecididoPor: u.email,
        DecididoEm: new Date(),
        MotivoDecisao: obs,
      });
      registrar_(e.itemId, u, 'Aprovou a cotação de ' + cotacao.Fornecedor + ' (' + moeda_(cotacao.ValorTotal) + ')', obs);
    });
  });
  return getDados();
}

/** Roda a ação em cada item; se algum falhar, junta os erros numa mensagem só. */
function emLote_(lista, fn) {
  const itens = Array.isArray(lista) ? lista : [lista];
  if (!itens.length) throw new Error('Nenhum item selecionado.');
  const erros = [];
  itens.forEach(x => {
    try {
      fn(x);
    } catch (e) {
      erros.push((x && x.itemId ? x.itemId : x) + ': ' + e.message);
    }
  });
  if (erros.length === itens.length) throw new Error(erros.join('\n'));
  if (erros.length) throw new Error('Alguns itens não foram processados:\n' + erros.join('\n'));
}

function reprovar(itemId, motivo) {
  comItem_(itemId, [STATUS.EM_COTACAO, STATUS.AGUARDANDO], (item, u, atualizar) => {
    exigirAprovador_(u);
    const m = obrigatorio_(motivo, 'Informe o motivo da reprovação.');
    atualizar({ Status: STATUS.REPROVADO, DecididoPor: u.email, DecididoEm: new Date(), MotivoDecisao: m });
    registrar_(itemId, u, 'Reprovou', m);
  });
  return getDados();
}

/** Aprovador devolve para os compradores cotarem de novo (ex.: pedir mais opções). */
function devolverParaCotacao(itemId, motivo) {
  comItem_(itemId, [STATUS.AGUARDANDO], (item, u, atualizar) => {
    exigirAprovador_(u);
    const m = obrigatorio_(motivo, 'Diga o que precisa mudar na cotação.');
    atualizar({ Status: STATUS.EM_COTACAO });
    registrar_(itemId, u, 'Devolveu para cotação', m);
  });
  return getDados();
}

function reabrir(itemId) {
  comItem_(itemId, [STATUS.REPROVADO], (item, u, atualizar) => {
    exigirDonoOuAprovador_(item.CriadoPor, u);
    atualizar({ Status: STATUS.EM_COTACAO, DecididoPor: '', DecididoEm: '', MotivoDecisao: '' });
    registrar_(itemId, u, 'Reabriu para nova cotação', '');
  });
  return getDados();
}

function marcarComprado(itemId, numeroPedido) {
  comItem_(itemId, [STATUS.APROVADO], (item, u, atualizar) => {
    const numero = texto_(numeroPedido);
    atualizar({ Status: STATUS.COMPRADO, CompradoPor: u.email, CompradoEm: new Date(), NumeroPedido: numero });
    registrar_(itemId, u, 'Marcou como comprado', numero ? 'Pedido/NF: ' + numero : '');
  });
  return getDados();
}

function marcarEntregue(itemId) {
  comItem_(itemId, [STATUS.COMPRADO], (item, u, atualizar) => {
    atualizar({ Status: STATUS.ENTREGUE, EntregueEm: new Date(), RecebidoPor: u.email });
    registrar_(itemId, u, 'Confirmou a entrega', '');
  });
  return getDados();
}

function cancelar(itemId, motivo) {
  comItem_(itemId, [STATUS.EM_COTACAO, STATUS.AGUARDANDO, STATUS.APROVADO], (item, u, atualizar) => {
    exigirDonoOuAprovador_(item.CriadoPor, u);
    const m = obrigatorio_(motivo, 'Informe o motivo do cancelamento.');
    atualizar({ Status: STATUS.CANCELADO });
    registrar_(itemId, u, 'Cancelou', m);
  });
  return getDados();
}

// ---------------------------------------------------------------------------
// Cotações
// ---------------------------------------------------------------------------

/** Cria (sem cotacao.id) ou edita uma cotação do item. */
function salvarCotacao(itemId, cotacao) {
  comItem_(itemId, [STATUS.EM_COTACAO, STATUS.AGUARDANDO], (item, u) => {
    const campos = validarCotacao_(cotacao);
    if (!cotacao.id) {
      inserirCotacao_(itemId, campos, item.Quantidade, u);
      registrar_(itemId, u, 'Adicionou cotação de ' + campos.Fornecedor, moeda_(totalCotacao_(campos, item.Quantidade)));
      return;
    }
    const existente = buscarCotacao_(cotacao.id, itemId);
    exigirDonoOuAprovador_(existente.CriadoPor, u);
    campos.ValorTotal = totalCotacao_(campos, item.Quantidade);
    atualizarPorId_(ABAS.COTACOES, existente.ID, campos);
    registrar_(itemId, u, 'Editou cotação de ' + campos.Fornecedor, moeda_(campos.ValorTotal));
  });
  return getDados();
}

function removerCotacao(itemId, cotacaoId) {
  comItem_(itemId, [STATUS.EM_COTACAO, STATUS.AGUARDANDO], (item, u, atualizar) => {
    const existente = buscarCotacao_(cotacaoId, itemId);
    exigirDonoOuAprovador_(existente.CriadoPor, u);
    atualizarPorId_(ABAS.COTACOES, existente.ID, { Ativa: 'NAO' });
    registrar_(itemId, u, 'Removeu cotação de ' + existente.Fornecedor, '');
    if (item.Status === STATUS.AGUARDANDO && !cotacoesDoItem_(itemId).length) {
      atualizar({ Status: STATUS.EM_COTACAO });
      registrar_(itemId, u, 'Voltou para cotação (sem cotações)', '');
    }
  });
  return getDados();
}

// ---------------------------------------------------------------------------
// Conversa e projetos
// ---------------------------------------------------------------------------

function comentar(itemId, textoComentario) {
  comLock_(() => {
    const u = usuarioAtual_();
    const t = obrigatorio_(textoComentario, 'Escreva uma mensagem.');
    if (!lerTabela_(ABAS.ITENS).some(i => String(i.ID) === String(itemId))) {
      throw new Error('Solicitação ' + itemId + ' não encontrada.');
    }
    inserir_(ABAS.COMENTARIOS, { ID: novoId_('M'), ItemID: itemId, Autor: u.email, Texto: t, CriadoEm: new Date() });
  });
  marcarLido(itemId);
  return getDados();
}

function criarProjeto(nome, cliente) {
  const projeto = comLock_(() => {
    const u = usuarioAtual_();
    const n = obrigatorio_(nome, 'Informe o nome do projeto.');
    const existentes = lerTabela_(ABAS.PROJETOS);
    const repetido = existentes.find(p => String(p.Nome).trim().toLowerCase() === n.toLowerCase());
    if (repetido) {
      if (!ativo_(repetido.Ativo)) atualizarPorId_(ABAS.PROJETOS, repetido.ID, { Ativo: 'SIM' });
      return String(repetido.Nome);
    }
    const numeros = existentes.map(p => parseInt(String(p.ID).replace(/\D/g, ''), 10) || 0);
    const id = 'P-' + String(Math.max.apply(null, numeros.concat(0)) + 1).padStart(3, '0');
    inserir_(ABAS.PROJETOS, {
      ID: id, Nome: n, Cliente: texto_(cliente), Ativo: 'SIM', CriadoPor: u.email, CriadoEm: new Date(),
    });
    return n;
  });
  return Object.assign(getDados(), { retorno: projeto });
}

/** Encerrar (some da tela inicial) ou reabrir um projeto. Só o aprovador. */
function definirProjetoAtivo(projetoId, ativo) {
  comLock_(() => {
    exigirAprovador_(usuarioAtual_());
    atualizarPorId_(ABAS.PROJETOS, projetoId, { Ativo: ativo ? 'SIM' : 'NAO' });
  });
  return getDados();
}

// ---------------------------------------------------------------------------
// Regras e validação
// ---------------------------------------------------------------------------

function validarItem_(d) {
  const quantidade = Number(d.quantidade);
  if (!(quantidade > 0)) throw new Error('Informe uma quantidade maior que zero.');
  return {
    Projeto: obrigatorio_(d.projeto, 'Escolha o projeto.'),
    Descricao: obrigatorio_(d.descricao, 'Descreva o item.'),
    Quantidade: quantidade,
    Unidade: texto_(d.unidade) || 'un',
    Urgencia: ['BAIXA', 'NORMAL', 'ALTA'].indexOf(d.urgencia) >= 0 ? d.urgencia : 'NORMAL',
    NecessarioAte: /^\d{4}-\d{2}-\d{2}$/.test(d.necessarioAte || '') ? new Date(d.necessarioAte + 'T12:00:00') : '',
    Observacao: texto_(d.observacao),
  };
}

function validarCotacao_(c) {
  const valor = Number(c.valorUnitario);
  const frete = Number(c.frete || 0);
  if (!(valor > 0)) throw new Error('Informe o valor unitário da cotação.');
  if (!(frete >= 0)) throw new Error('Frete inválido.');
  return {
    Fornecedor: obrigatorio_(c.fornecedor, 'Informe o fornecedor.'),
    Canal: obrigatorio_(c.canal, 'Informe onde foi cotado.'),
    Link: texto_(c.link),
    ValorUnitario: valor,
    Frete: frete,
    FormaPagamento: obrigatorio_(c.formaPagamento, 'Informe a forma de pagamento.'),
    CondicaoPagamento: texto_(c.condicaoPagamento),
    PrazoEntrega: texto_(c.prazoEntrega),
    Observacao: texto_(c.observacao),
  };
}

function inserirCotacao_(itemId, campos, quantidade, u) {
  inserir_(ABAS.COTACOES, Object.assign({}, campos, {
    ID: novoId_('C'),
    ItemID: itemId,
    ValorTotal: totalCotacao_(campos, quantidade),
    CriadoPor: u.email,
    CriadoEm: new Date(),
    Ativa: 'SIM',
  }));
}

function totalCotacao_(campos, quantidade) {
  return Math.round((Number(campos.ValorUnitario) * Number(quantidade) + Number(campos.Frete || 0)) * 100) / 100;
}

function recalcularTotais_(itemId, quantidade) {
  cotacoesDoItem_(itemId).forEach(c => {
    atualizarPorId_(ABAS.COTACOES, c.ID, { ValorTotal: totalCotacao_(c, quantidade) });
  });
}

function cotacoesDoItem_(itemId) {
  return lerTabela_(ABAS.COTACOES).filter(c => String(c.ItemID) === String(itemId) && ativo_(c.Ativa));
}

function buscarCotacao_(cotacaoId, itemId) {
  const c = cotacoesDoItem_(itemId).find(x => String(x.ID) === String(cotacaoId));
  if (!c) throw new Error('Cotação não encontrada. Atualize a página.');
  return c;
}

function exigirAprovador_(u) {
  if (u.papel !== PAPEL.APROVADOR) throw new Error('Só o aprovador pode fazer isso.');
}

function exigirDonoOuAprovador_(criadoPor, u) {
  if (u.papel !== PAPEL.APROVADOR && normalizarEmail_(criadoPor) !== u.email) {
    throw new Error('Só quem criou ou o aprovador pode fazer isso.');
  }
}

/** Abre a linha do item, confere o status e roda a ação dentro do lock. */
function comItem_(itemId, statusPermitidos, fn) {
  return comLock_(() => {
    const u = usuarioAtual_();
    const item = lerTabela_(ABAS.ITENS).find(i => String(i.ID) === String(itemId));
    if (!item) throw new Error('Solicitação ' + itemId + ' não encontrada.');
    if (statusPermitidos.indexOf(item.Status) < 0) {
      throw new Error('Esta solicitação está "' + (ROTULO_STATUS[item.Status] || item.Status) +
        '" e não permite essa ação. Atualize a página.');
    }
    const atualizar = campos => atualizarPorId_(ABAS.ITENS, item.ID, Object.assign({ AtualizadoEm: new Date() }, campos));
    return fn(item, u, atualizar);
  });
}

function registrar_(itemId, u, acao, detalhe) {
  inserir_(ABAS.HISTORICO, { ItemID: itemId, Data: new Date(), Usuario: u.email, Acao: acao, Detalhe: detalhe || '' });
}

// ---------------------------------------------------------------------------
// Usuário
// ---------------------------------------------------------------------------

function emailAtual_() {
  const email = normalizarEmail_(Session.getActiveUser().getEmail());
  if (!email) throw new Error('Não foi possível identificar sua conta Google. Entre com sua conta e recarregue.');
  return email;
}

function buscarUsuario_(email) {
  const u = lerTabela_(ABAS.USUARIOS).find(r => normalizarEmail_(r.Email) === email);
  if (!u || !ativo_(u.Ativo)) return null;
  return { email: email, nome: String(u.Nome || email), papel: papel_(u.Papel) };
}

function usuarioAtual_() {
  const email = emailAtual_();
  const u = buscarUsuario_(email);
  if (!u) throw new Error('O e-mail ' + email + ' não está liberado no sistema de compras.');
  return u;
}

function papel_(v) {
  return String(v).trim().toUpperCase() === PAPEL.APROVADOR ? PAPEL.APROVADOR : PAPEL.COMPRADOR;
}

function normalizarEmail_(v) {
  return String(v || '').trim().toLowerCase();
}

// ---------------------------------------------------------------------------
// Acesso à planilha
// ---------------------------------------------------------------------------

function planilha_() {
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  return id ? SpreadsheetApp.openById(id) : SpreadsheetApp.getActive();
}

function aba_(def) {
  const sh = planilha_().getSheetByName(def.nome);
  if (!sh) throw new Error('Aba "' + def.nome + '" não encontrada. Rode Sistema de Compras > Configurar planilha.');
  return sh;
}

function cabecalho_(sh) {
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String);
}

/** Lê a aba como lista de objetos { Coluna: valor, _linha: n }. */
function lerTabela_(def) {
  const sh = aba_(def);
  const valores = sh.getDataRange().getValues();
  const cab = (valores.shift() || []).map(String);
  const linhas = [];
  valores.forEach((l, i) => {
    if (!l.some(v => v !== '' && v !== null)) return;
    const obj = { _linha: i + 2 };
    cab.forEach((c, j) => { if (c) obj[c] = l[j]; });
    linhas.push(obj);
  });
  return linhas;
}

function inserir_(def, obj) {
  const sh = aba_(def);
  const cab = cabecalho_(sh);
  sh.appendRow(cab.map(c => paraCelula_(obj[c])));
}

function atualizarPorId_(def, id, campos) {
  const sh = aba_(def);
  const cab = cabecalho_(sh);
  const linha = lerTabela_(def).find(r => String(r.ID) === String(id));
  if (!linha) throw new Error('Registro ' + id + ' não encontrado em ' + def.nome + '.');
  const range = sh.getRange(linha._linha, 1, 1, cab.length);
  const valores = range.getValues()[0];
  Object.keys(campos).forEach(k => {
    const i = cab.indexOf(k);
    if (i < 0) throw new Error('Coluna "' + k + '" não existe em ' + def.nome + '. Rode Configurar planilha.');
    valores[i] = paraCelula_(campos[k]);
  });
  range.setValues([valores]);
}

/** Evita que textos como "=..." ou "+55..." virem fórmula na planilha. */
function paraCelula_(v) {
  if (v === undefined || v === null) return '';
  if (typeof v === 'string' && /^[=+\-@]/.test(v)) return "'" + v;
  return v;
}

function limpar_(obj) {
  const out = {};
  Object.keys(obj).forEach(k => {
    if (k === '_linha') return;
    const v = obj[k];
    out[k] = v instanceof Date ? v.toISOString() : v;
  });
  return out;
}

function proximoNumeroItem_() {
  const numeros = lerTabela_(ABAS.ITENS).map(i => parseInt(String(i.ID).replace(/\D/g, ''), 10) || 0);
  return Math.max.apply(null, numeros.concat(0)) + 1;
}

function novoId_(prefixo) {
  return prefixo + '-' + Utilities.getUuid().replace(/-/g, '').slice(0, 10).toUpperCase();
}

function comLock_(fn) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) throw new Error('O sistema está ocupado salvando outra alteração. Tente de novo.');
  try {
    return fn();
  } finally {
    SpreadsheetApp.flush();
    lock.releaseLock();
  }
}

function ativo_(v) {
  if (v === false) return false;
  return !/^(n|nao|não|false|0|inativo)$/i.test(String(v).trim());
}

function texto_(v) {
  return String(v === undefined || v === null ? '' : v).trim();
}

function obrigatorio_(v, mensagem) {
  const t = texto_(v);
  if (!t) throw new Error(mensagem);
  return t;
}

function moeda_(v) {
  return 'R$ ' + (Number(v) || 0).toFixed(2).replace('.', ',').replace(/\B(?=(\d{3})+(?!\d))/g, '.');
}
