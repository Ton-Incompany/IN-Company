/**
 * IN Company Chilli Beans - Prospecção B2B
 * Guarda os dados em uma planilha do Google criada na sua conta.
 * Implante como "Aplicativo da Web": executar como VOCÊ, quem pode acessar: QUALQUER PESSOA (a chave abaixo protege os dados).
 */
var NOME_PLANILHA = 'IN Company Chilli Beans - Prospecção';
var ESTRUTURA = '3'; // muda quando as colunas da planilha mudam (a planilha é reorganizada sozinha)
var STATUS = ['Nova empresa', 'Em follow-up', 'Em conversa', 'Reunião agendada', 'Em orçamento',
              'Concluída', 'Encerrado sem sucesso'];
// Aba Empresas: A = id, B = dados (JSON, oculta) e, a partir da C, as colunas abaixo.
var COLS = ['Empresa', 'Status', 'Origem do 1º contato', '1º contato', 'Franqueado', 'Canal', 'Consultor',
            'Estado', 'Cidade', 'SAP', 'Loja', 'Telefone do franqueado', 'Contatos',
            'Valor a faturar', 'Valor faturado', 'Data da venda', 'Último follow-up'];
// Números das colunas (1 = A): F = 1º contato, P = Valor a faturar, Q = Valor faturado, R = Data da venda.
var COL_D1 = 6, COL_AFATURAR = 16, COL_FATURADO = 17, COL_DATAVENDA = 18;

// CHAVE DE ACESSO: invente uma senha longa (12+ caracteres) e use a mesma no app, na tela "Conexão com a planilha".
var CHAVE = 'TROQUE-ESTA-CHAVE';

var FUNCOES = { getAll: getAll, saveItem: saveItem, removeItem: removeItem, saveCfg: saveCfg, enviarCodigo: enviarCodigo };

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return ContentService.createTextOutput('IN Company Chilli Beans: serviço de dados ativo.');
}

function doPost(e) {
  try {
    var b = JSON.parse(e.postData.contents);
    if (!CHAVE || CHAVE === 'TROQUE-ESTA-CHAVE' || b.k !== CHAVE) return json_({ erro: 'chave' });
    var f = FUNCOES[b.fn];
    if (!f) return json_({ erro: 'função desconhecida' });
    return json_({ r: f.apply(null, b.a || []) });
  } catch (err) {
    return json_({ erro: String((err && err.message) || err) });
  }
}

/** Rode esta função uma vez no editor para criar/atualizar a planilha antes de abrir o aplicativo. */
function montarPlanilha() {
  var ss = planilha_();
  Logger.log('Planilha pronta: ' + ss.getUrl());
}

/** Abre (ou cria na primeira vez) a planilha de dados e mantém a estrutura atualizada. */
function planilha_() {
  var props = PropertiesService.getScriptProperties();
  var id = props.getProperty('SHEET_ID');
  var ss = null;
  if (id) {
    try { ss = SpreadsheetApp.openById(id); } catch (e) { ss = null; }
  }
  if (!ss) {
    ss = SpreadsheetApp.create(NOME_PLANILHA);
    ss.setSpreadsheetTimeZone('America/Sao_Paulo');
    props.setProperty('SHEET_ID', ss.getId());
    var a = ss.getSheets()[0];
    a.setName('Empresas');
    a.getRange('A:B').setNumberFormat('@');
    cabecalho_(a);
    var c = ss.insertSheet('Config');
    c.getRange('A:B').setNumberFormat('@');
    c.getRange(1, 1, 1, 2).setValues([['chave', 'valor']]).setFontWeight('bold');
    c.hideSheet();
    resumo_(ss);
    props.setProperty('ESTRUTURA', ESTRUTURA);
  } else if (props.getProperty('ESTRUTURA') !== ESTRUTURA) {
    migrar_(ss, props);
  }
  return ss;
}

function cabecalho_(a) {
  a.getRange(1, 1, 1, 2 + COLS.length).setValues([['id', 'dados (JSON)'].concat(COLS)])
    .setFontWeight('bold').setBackground('#111111').setFontColor('#ffffff');
  a.setFrozenRows(1);
  try { a.hideColumns(2); } catch (e) {}
  a.setColumnWidth(3, 220);
  a.setColumnWidth(15, 320);
}

function formatar_(a, primeira, qtd) {
  if (qtd < 1) return;
  a.getRange(primeira, COL_D1, qtd, 1).setNumberFormat('dd/mm/yyyy');
  a.getRange(primeira, COL_DATAVENDA, qtd, 1).setNumberFormat('dd/mm/yyyy');
  a.getRange(primeira, COL_AFATURAR, qtd, 2).setNumberFormat('R$ #,##0.00');
}

/** Reorganiza a planilha quando as colunas mudam: reescreve tudo a partir dos dados (JSON). */
function migrar_(ss, props) {
  var a = ss.getSheetByName('Empresas');
  var last = a.getLastRow();
  var itens = [];
  if (last > 1) {
    a.getRange(2, 1, last - 1, 2).getValues().forEach(function (r) {
      if (r[0] && r[1]) {
        try { var o = JSON.parse(r[1]); o.id = String(r[0]); itens.push(o); } catch (e) {}
      }
    });
  }
  var larg = Math.max(a.getMaxColumns() - 2, 1);
  a.getRange(1, 3, Math.max(a.getMaxRows(), 1), larg).clearContent();
  cabecalho_(a);
  itens.forEach(function (o, i) {
    var linha = legivel_(ss, o);
    a.getRange(i + 2, 3, 1, linha.length).setValues([linha]);
  });
  formatar_(a, 2, itens.length);
  var r = ss.getSheetByName('Resumo');
  if (r) ss.deleteSheet(r);
  resumo_(ss);
  props.setProperty('ESTRUTURA', ESTRUTURA);
}

/** Aba com totais e quantidades, calculados por fórmulas (atualizam sozinhos). */
function resumo_(ss) {
  var r = ss.insertSheet('Resumo', 0);
  var de = 'IF($B$3="",DATE(2000,1,1),$B$3)';
  var ate = 'IF($B$4="",DATE(2100,1,1),$B$4)';
  r.getRange('A1').setValue('Resumo - IN Company Chilli Beans').setFontWeight('bold').setFontSize(14);
  r.getRange('A3').setValue('Vendas de (data):');
  r.getRange('A4').setValue('Vendas até (data):');
  r.getRange('B3:B4').setNumberFormat('dd/mm/yyyy').setBackground('#fff2cc');
  r.getRange('C3').setValue('Preencha as datas para filtrar o período. Em branco = todo o período.')
    .setFontColor('#666666');
  r.getRange('A6').setValue('Valor faturado no período').setFontWeight('bold');
  r.getRange('B6').setFormula('=SUMIFS(Empresas!Q:Q,Empresas!D:D,"Concluída",Empresas!R:R,">="&' + de +
    ',Empresas!R:R,"<="&' + ate + ')').setNumberFormat('R$ #,##0.00').setFontWeight('bold');
  r.getRange('A7').setValue('Vendas concluídas no período');
  r.getRange('B7').setFormula('=COUNTIFS(Empresas!D:D,"Concluída",Empresas!R:R,">="&' + de +
    ',Empresas!R:R,"<="&' + ate + ')');
  r.getRange('A8').setValue('Valor a faturar (em andamento)');
  r.getRange('B8').setFormula('=SUM(Empresas!P:P)').setNumberFormat('R$ #,##0.00');
  r.getRange('A10').setValue('Empresas por status').setFontWeight('bold');
  for (var i = 0; i < STATUS.length; i++) {
    r.getRange(11 + i, 1).setValue(STATUS[i]);
    r.getRange(11 + i, 2).setFormula('=COUNTIF(Empresas!D:D,A' + (11 + i) + ')');
  }
  var t = 11 + STATUS.length;
  r.getRange(t, 1).setValue('Total de empresas').setFontWeight('bold');
  r.getRange(t, 2).setFormula('=SUM(B11:B' + (t - 1) + ')').setFontWeight('bold');
  // Faturado por franqueado, canal, consultor e origem (todo o período)
  var blocos = [['D', 'G', 'Franqueado'], ['H', 'H', 'Canal'], ['L', 'I', 'Consultor'], ['P', 'E', 'Origem']];
  blocos.forEach(function (b) {
    r.getRange(b[0] + '10').setValue('Faturado por ' + b[2].toLowerCase()).setFontWeight('bold');
    r.getRange(b[0] + '11').setFormula('=IFERROR(QUERY(Empresas!A:S,"select ' + b[1] + ', count(A), sum(Q) ' +
      'where D = \'Concluída\' and ' + b[1] + ' <> \'\' group by ' + b[1] + ' order by sum(Q) desc ' +
      'label ' + b[1] + ' \'' + b[2] + '\', count(A) \'Vendas\', sum(Q) \'Total faturado\'",1),"Sem vendas")');
  });
  r.setColumnWidth(1, 260);
  r.setColumnWidth(2, 140);
}

/** Converte "aaaa-mm-dd" em data real da planilha (ou vazio). */
function data_(ss, txt) {
  if (!txt || !/^\d{4}-\d{2}-\d{2}$/.test(txt)) return '';
  return Utilities.parseDate(txt, ss.getSpreadsheetTimeZone(), 'yyyy-MM-dd');
}

/** Colunas legíveis (somente para consulta na planilha). Valor a faturar x Valor faturado: depende do status. */
function legivel_(ss, o) {
  var fr = o.fr || {};
  var cs = (o.contatos || []).map(function (c) {
    return [c.nome, c.cargo, c.telefone, c.email, c.linkedin].filter(Boolean).join(' / ');
  }).join(' | ');
  var fu = (o.followups || []).slice().sort(function (a, b) { return a.data < b.data ? 1 : -1; })[0];
  var valor = Number(o.valorVenda) || 0;
  var vendida = o.status === 'Concluída';
  return [o.empresa || '', o.status || '', o.canal || '', data_(ss, o.d1), fr.nome || '', fr.canal || '',
          fr.consultor || '', fr.uf || fr.regiao || '', fr.cidade || '', fr.sap || '',
          fr.loja || '', fr.telefone || '', cs,
          vendida ? '' : (valor || ''), vendida ? valor : '', vendida ? data_(ss, o.dataVenda) : '',
          fu ? (fu.data + ' ' + fu.canal) : ''];
}

function linhaDoId_(sh, id) {
  var last = sh.getLastRow();
  if (last < 2) return 0;
  var ids = sh.getRange(2, 1, last - 1, 1).getValues();
  for (var i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) return i + 2;
  }
  return 0;
}

/** Devolve todos os dados para o aplicativo. */
function getAll() {
  var ss = planilha_();
  var sh = ss.getSheetByName('Empresas');
  var items = [];
  var last = sh.getLastRow();
  if (last > 1) {
    sh.getRange(2, 1, last - 1, 2).getValues().forEach(function (r) {
      if (r[0] && r[1]) {
        try { var o = JSON.parse(r[1]); o.id = String(r[0]); items.push(o); } catch (e) {}
      }
    });
  }
  return { items: items, senha: lerCfg_(ss, 'senha'), app: lerCfg_(ss, 'app'), url: ss.getUrl() };
}

function saveItem(o) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var ss = planilha_();
    var sh = ss.getSheetByName('Empresas');
    var row = linhaDoId_(sh, o.id) || (sh.getLastRow() + 1);
    var linha = [String(o.id), JSON.stringify(o)].concat(legivel_(ss, o));
    sh.getRange(row, 1, 1, linha.length).setValues([linha]);
    formatar_(sh, row, 1);
    return true;
  } finally {
    lock.releaseLock();
  }
}

function removeItem(id) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = planilha_().getSheetByName('Empresas');
    var row = linhaDoId_(sh, id);
    if (row) sh.deleteRow(row);
    return true;
  } finally {
    lock.releaseLock();
  }
}

function lerCfg_(ss, chave) {
  var sh = ss.getSheetByName('Config');
  var last = sh.getLastRow();
  if (last < 2) return null;
  var v = sh.getRange(2, 1, last - 1, 2).getValues();
  for (var i = 0; i < v.length; i++) {
    if (v[i][0] === chave) {
      try { return JSON.parse(v[i][1]); } catch (e) { return null; }
    }
  }
  return null;
}

/** Guarda a senha (hash) e as configurações do aplicativo. */
function saveCfg(chave, obj) {
  if (chave !== 'senha' && chave !== 'app') throw new Error('chave inválida');
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    var sh = planilha_().getSheetByName('Config');
    var last = sh.getLastRow(), row = 0;
    if (last > 1) {
      var ch = sh.getRange(2, 1, last - 1, 1).getValues();
      for (var i = 0; i < ch.length; i++) { if (ch[i][0] === chave) { row = i + 2; break; } }
    }
    if (!row) row = last + 1;
    sh.getRange(row, 1, 1, 2).setValues([[chave, JSON.stringify(obj)]]);
    return true;
  } finally {
    lock.releaseLock();
  }
}

/** Envia o código de recuperação de senha para o seu e-mail. */
function enviarCodigo(email, codigo) {
  if (!/^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(email || '')) throw new Error('e-mail inválido');
  MailApp.sendEmail({
    to: email,
    subject: 'Código de recuperação - IN Company Chilli Beans',
    body: 'Código de recuperação da senha de exportação do app IN Company Chilli Beans:\n\n' +
          codigo + '\n\nGuarde este e-mail. Use o código em Configurações > Senha e recuperação > Esqueci a senha.'
  });
  return true;
}
