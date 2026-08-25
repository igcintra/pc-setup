/**
 * CONTADOR DE USO DOS SCRIPTS DE TI — v2 (25/08/2026)
 * Planilha no Gmail pessoal do Gabriel. Editar em: Extensoes > Apps Script.
 *
 * COMO USAR: apagar TODO o conteudo do editor, colar este arquivo, salvar e
 * fazer "Implantar > Nova implantacao > App da Web"
 *   Executar como: Eu    |    Quem tem acesso: Qualquer pessoa
 * Isso gera uma URL NOVA — e a URL antiga (que vazou no historico do git)
 * deixa de responder.
 *
 * O QUE MUDOU DA v1
 *  - Nao usa mais getActiveSheet(): as abas sao buscadas PELO NOME e criadas se
 *    faltarem. Era isso que fazia o contador escrever na aba errada depois de
 *    qualquer insertSheet().
 *  - Trava de concorrencia (LockService): dois PCs rodando o setup ao mesmo
 *    tempo nao se sobrescrevem mais.
 *  - Aba "historico": uma linha por execucao (data, script, PC, usuario, versao).
 *    O contador so diz "quantas vezes"; o historico diz QUANDO, ONDE e COM QUEM
 *    — que e o que serve pra decidir qual script vale manter.
 *  - Aceita GET e POST e responde JSON, nunca HTML.
 *  - Qualquer erro vira {"ok":false}: o script do PC nunca trava por causa daqui.
 */

var ABA_CONTADOR  = 'contador';
var ABA_HISTORICO = 'historico';
var MAX_HISTORICO = 20000;   // linhas; acima disso apaga as mais antigas
var FUSO          = 'America/Sao_Paulo';

function doGet(e)  { return registrar(e); }
function doPost(e) { return registrar(e); }

function registrar(e) {
  var lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);                       // espera a vez, ate 20s

    var p = (e && e.parameter) ? e.parameter : {};
    var dados = {
      script:  limpar(p.script,  60) || 'desconhecido',
      host:    limpar(p.host,    60),
      usuario: limpar(p.usuario, 60),
      versao:  limpar(p.versao,  20),
      extra:   extras(p)
    };

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var total = contar(ss, dados.script);
    historiar(ss, dados);

    return json({ ok: true, script: dados.script, total: total });

  } catch (err) {
    return json({ ok: false, erro: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

/** Soma 1 na linha do script (cria se for a primeira vez). Devolve o total. */
function contar(ss, script) {
  var aba = abaPorNome(ss, ABA_CONTADOR, ['script', 'execucoes', 'primeiro uso', 'ultimo uso']);
  var valores = aba.getDataRange().getValues();

  for (var i = 1; i < valores.length; i++) {
    if (String(valores[i][0]) === script) {
      var novo = Number(valores[i][1] || 0) + 1;
      aba.getRange(i + 1, 2).setValue(novo);
      aba.getRange(i + 1, 4).setValue(agora());
      return novo;
    }
  }
  aba.appendRow([script, 1, agora(), agora()]);
  return 1;
}

/** Uma linha por execucao. E aqui que da pra ver uso real por semana. */
function historiar(ss, d) {
  var aba = abaPorNome(ss, ABA_HISTORICO, ['data/hora', 'script', 'PC', 'usuario', 'versao', 'extras']);
  aba.appendRow([agora(), d.script, d.host, d.usuario, d.versao, d.extra]);

  var linhas = aba.getLastRow() - 1;
  if (linhas > MAX_HISTORICO) {
    aba.deleteRows(2, linhas - MAX_HISTORICO);   // apaga as mais antigas
  }
}

/** Devolve a aba pelo NOME; cria com cabecalho se nao existir. */
function abaPorNome(ss, nome, cabecalho) {
  var s = ss.getSheetByName(nome);
  if (!s) {
    s = ss.insertSheet(nome);
    s.appendRow(cabecalho);
    s.getRange(1, 1, 1, cabecalho.length).setFontWeight('bold');
    s.setFrozenRows(1);
  }
  return s;
}

/** Junta os parametros que nao sao os fixos, pra nao perder nada que o script mandar. */
function extras(p) {
  var fixos = { script: 1, host: 1, usuario: 1, versao: 1 };
  var out = [];
  for (var k in p) {
    if (!fixos[k]) out.push(k + '=' + limpar(p[k], 120));
  }
  return out.join(' | ').substring(0, 400);
}

/** Corta tamanho e tira quebra de linha — planilha nao gosta de texto solto. */
function limpar(v, max) {
  if (v === undefined || v === null) return '';
  return String(v).replace(/[\r\n\t]/g, ' ').trim().substring(0, max);
}

function agora() {
  return Utilities.formatDate(new Date(), FUSO, 'yyyy-MM-dd HH:mm:ss');
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Teste sem sair do editor: rodar esta funcao e ver o Registro de execucoes.
 * Deve criar as duas abas e escrever uma linha de teste.
 */
function testar() {
  var r = registrar({ parameter: { script: 'teste-editor', host: 'PC-TESTE', usuario: 'gcintra', versao: 'v2' } });
  Logger.log(r.getContent());
}
