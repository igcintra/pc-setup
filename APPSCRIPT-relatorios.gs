/**
 * RECEPTOR DE RELATORIOS DO RASTREADOR — 01/09/2026
 *
 * SEPARADO DO CONTADOR DE PROPOSITO. Projeto proprio, implantacao propria, URL
 * propria. Se um cair, o outro continua. NAO colar isto no projeto do contador.
 *
 * O que faz: recebe por POST o TXT que o rastreador gera e cria o arquivo numa
 * pasta do Drive. Devolve JSON, nunca HTML.
 *
 * ONDE IMPLANTAR: numa conta @ignetworks.com (a do Gabriel). Quem cria o arquivo
 * e' a conta que RODA o script — implantando com a conta corporativa, o arquivo
 * nasce dentro da IGN. E' a diferenca em relacao ao contador, que vive na conta
 * pessoal por causa do bloqueio de login de 2026.
 *
 * COMO IMPLANTAR: script.google.com > Novo projeto > colar isto > Salvar >
 *   Implantar > Nova implantacao > App da Web
 *     Executar como: EU        (assim o arquivo nasce como gcintra@)
 *     Quem tem acesso: QUALQUER PESSOA     <- se essa opcao NAO aparecer, avisar
 * Depois de qualquer alteracao no codigo: Implantar > GERENCIAR implantacoes >
 * editar (lapis) > Versao: Nova > Implantar. Salvar o codigo NAO muda o que a
 * URL ja publicada serve.
 */

var PASTA_ID    = '18G-6-wN-qdQ--AujHjqzGFUlhcpW3duh';  // "Verificacoes de seguranca do PC"
var MAX_BYTES   = 200 * 1024;   // relatorio real tem 2-8 KB
var MAX_POR_DIA = 200;          // freio de vandalismo (a URL e' publica)

// A URL fica dentro de um script em repositorio PUBLICO, entao qualquer um pode
// chamar. Nao da pra proteger com senha (a senha estaria no mesmo arquivo publico).
// Defesa possivel: so aceitar coisa que TEM A CARA do nosso relatorio, limitar
// tamanho e limitar quantidade por dia. Isso reduz a nuisance, nao elimina.
var ASSINATURAS = ['RASTREADOR DE ACESSOS REMOTOS', '[DADOS-JSON]', '[COBERTURA-E-LIMITES]'];

function doPost(e) { return receber(e); }
function doGet(e)  { return json({ ok: false, erro: 'use POST' }); }

function receber(e) {
  try {
    var corpo = (e && e.postData && e.postData.contents) ? e.postData.contents : '';
    if (!corpo)                   return json({ ok: false, erro: 'vazio' });
    if (corpo.length > MAX_BYTES) return json({ ok: false, erro: 'grande demais' });

    for (var i = 0; i < ASSINATURAS.length; i++) {
      if (corpo.indexOf(ASSINATURAS[i]) === -1) {
        return json({ ok: false, erro: 'nao parece um relatorio do rastreador' });
      }
    }

    var pasta = DriveApp.getFolderById(PASTA_ID);
    if (contarHoje(pasta) >= MAX_POR_DIA) return json({ ok: false, erro: 'cota do dia' });

    var p = (e && e.parameter) ? e.parameter : {};
    // HHmm no fim: quem roda 3x no mesmo dia (aconteceu em 01/09) gera 3 arquivos,
    // e da pra ver a sequencia em vez de sobrescrever a primeira passada.
    var nome = [
      'verificacao-seguranca',
      limpar(p.usuario) || 'usuario',
      limpar(p.host)    || 'pc',
      Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'yyyy-MM-dd-HHmm'),
      (limpar(p.resultado) || 'SEMVEREDITO')
    ].join('-') + '.txt';

    var arq = pasta.createFile(nome, corpo, MimeType.PLAIN_TEXT);
    return json({ ok: true, arquivo: nome, id: arq.getId() });

  } catch (err) {
    return json({ ok: false, erro: String(err) });
  }
}

function contarHoje(pasta) {
  var limite = new Date(); limite.setHours(0, 0, 0, 0);
  var n = 0, it = pasta.getFiles();
  while (it.hasNext()) { if (it.next().getDateCreated() >= limite) n++; }
  return n;
}

function limpar(v) {
  return v ? String(v).replace(/[^A-Za-z0-9._-]/g, '').slice(0, 40) : '';
}

function json(o) {
  return ContentService.createTextOutput(JSON.stringify(o))
                       .setMimeType(ContentService.MimeType.JSON);
}
