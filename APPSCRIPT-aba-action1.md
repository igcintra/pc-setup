# Aba nova na planilha do contador — resultados da varredura Action1

O `verificar-action1.ps1` chama o **mesmo Web App** do contador de scripts, mas manda
parâmetros extras:

```
?script=verificar-action1&host=NOME-PC&usuario=fulano&veredito=ACHADO|LIMPO&marcadores=registro|programa|log-instalador&guid=xxxx
```

Hoje o Apps Script provavelmente só incrementa o contador e ignora o resto. Para gravar
numa aba própria, edite o projeto do Apps Script (Extensões → Apps Script na planilha) e
acrescente o trecho abaixo dentro do `doGet`, **antes** do `return`:

```javascript
function doGet(e) {
  var p = (e && e.parameter) ? e.parameter : {};

  // ---- contador existente: NAO MEXER, deixe o codigo que ja esta aqui ----

  // ---- NOVO: varredura Action1 ----
  if (p.script === 'verificar-action1') {
    var ss  = SpreadsheetApp.getActiveSpreadsheet();
    var aba = ss.getSheetByName('action1-varredura');
    if (!aba) {
      aba = ss.insertSheet('action1-varredura');
      aba.appendRow(['Quando','Computador','Usuario','Veredito','Marcadores','agent.guid']);
      aba.getRange('A1:F1').setFontWeight('bold').setBackground('#1C4587').setFontColor('#FFFFFF');
      aba.setFrozenRows(1);
      aba.setColumnWidths(1, 6, 150);
    }
    aba.appendRow([
      Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm:ss'),
      p.host       || '',
      p.usuario    || '',
      p.veredito   || '',
      p.marcadores || '',
      p.guid       || ''
    ]);
    // destaca em vermelho a linha de quem acusou
    if (p.veredito === 'ACHADO') {
      aba.getRange(aba.getLastRow(), 1, 1, 6).setBackground('#F4CCCC');
    }
  }

  return ContentService.createTextOutput('ok');
}
```

Depois de colar: **Implantar → Gerenciar implantações → editar → Nova versão**. Sem isso
o Web App continua servindo o código antigo (pegadinha clássica do Apps Script).

## Como acompanhar a varredura

A aba `action1-varredura` responde as três perguntas que importam:

- **Quantas máquinas já rodaram** — conte as linhas
- **Quem acusou** — linhas em vermelho, coluna `Veredito = ACHADO`
- **Quem ainda não rodou** — cruze a coluna `Usuario` com a lista de funcionários

O `agent.guid` é útil comparar entre máquinas: **se várias acusarem o mesmo GUID de
organização, é a mesma campanha**; GUIDs diferentes por máquina são normais (cada agente
tem o seu), então o que interessa é a existência, não a igualdade.

## Ordem sugerida de distribuição

1. **Accounting e Finance primeiro** — é onde a isca de abril mirou.
2. Depois quem estava entre os 62 destinatários internos do disparo de 14/08.
3. Por último o resto da empresa.

Não distribua para todos de uma vez sem antes ver a primeira dezena de respostas — se o
script tiver algum atrito (bloqueio de execução, política de PowerShell), é melhor
descobrir com 10 pessoas do que com 190.
