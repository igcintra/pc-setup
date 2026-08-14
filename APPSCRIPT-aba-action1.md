# Aba nova na planilha do contador — resultados da varredura Action1

O `verificar-action1.ps1` chama o **mesmo Web App** do contador de scripts, mas manda
parâmetros extras:

```
?script=verificar-action1&host=NOME-PC&usuario=fulano&veredito=ACHADO|LIMPO&marcadores=registro|programa|log-instalador&guid=xxxx
```

## Código completo do `doGet` (substitui o arquivo inteiro)

Editar em **Extensões → Apps Script** na planilha.

> ⚠️ **Pegadinha que este código já resolve:** o contador original usava
> `SpreadsheetApp.getActiveSpreadsheet().getActiveSheet()`, e **`insertSheet()` torna a
> nova aba a ativa**. Se a função nova fosse só acrescentada, na primeira execução o
> contador passaria a escrever dentro da aba `action1-varredura`. Por isso a referência da
> aba do contador é capturada **antes** de tudo, e o foco é devolvido depois de criar a aba.
> A parte nova também vai em `try/catch`, para nunca derrubar o contador dos outros scripts.

```javascript
function doGet(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var abaContador = ss.getActiveSheet();   // capturado ANTES de qualquer insertSheet()

  var script = (e && e.parameter && e.parameter.script) ? e.parameter.script : "desconhecido";

  // ---------------- CONTADOR (comportamento original) ----------------
  var data  = abaContador.getDataRange().getValues();
  var achou = false;

  for (var i = 1; i < data.length; i++) {
    if (data[i][0] == script) {
      abaContador.getRange(i + 1, 2).setValue(data[i][1] + 1);
      abaContador.getRange(i + 1, 3).setValue(new Date().toLocaleString("pt-BR"));
      achou = true;
      break;
    }
  }

  if (!achou) {
    var novaLinha = data.length + 1;
    abaContador.getRange(novaLinha, 1).setValue(script);
    abaContador.getRange(novaLinha, 2).setValue(1);
    abaContador.getRange(novaLinha, 3).setValue(new Date().toLocaleString("pt-BR"));
  }

  // ---------------- VARREDURA ACTION1 (aba propria) ----------------
  try {
    registrarAction1(e, ss, abaContador);
  } catch (err) {
    // se algo falhar aqui, o contador acima ja rodou e nao e afetado
  }

  return ContentService.createTextOutput("ok");
}


function registrarAction1(e, ss, abaContador) {
  var p = (e && e.parameter) ? e.parameter : {};
  if (p.script !== 'verificar-action1') { return; }

  var aba = ss.getSheetByName('action1-varredura');

  if (!aba) {
    aba = ss.insertSheet('action1-varredura');
    aba.appendRow(['Quando', 'Computador', 'Usuario', 'Veredito', 'Marcadores', 'agent.guid']);
    aba.getRange('A1:F1')
       .setFontWeight('bold')
       .setBackground('#1C4587')
       .setFontColor('#FFFFFF');
    aba.setFrozenRows(1);
    aba.setColumnWidths(1, 6, 160);
    ss.setActiveSheet(abaContador);   // devolve o foco, senao o contador quebra
  }

  aba.appendRow([
    Utilities.formatDate(new Date(), 'America/Sao_Paulo', 'dd/MM/yyyy HH:mm:ss'),
    p.host       || '',
    p.usuario    || '',
    p.veredito   || '',
    p.marcadores || '',
    p.guid       || ''
  ]);

  if (p.veredito === 'ACHADO') {
    aba.getRange(aba.getLastRow(), 1, 1, 6).setBackground('#F4CCCC');
  }
}
```

Depois de colar: **Salvar** → **Implantar → Gerenciar implantações → lápis → Versão: Nova
versão → Implantar**. Sem isso o Web App continua servindo o código antigo (pegadinha
clássica do Apps Script, e a causa nº 1 de "o script não funciona").

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
