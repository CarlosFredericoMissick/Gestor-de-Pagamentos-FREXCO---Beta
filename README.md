# Gestor de Pagamentos Frexco — estrutura

Página estática, sem build. Abra `index.html` servido por HTTP no mesmo origin cadastrado no
app do Azure (o `redirectUri` do MSAL é `window.location.origin`). As bibliotecas vêm por CDN.

```
index.html              marcação + <script> na ordem de carga
logo-frexco-branco.png  logo do cabeçalho
css/styles.css          estilos próprios (o resto é Tailwind via CDN)
js/vendor-config.js     config do Tailwind e do worker do pdf.js (no <head>)
js/config.js            SP_SITE_ID, escopos e config do MSAL, URLs do Graph
js/state.js             estado global compartilhado (notas, prestadores, recebível, conexão)
js/regras.js            regras puras: cruzamento por CNPJ/ID, vínculo manual x automático
                        dos motoristas da nota e cálculo do valor devido (testado em ../testes)
js/utils.js             formatação de moeda, argumento seguro para onclick, loading de botão
js/auth.js              login MSAL, renovação de token, banner de sessão expirada, Atualizar
js/sharepoint.js        leitura/escrita no Graph (prestadores.json, recebivel.json, notas/, pdfs/, contratos)
js/armazenamento-local.js  IndexedDB de prestadores (modo sem SharePoint) e migrações
js/prestadores.js       cruzamento, importações (base, CNPJ/PIX, contratos ZIP), cadastro e lista
js/lideres.js           líderes de time: sincronização, edição, contrato
js/tabela.js            tabela de notas: ordenação, linhas, filtro por time
js/notas.js             reprocessamento, seleção/exclusão, cards de cobertura, edição OCR
js/extracao.js          fila de PDFs, leitura geométrica/OCR e parsers por prefeitura
js/motoristas-nota.js   modal "Motoristas na Nota"
js/conciliacao.js       recebível, conciliação, filtro de times e detalhe do valor devido
js/exportacao.js        Excel/ZIP de notas, prestadores, contratos, conciliação e detalhe
js/main.js              listeners de upload e inicialização (sempre por último)
```

Os scripts são clássicos (não ES modules): funções de topo ficam em `window`, então os
`onclick="..."` do HTML continuam funcionando, e `let`/`const` de topo são visíveis nos
scripts carregados depois. Mantenha a ordem do `index.html`.

## Dados no SharePoint (pasta da aplicação)

Mesmos arquivos da versão anterior. Campos novos em cada `notas/*.json`: `motoristasManual`
(true/false), e, quando os motoristas são editados no modal, `motoristasEditadoEm` e
`motoristasEditadoPor`.

- Notas antigas (sem `motoristasManual`) são classificadas na primeira carga com o SharePoint
  conectado por `migrarNotaLegada` (`js/regras.js`), sem mudar o valor devido que a versão
  anterior calculava. A classificação é gravada uma única vez, só nessas notas, levando apenas
  a flag (e a lista explícita do líder, quando aplicável).
- Gravação de nota usa concorrência otimista: o eTag lido do SharePoint vai em `If-Match`; se
  outro usuário gravou antes (412), a nota é relida, só os campos editados são reaplicados e o
  usuário é avisado se o mesmo campo foi alterado pelos dois.

## Regras decididas pelo usuário

- **Comissão do time: só na nota mais recente do líder.** Entre todas as notas em que o líder
  conta, a comissão do time no período vai para a de data de emissão mais recente (desempate:
  `dataExtracao` mais recente, depois nome do arquivo). A nota não precisa ter sido emitida
  dentro do período, porque as notas costumam sair depois da apuração e a base de notas é
  limpa antes de cada agendamento de pagamento. A atribuição considera todas as notas,
  independente do filtro de time. Implementada em `politicaComissaoNotaMaisRecente` (padrão em
  `politicaComissaoPadrao`, `js/regras.js`); `politicaComissaoPorNota` reproduz a regra anterior.
- **Lista manual é a fonte de verdade.** Se o líder ou os vinculados forem removidos no modal
  "Motoristas na Nota", eles saem do valor devido.

## Pendente de decisão do usuário

- **Motorista em várias notas.** O total do período de um motorista continua sendo somado em
  cada nota em que ele aparece (como na versão anterior).

## Testes

```
node testes/regras.test.js        # regras de vínculo e valor devido
node testes/fluxo-app.smoke.js    # carrega os scripts num DOM falso com Graph simulado
```
