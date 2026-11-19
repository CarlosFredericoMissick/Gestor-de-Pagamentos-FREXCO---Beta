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
js/divisao-recebivel.js modal "Dividir recebível" (alocação do recebível de um motorista entre notas)
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
- `alocacoesRecebivel`: divisão do recebível de um motorista entre notas, uma entrada por
  motorista e período de conciliação:
  `[{ vehicleId, valor, periodoInicio: 'AAAA-MM-DD', periodoFim, editadoEm, editadoPor }]`.
  Só as notas com valor fixado têm a entrada; a nota que recebe o restante não precisa dela.
  Em conflito (412) a mescla é por motorista e período: alocações de outros motoristas na
  mesma nota, gravadas por outro usuário, são preservadas.

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
- **Rótulos "Comissão" e "Rotas".** Na tela e nos exports, prestador líder aparece como
  "<nome> Comissão" (traz a comissão do time) e o motorista com o mesmo nome de um líder
  presente na mesma lista aparece como "<nome> Rotas". Sem ambiguidade, o motorista aparece só
  com o nome. É só exibição (`rotulosMotoristas`, `js/regras.js`); o JSON guarda o nome puro.
  Os nomes dos PDFs no "Baixar ZIP" usam o nome puro, como na versão original.
- **Motorista em várias notas: divisão do recebível.** Se um motorista com recebível no
  período está em 2 ou mais notas (todas as notas, independente do filtro de time), a linha
  da conciliação mostra o botão "Dividir recebível". No modal o usuário informa quanto vai
  para cada nota; a nota deixada em branco recebe o restante (`dividirRecebivelEntreNotas`):
  - as rotas do motorista no período são ordenadas por data e posição na planilha;
  - as notas com valor são preenchidas em ordem (emissão, extração, arquivo) até o valor; a
    rota que passa do valor é fracionada ("parcial"): uma parte fica na nota e a sobra vai
    para a nota do restante. Cada nota com valor fica com no máximo uma rota fracionada,
    exceto quando só restam partes de rotas já fracionadas para completar o valor;
  - tudo em centavos inteiros: a soma das partes é sempre o total do motorista;
  - 2 ou mais notas em branco: elas continuam com o total do motorista e a linha avisa
    "aloque as demais notas deste motorista";
  - soma maior que o total, ou nenhuma nota em branco: o modal bloqueia ao salvar;
  - a divisão vale só para o período em que foi feita; em outro período é ignorada e a linha
    mostra "alocação de outro período ignorada";
  - se o motorista sai da nota, a alocação dele é descartada.

  Exemplo: motorista com R$ 3.000 (rotas 02/10 R$ 1.200, 06/10 R$ 900, 10/10 R$ 900) e notas de
  R$ 2.000 e R$ 1.000. Alocando R$ 2.000 na primeira, ela fica com 02/10 (1.200) + 800 de 06/10
  (parcial); a segunda recebe o restante: 100 de 06/10 (parcial) + 10/10 (900) = R$ 1.000.

- **DANFSe Nacional: líquido fora da coluna fica marcado.** O líquido é lido da coluna
  "VALOR LÍQUIDO DA NFS-e" (caminho normal, sem aviso). Se essa leitura falhar (colunas
  perdidas, valor incoerente), o extrator calcula
  `Serv − Desconto incondicionado − Desconto condicionado − Retenções` (fórmula da NFS-e
  nacional/ABRASF: os dois descontos reduzem o líquido). Os descontos são lidos pela ordem dos
  rótulos da linha "VALOR DA OPERAÇÃO / SERVIÇO …", por token, ou no formato
  "Desconto Incondicionado: R$ x". Desconto ilegível é considerado zero, e o aviso diz isso.
  Qualquer líquido que não veio da coluna grava na nota `avisoExtracao: 'liquido-calculado'` e
  `avisoExtracaoDetalhe` (a conta usada) e a linha fica amarela até ser revisada.

## Tabela de notas: destaque das linhas

Derivado na hora de desenhar a linha (`estadoLinhaNota`, `js/regras.js`), não gravado.
Fundo, em ordem de prioridade (vale a primeira que se aplica):
1. **vermelho claro**: problema que impede o pagamento: CNPJ sem prestador cadastrado (coluna
   Responsável) ou PIX faltante (coluna Chave PIX). Some ao cadastrar e reprocessar.
2. **amarelo claro**: valor líquido calculado pelo extrator, a conferir (`avisoExtracao`,
   ver abaixo). O valor a pagar pode estar errado, por isso vem antes do OCR.
3. **laranja claro**: nota extraída por OCR ainda não revisada.

Indicadores que aparecem mesmo sob outra cor: ícone de alerta ao lado do R$ Líq (líquido
calculado) e borda esquerda laranja (OCR pendente). O tooltip da linha lista tudo.

A revisão é uma só ("Confirmo que revisei os dados desta nota") e limpa o amarelo e o laranja:
- notas OCR: modal "Editar Campos — Nota OCR" (como antes);
- notas com líquido calculado: o mesmo modal, como "Revisar Valores — Líquido calculado",
  mostra o motivo e permite corrigir os valores à mão;
- gravada na nota no SharePoint: `ocrRevisado`, `ocrRevisadoEm`, `ocrRevisadoPor` (nome mantido
  por compatibilidade com as notas OCR já revisadas). Desmarcar volta a destacar.

Acima da tabela: "X com problema · Y OCR pendente de revisão · Z líquido calculado a conferir"
das notas visíveis.

## Gravação de prestadores.json

As gravações passam por uma fila (`savePrestadoresSP_internal`): nunca há dois PUTs
simultâneos no mesmo arquivo. Cada gravação leva um `idGravacao`; se o PUT responde erro, o
arquivo é relido e, se tiver esta versão, a gravação é considerada feita (o SharePoint pode
gravar e responder 5xx/timeout). Os botões Salvar de líder e prestador ignoram clique repetido
enquanto gravam.

## Comportamentos conhecidos (mantidos de propósito)

- **Alocação que volta a valer.** A alocação de um motorista que deixou de estar na nota por
  mudança de cadastro (nota automática, sem edição no modal) só é ignorada no cálculo, não é
  apagada. Se o motorista voltar à nota, ela volta a valer. Remover o motorista pelo modal
  "Motoristas na Nota" apaga a alocação.
- **Rota negativa na nota alocada.** Rotas com Total Motorista negativo (ajustes) entram na
  ordem de data como as outras; podem cair na nota com valor fixado (o valor total dela
  continua exato, completado pelas rotas seguintes).
- **2 ou mais notas em branco.** Até o usuário alocar as demais notas do motorista, cada nota
  em branco soma o total do motorista; a linha avisa "aloque as demais notas deste motorista".
- **Notas DANFSe antigas não ganham o aviso.** `avisoExtracao` só é gravado na extração; notas
  extraídas antes desta versão não são reavaliadas. Para conferir, reextraia o PDF.
- **Revisão marca a nota inteira.** Corrigir o líquido e marcar "revisei" limpa o destaque;
  `avisoExtracao` continua gravado como registro de que o valor foi calculado.

## Armazenamento local (IndexedDB)

Sem SharePoint conectado, os prestadores ficam no IndexedDB `FrexcoPrestadoresDB`, versão 2
(`js/armazenamento-local.js`). A versão anterior criava um banco v1 vazio no primeiro acesso
e a partir daí o app mostrava "Erro ao inicializar sistema"; a v2 cria a store que faltar e
mantém os dados de quem já tinha a v1. A migração do banco legado `FrexcoPixDB` só lê bancos
que já existem e fecha as conexões que abre.

Teste manual: no DevTools > Application > IndexedDB, apague os bancos `FrexcoPrestadoresDB` e
`FrexcoPixDB` e recarregue (primeiro acesso: abre sem erro e cria só `FrexcoPrestadoresDB` v2).
Para simular o estado quebrado, no console de uma aba sem o app aberto em outra aba:
`indexedDB.deleteDatabase('FrexcoPrestadoresDB'); indexedDB.open('FrexcoPrestadoresDB', 1)` e
recarregue: o app abre sem erro e o banco passa para a v2 com a store `prestadores`.

## Testes

```
node testes/regras.test.js        # regras de vínculo, valor devido, divisão do recebível e rótulos
node testes/fluxo-app.smoke.js    # carrega os scripts num DOM falso com Graph simulado
node testes/extracao-danfse.test.js  # extrator DANFSe Nacional com linhas reais do pdf.js (fixtures/nfse-desconto)
```
