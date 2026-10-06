# MUNIN — Marco 0
### Motor de Unificação, Navegação e Inspeção de Notas
### Registro do que foi o MVP e o que vem a seguir

---

## O que é o Corvos

Suíte de auditoria fiscal construída para cruzar o XML de nota fiscal
recebido de fornecedores contra a escrituração no SPED Fiscal (EFD
ICMS/IPI), identificando divergências antes que virem risco de autuação
ou prejuízo financeiro. Não substitui a conferência fiscal/contábil — é
uma segunda camada de verificação independente, que direciona o
profissional para revisão humana quando algo não bate.

Regra de linguagem, válida em todo o sistema:
> Nunca "SPED errado" ou "XML errado". Sempre "divergência
> identificada — revisão necessária".

---

## MVP (Marco 0) — o que já existe e já funciona

**Escopo:** confronto de duas fontes — XML de fornecedores × SPED
Fiscal (EFD ICMS/IPI) — no nível de documento (não item).

**Arquitetura:**
- Modelo `MUNIN.xlsx` pré-construído (Power Query já configurado,
  abas de output protegidas contra edição manual).
- `Munin.bat` + script PowerShell gerado em tempo de execução: o
  orquestrador confere se os arquivos necessários estão presentes,
  aguarda em loop se não estiverem, grava os caminhos corretos na aba
  `CONFIGURACAO` e aciona a atualização do Excel via COM.
- `Hugin.bat` + `Huginn_Gerador.ps1`: gerador de massa de teste
  (XML + SPED fictícios, com quantidade e posição de erro
  aleatorizadas a cada execução), usado para validar a lógica sem
  depender de dado real de cliente.
- Uma cópia de `Munin.bat` + `MUNIN.xlsx` por pasta de cliente —
  isolamento total entre carteiras, sem misturar dado.
- Modelo mestre centralizado (`MUNIN_template.xlsx`), copiado
  automaticamente para clientes novos na primeira execução.
- Arquivo `config.txt` local guarda os parâmetros da última execução
  (CNPJ auditado, tipo de operação, período, caminhos das pastas).

**O que o MVP compara:** `VALOR_TOTAL` da nota, via Join pela Chave de
Acesso (44 dígitos). Resultado classificado em: `OK`, `DIVERGÊNCIA`,
`NÃO LOCALIZADO`, `REVISAR`, `DUPLICIDADE`.

**Testado com sucesso:** massa de 80 notas geradas pelo Huginn, com 3
divergências de valor, 3 não localizadas e 1 chave divergente
propositalmente injetadas — o painel do MUNIN reproduziu os números
esperados.

---

## Robustez adicionada ao MVP

Após o MVP rodar ponta-a-ponta, foram incorporadas as seguintes
melhorias de infraestrutura — todas sem alterar a lógica de negócio
(a comparação continua sendo apenas `VALOR_TOTAL`):

**Validação de CNPJ (NT COCAD/SUARA nº 49/2024)**
- Antes: o orquestrador só checava se o CNPJ tinha 14 dígitos.
- Agora: valida os 2 dígitos verificadores com a regra oficial,
  incluindo suporte a **CNPJ alfanumérico** (12 primeiros caracteres
  em `[0-9A-Z]`, 2 últimos em `[0-9]`).
- O validador é isolado em arquivo `.ps1` temporário, evitando
  problemas de escape do `cmd` com `^`, `(`, `)`, `!`.

**Senha dinâmica para proteção das abas**
- Antes: senha hardcoded no `.bat` (`SenhaCorvos2026`), visível a
  qualquer um que abrisse o arquivo.
- Agora: a senha é lida de `.senha.txt` ou gerada automaticamente na
  primeira execução (`Corvos<random>`). O arquivo `.senha.txt` é
  ignorado pelo Git — cada máquina tem a sua.

**Re-proteção automática das abas**
- Antes: se o template tinha abas protegidas, elas voltavam
  desprotegidas após a execução.
- Agora: o script registra quais abas estavam protegidas e as
  re-protege antes do save final.

**Tratamento de erro no COM do Excel**
- Antes: um erro no meio da automação deixava `EXCEL.EXE` pendurado
  no Gerenciador de Tarefas.
- Agora: o `catch` fecha o workbook (`$wb.Close($false)`) e encerra
  o processo (`$excel.Quit()`).

**Bugfix — leitura do `config.txt`**
- Corrigido `useblock` → `usebackq` no `for /f`, que impedia a
  leitura correta do arquivo de configuração entre execuções.

**Modo DEMO (integração Huginn ↔ Munin)**
- O `Huginn_Gerador.ps1` agora escreve um marcador `.huginn_demo` ao
  gerar massa de teste, com CNPJ + período usados.
- O `Munin.bat` detecta esse marcador na inicialização e pergunta:
  *"Este é um protocolo DEMO?"* — com **Enter = Sim** (fricção
  mínima). Se sim, aplica os parâmetros do marcador automaticamente,
  sem exigir configuração manual.
- Comando `Munin.bat /CONFIG` força o modo visível (ignora demo) e
  entra no fluxo de configuração manual.

**Huginn aceita qualquer CNPJ**
- Antes: CNPJ fixo `12345678000199` — inválido pelos dígitos
  verificadores, rejeitado por consulta pública (RedeSim).
- Agora: aceita argumento direto (`Hugin.bat 11222333000181`), prompt
  interativo com validação em loop, ou **Enter = padrão**
  (`11222333000181`, DV válido, usado em homologação no Brasil).

**Template limpo para distribuição**
- Removidos do `MUNIN_template.xlsx`: CNPJ de teste inválido e
  caminhos pessoais (`C:\Users\...\...`).
- Removidas chaves duplicadas com espaço (`TIPO DE OPERAÇÃO`,
  `DATA INICIO`, `DATA FIM`) que causavam divergência silenciosa
  entre `MATCH` do Excel e leitura do Power Query.
- Células formatadas corretamente (Texto vs Número).

**Preparação para publicação**
- `.gitignore` protege `config.txt`, `MUNIN.xlsx`, `MUNIN.xlsx.bak`,
  `.senha.txt`, `.huginn_demo` e as pastas de dados
  (`01_XML_Fornecedores/`, `02_Dados_ERP/`).
- `README.md` com descrição, problema resolvido, stack, instruções.
- `LICENSE` MIT (bilíngue).
- `DEDICATORIA.md` — homenagem à Lih, fora do LICENSE para preservar
  a integridade da licença.

**Importante:** a lógica de negócio do MVP **não foi alterada**.
A comparação continua sendo apenas `VALOR_TOTAL`. Tudo o que foi
adicionado é infraestrutura em volta — entrada de dados, segurança,
tratamento de erro, testabilidade e distribuição.

---

## Decisões de escopo tomadas conscientemente

- ICMS, IPI, PIS, COFINS, CFOP, NCM, CST — fora do MVP. Motivo:
  exigem lógica de alíquota esperada por NCM, mais complexa; o MVP já
  entrega valor real só com divergência de valor total.
- Nenhum agendamento automático (rodar sozinho ao ligar o PC, ou em
  horário fixo) — escopo descartado por ser desproporcional ao
  volume real de uso (carteira de cliente pequeno/médio).
- Nenhuma nota de "criticidade" ou gravidade nesta fase — o sistema
  descreve o fato (qual campo divergiu, de que natureza), sem atribuir
  peso de gravidade sem dado real que sustente esse peso.

---

## Sprints da Linhagem Ymir

### Sprint 1 — AUDUMBLA: Fundação operacional
**Status:** ✅ Concluído

Fundação sobre a qual todos os sprints seguintes se apoiam. Sem
alterar a lógica de negócio do MVP (que continua comparando apenas
`VALOR_TOTAL`), este sprint blindou o orquestrador em cinco frentes:

- **Entrada de dados:** validação de CNPJ com DV, Huginn aceita
  qualquer CNPJ.
- **Segurança:** senha dinâmica via `.senha.txt`, re-proteção de abas.
- **Robustez:** tratamento de erro no COM, bugfix do `config.txt`.
- **Testabilidade:** modo DEMO integrado.
- **Distribuição:** template limpo, `.gitignore`, README, LICENSE.

---

### Sprint 2 — HEIMDALL: Mapeamento completo de campos de cabeçalho
**Status:** ⏳ Planejado

**Objetivo:** ampliar a comparação campo a campo (mesmo mecanismo já
validado no MVP, só que para mais variáveis), e fazer a aba de Revisão
informar **exatamente qual campo divergiu**, não apenas um motivo
genérico.

**Critério usado para escolher os campos:** o elo mais fraco de
qualquer SPED/ERP é o ponto onde existe intervenção humana manual —
campo digitado à mão carrega risco real de erro; campo calculado
automaticamente pelo próprio sistema, não.

**Campos mapeados:**
IND_OPER, CNPJ_EMIT, COD_SIT, SERIE, NUM_DOC,
DT_EMISSAO, DT_ENT_SAI, VALOR_TOTAL, VALOR_PROD,
IND_FRETE, VALOR_FRETE, VALOR_SEGURO,
VALOR_DESPESAS, VALOR_DESCONTO

(`CHAVE_NFE` fica de fora — é a chave de junção; `COD_MOD` e `IND_EMIT`
também ficam de fora — geralmente fixos ou definidos automaticamente
pelo próprio sistema.)

**Mecanismo:** nota já pareada pela Chave de Acesso (Camada 1 existente)
→ comparar cada um dos 14 campos entre XML e SPED → se divergir,
registrar qual campo específico no `MOTIVO_REVISAO`, podendo listar
mais de um por nota (ex.: *"Data de emissão divergente; Valor de frete
divergente"*).

**O que NÃO entra neste sprint:** nenhuma nota de gravidade/criticidade
(0-10) — decisão consciente de que isso só deve nascer de dado real
acumulado, não de estimativa a priori.

---

### Sprint 3 — ASGARD: Padrões, contagem e integridade do DANFE
**Status:** ⏳ Planejado

- **Casa decimal deslocada**: identificar quando `VALOR_SPED` é
  aproximadamente `VALOR_XML × 10` ou `÷ 10`.
- **Duplicidade no SPED**: agrupar por Chave de Acesso e sinalizar
  quando a mesma chave aparece mais de uma vez.
- **Integridade proposicional do DANFE**: validar que os dígitos
  23-25 (série) e 26-34 (número) embutidos na própria Chave de Acesso
  batem com as tags `<serie>` e `<nNF>` do XML.

---

### Sprint 4 — NIFLHEIM: Anomalias estruturais e status SEFAZ
**Status:** ⏳ Planejado

- **Omissão de lançamento**: XML presente na pasta sem escrituração
  correspondente no SPED.
- **Nota "fantasma"**: lançamento no SPED sem o XML correspondente na
  pasta local.
- **Status SEFAZ (`cStat`)**: identificar notas canceladas (101/135)
  ou denegadas (110, em homologação — 301/302/303 em produção) que
  continuam escrituradas como regulares (`COD_SIT = 00`).

---

### Sprint 5 — MIDGARD: Camada de itens e fuzzy matching
**Status:** ⏳ Planejado

- **Validação de NCM por item**: comparar o NCM do item (registro
  C170/0200 do SPED) contra a tag `<det><prod><NCM>` do XML.
- **Camada 2 — fuzzy matching**: para as notas que sobraram sem par
  (status `NÃO LOCALIZADO` dos dois lados), tentar reconectá-las por
  semelhança de CNPJ do emitente + valor total + data de emissão —
  recuperando casos de Chave de Acesso digitada com erro. Roda apenas
  sobre os órfãos, não sobre o volume total. Resultado desta camada:
  status `POSSÍVEL DIVERGÊNCIA — revisar correspondência`.

---

## Princípio geral, válido para todos os sprints futuros

Antes de atribuir peso ou gravidade a qualquer tipo de erro, primeiro
**descrever o fato com precisão** (qual campo, que natureza, manual ou
automático). Qualquer priorização por criticidade só deve ser
introduzida depois de dado real acumulado o suficiente para validar
empiricamente quais campos realmente divergem com mais frequência e
maior impacto — nunca por estimativa a priori sem lastro.
