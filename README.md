# Corvos

Suíte de auditoria fiscal — conciliação entre XML de NF-e e SPED Fiscal.

> **Status:** V1.0 — Fundação operacional completa (sprint AUDUMBLA).
> Hoje o motor compara apenas `VALOR_TOTAL` por Chave de Acesso.
> A expansão para 14 campos (sprint HEIMDALL) está em desenvolvimento.

## O que é

O Corvos verifica a consistência entre documentos fiscais recebidos em XML
e informações escrituradas no SPED Fiscal. Ele **não substitui** a conferência
fiscal — ele **sinaliza** divergências para revisão humana.

**Regra de linguagem, válida em todo o sistema:**
> Nunca "SPED errado" ou "XML errado". Sempre "divergência identificada —
> revisão necessária".

## Como funciona

1. Coloque os XMLs em `01_XML_Fornecedores/`
2. Coloque o SPED Fiscal (`.txt`) em `02_Dados_ERP/`
3. Duplo-clique em `Munin.bat`
4. Informe CNPJ auditado, tipo de operação e período
5. O painel abre com o resumo da auditoria

### Modo DEMO

Para testar sem dado real:

1. Duplo-clique em `Hugin.bat` → aperte Enter (usa CNPJ padrão)
2. Duplo-clique em `Munin.bat` → Enter no prompt de demo
3. O painel abre com a massa de teste

## Requisitos

- Windows 10 ou 11
- Microsoft Excel 2016+
- PowerShell 5.1+

## Estrutura

| Arquivo | Função |
|:---|:---|
| `Munin.bat` | Orquestrador principal |
| `Hugin.bat` | Launcher do gerador de massa de teste |
| `Huginn_Gerador.ps1` | Gerador de XML + SPED fictícios |
| `MUNIN_template.xlsx` | Template mestre do workbook |

## Sprints

| # | Nome | Escopo | Status |
|:---|:---|:---|:---|
| 1 | **AUDUMBLA** | Fundação operacional | ✅ Concluído |
| 2 | **HEIMDALL** | 14 campos de cabeçalho | ⏳ Planejado |
| 3 | **ASGARD** | Padrões e integridade do DANFE | ⏳ Planejado |
| 4 | **NIFLHEIM** | Anomalias e status SEFAZ | ⏳ Planejado |
| 5 | **MIDGARD** | Camada de itens e fuzzy matching | ⏳ Planejado |

## Documentação

- [`docs/MARCO_0.md`](docs/MARCO_0.md) — contexto e roadmap
- [`docs/DOCUMENTACAO_TECNICA.md`](docs/DOCUMENTACAO_TECNICA.md) — arquitetura completa

## Privacidade

O Corvos roda **100% local**. Nenhum dado sai da máquina.
`config.txt`, `MUNIN.xlsx`, `.senha.txt` e as pastas de dados não são versionados.

## Licença

MIT — veja [`LICENSE`](LICENSE).
