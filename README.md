# Apuracao Simples Nacional

> Projeto de portfólio de **Victória Pedrosa** (Automação, Processos e Dados). Automação desenvolvida para um escritório de contabilidade; **esta é uma versão com dados fictícios** — nomes, CNPJs, e-mails e IDs internos foram substituídos.

## Problema de negócio
A apuração do Simples envolve RBT12, alíquotas efetivas, antecipação do DAS e conferência de faturamento — muitos passos manuais.

## Antes x depois
| | Antes | Depois |
|---|---|---|
| Como é feito | Cálculos em planilha e conferência manual com a Domínio e o Drive. | Scripts importam faturamento, calculam RBT12 e alíquotas efetivas (LC 123/2006), conferem com a Domínio e varrem o Drive atrás de DAS e extratos. |

## Ganho
- Apuração mais rápida e com conferência automática.

## Tecnologias
Google Apps Script, Google Drive, Google Sheets, Web App (HtmlService)

## Arquivos
- `AliquotaEfetivaDAS_ISS.gs`
- `AntecipacaoDAS.gs`
- `CalculoRBT12.gs`
- `SistemaFiscalIntegrado_V30.gs`

## Como usar
Crie um projeto no Google Apps Script, copie os arquivos `.gs`/`.html` e configure as Propriedades do script indicadas no código.

## Autora
Victória Pedrosa — Product Owner do Time de IA, automação de processos contábeis e fiscais.
