# 02 - Vendedores e Comissões

Ferramentas para acompanhar desempenho individual de vendedores, calcular comissões com regras comerciais reais e apoiar negociações de fechamento.  
Respondem perguntas como: *Qual vendedor bateu a meta? Qual a comissão correta de cada um? Como registrar uma negociação de ajuste?*

## Scripts deste grupo

| Script | O que faz | Saída |
|---|---|---|
| `analise_vendedores_excel.py` | Extrai e consolida indicadores de vendedores (faturamento, margem, mix, positivação). | Excel |
| `analise_vendedores_excel_procv.py` | Versão com lógica adicional de cruzamento de dados para análise comparativa. | Excel |
| `analise_vendedores_gui.py` | Interface desktop para analisar performance de qualquer vendedor em qualquer período. | Excel + PDF |
| `analise_vendedores_outubro.py` | Script pontual para leitura rápida de desempenho do mês de outubro. | Console / Excel |
| `relatório_vendedores_excel.py` | Relatório estruturado com todos os indicadores comerciais por vendedor. | Excel formatado |
| `comissão_dados_oracle.py` | Extrai a base de cálculo de comissão diretamente do Oracle. | DataFrame / CSV |
| `atualizar_comissão_sheets.py` | Sincroniza os dados de comissão calculados com a planilha Google Sheets da equipe. | Google Sheets |
| `sistema_negociações_comissão.py` | Consolida negociações de ajuste de comissão e gera visão final do comissionamento. | Excel |

## Tecnologias utilizadas

- `oracledb`, `pandas`, `openpyxl`
- `tkinter`, `gspread`, `oauth2client`

## Valor para o negócio

- Cálculo de comissão padronizado, evitando erros e disputas no fechamento mensal
- Transparencia de performance para cada vendedor e supervisor
- Redução de retrabalho manual: dados do Oracle vao direto para o Google Sheets da equipe

