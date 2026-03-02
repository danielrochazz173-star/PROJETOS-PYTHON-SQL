# 04 - Dashboards e BI

Ferramentas de visualização executiva e disponibilização de dados para análise gerencial.  
Respondem perguntas como: *Como está o faturamento hoje versus o mês passado? Como alimentar o Power BI com dados reais do ERP?*

## Scripts deste grupo

| Script | O que faz | Como executar | Saída |
|---|---|---|---|
| `dashboard_analise_vendedores.py` | Dashboard anual com comparativos de tendência por vendedor. | `streamlit run dashboard_analise_vendedores.py` | Web em `localhost:8501` |
| `dashboard_vendedores_streamlit.py` | Dashboard com comparação mes atual vs anterior, cards de KPIs e exportação Excel. | `streamlit run dashboard_vendedores_streamlit.py` | Web em `localhost:8501` |
| `dashboard_crescimento.py` | Dashboard Flask para comparativo de crescimento 2024 vs 2025. | `python dashboard_crescimento.py` | Web em `localhost:5000` |
| `gerar_dados_bi_vendas.py` | Gera tabelas fato e dimensões (modelo estrela) para consumo no Power BI. | `python gerar_dados_bi_vendas.py` | CSVs na pasta `DADOS_BI_VENDAS/` |

## Projeto complementar (pasta dedicada)

- [`projeto_bigquery_risco_comercial/`](../../projeto_bigquery_risco_comercial/README.md)  
  Pipeline cloud com BigQuery para consolidação comercial, carga incremental com `MERGE`, detecção de quebra de margem (z-score) e ativação de alertas no Salesforce Marketing Cloud.

## Tecnologias utilizadas

- `streamlit`, `flask`, `plotly`, `matplotlib`
- `oracledb`, `pandas`, `openpyxl`

## Valor para o negócio

- Leitura executiva rápida de indicadores sem abrir o ERP
- Base de dados estruturada (modelo estrela) para dashboards corporativos no Power BI
- Comparativos automáticos que apoiam reuniões de performance e planejamento comercial

