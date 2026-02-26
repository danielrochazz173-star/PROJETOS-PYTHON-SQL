# PROJETOS-PYTHON

Colecao de projetos Python focados em analise comercial, automacao de processos e integracao com Oracle/Google Sheets/BigQuery/Salesforce Marketing Cloud.  
Organizei este repositorio para leitura de recrutadores e avaliacao tecnica rapida.

## Visao geral

- Total de scripts mapeados: **42**
- Foco principal: **dados comerciais (vendas, margem, comissao, compras x venda)**
- Tecnologias mais usadas: **Python, Oracle DB, BigQuery, Salesforce Marketing Cloud, Pandas, OpenPyXL, Tkinter, Streamlit**
- Tipos de entrega: **Excel, PDF, dashboard web, automacoes e monitoramento**

## Projeto destaque (Cloud / BigQuery)

- [`projeto_bigquery_risco_comercial/README.md`](projeto_bigquery_risco_comercial/README.md)  
  **RiskPulse BQ**: pipeline Python + BigQuery + Salesforce Marketing Cloud com carga incremental (`MERGE`), data quality, alertas de quebra de margem (z-score), particionamento/clustering, ativacao em Data Extension e auditoria de execucao.

## Organizacao por pastas (documentacao)

Os scripts continuam na raiz para manter compatibilidade atual.  
Dividi a documentacao por pasta em `docs/projetos/`:

- [`docs/projetos/README.md`](docs/projetos/README.md)
- [`docs/projetos/01-analises-produtos-vendas/README.md`](docs/projetos/01-analises-produtos-vendas/README.md)
- [`docs/projetos/02-vendedores-comissoes/README.md`](docs/projetos/02-vendedores-comissoes/README.md)
- [`docs/projetos/03-compras-departamentos/README.md`](docs/projetos/03-compras-departamentos/README.md)
- [`docs/projetos/04-dashboards-bi/README.md`](docs/projetos/04-dashboards-bi/README.md)
- [`docs/projetos/05-integracoes-automacoes/README.md`](docs/projetos/05-integracoes-automacoes/README.md)
- [`docs/projetos/06-monitoramento-utilitarios/README.md`](docs/projetos/06-monitoramento-utilitarios/README.md)

## Competencias tecnicas que demonstro

- Modelagem de consultas SQL em ambiente Oracle para regras comerciais reais
- Transformacao e consolidacao de dados com Pandas
- Geracao de relatorios executivos em Excel (formatacao e formulas) e PDF
- Interfaces desktop com Tkinter para uso operacional
- Dashboards analiticos com Streamlit e Flask
- Integracao com Google Sheets para fluxos de comissao e negociacao
- Automacao orientada a produtividade de times comerciais/administrativos

## Stack utilizada

- Banco e dados: `oracledb`, `pandas`, `numpy`
- Banco e dados cloud: `google-cloud-bigquery`, `google-auth`
- Planilhas/arquivos: `openpyxl`, `reportlab`, `weasyprint`
- Interfaces e apps: `tkinter`, `tkcalendar`, `streamlit`, `flask`, `plotly`
- Integracoes: `gspread`, `oauth2client`, modulo local `google_sheets_api`
- Utilitarios: `plyer`, `win10toast`, `python-dateutil`

## Execucao local (guia rapido)

1. Criar ambiente virtual:

```powershell
python -m venv .venv
.venv\Scripts\activate
```

2. Instalar dependencias principais:

```powershell
pip install pandas oracledb openpyxl reportlab weasyprint streamlit flask plotly gspread oauth2client python-dateutil tkcalendar plyer win10toast
```

3. Configurar credenciais/conexao Oracle e Google (quando aplicavel) nos scripts.

4. Executar o script desejado:

```powershell
python nome_do_script.py
```

## Destaques do meu portfolio

- Projetos com impacto direto em indicadores de negocio (faturamento, margem, positivacao, mix e comissao)
- Solucoes com interface para uso por areas nao tecnicas
- Entregas de BI e dashboards para tomada de decisao
- Automacoes de rotina com foco em ganho operacional
