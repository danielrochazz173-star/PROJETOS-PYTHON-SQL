# Portfolio Python + SQL — Análise Comercial e Automação

Coleção de projetos Python desenvolvidos para resolver problemas reais de uma operação comercial B2B:
análise de vendas e margem, acompanhamento de vendedores, controle de compras x venda,
dashboards executivos e integração com Oracle ERP, Google Sheets, BigQuery e Salesforce Marketing Cloud.

> Repositório organizado para avaliação técnica rápida por recrutadores e gestores.

---

## O que cada grupo de projetos resolve

| Grupo | Problema de negócio resolvido | Entrega |
|---|---|---|
| [Análises de Produtos e Vendas](docs/projetos/01-analises-produtos-vendas/README.md) | Quais produtos vendem mais? Qual a margem real após devoluções? Quais clientes compraram? | Excel formatado, PDF, interface desktop |
| [Vendedores e Comissões](docs/projetos/02-vendedores-comissoes/README.md) | Como cada vendedor performou? Qual a comissão correta? Como negociar ajustes? | Excel, Google Sheets, dashboard Streamlit |
| [Compras e Departamentos](docs/projetos/03-compras-departamentos/README.md) | Estamos comprando mais do que vendemos? Onde há risco de margem negativa? | Excel de comparativo, relatórios gerenciais |
| [Dashboards e BI](docs/projetos/04-dashboards-bi/README.md) | Como visualizar indicadores em tempo real? Como alimentar o Power BI? | Dashboards Streamlit/Flask, modelo estrela CSV |
| [Integrações e Automações](docs/projetos/05-integracoes-automacoes/README.md) | Como reduzir retrabalho entre ERP, XML e planilhas? | Rotinas automatizadas, atualização de cadastros |
| [Monitoramento e Utilitários](docs/projetos/06-monitoramento-utilitarios/README.md) | Como ser avisado quando um pedido fica parado? | Notificações desktop em tempo real |

---

## Projeto destaque — Pipeline Cloud (BigQuery + Salesforce Marketing Cloud)

**[RiskPulse BQ](projeto_bigquery_risco_comercial/README.md)**: pipeline Python end-to-end que:

- consolida pedidos, devoluções e custos do ERP no BigQuery;
- calcula margem líquida por cliente/produto/departamento;
- detecta quebras de margem com z-score estatístico;
- faz carga incremental com `MERGE` (SCD simplificado);
- publica alertas no Salesforce Marketing Cloud para ativação de jornada CRM;
- registra auditoria de cada execução.

> Demonstra: arquitetura de dados cloud, SQL analítico, integração martech e operação rastreável.

---

## Visão geral dos números

- **42 scripts** documentados e organizados por domínio funcional
- **6 grupos temáticos** cobrindo toda a cadeia comercial (venda → margem → comissão → compra → BI → alerta)
- **3 tipos de interface**: linha de comando, desktop (Tkinter) e web (Streamlit / Flask)
- **4 sistemas integrados**: Oracle ERP, Google Sheets, BigQuery, Salesforce Marketing Cloud

---

## Stack utilizada

| Camada | Bibliotecas |
|---|---|
| Banco relacional | `oracledb`, `pandas`, `numpy` |
| Banco cloud | `google-cloud-bigquery`, `google-auth` |
| Planilhas e PDF | `openpyxl`, `reportlab`, `weasyprint` |
| Interfaces | `tkinter`, `tkcalendar`, `streamlit`, `flask`, `plotly` |
| Integrações | `gspread`, `oauth2client` |
| Utilitários | `plyer`, `win10toast`, `python-dateutil` |

---

## Como executar localmente

### 1. Criar ambiente virtual

```powershell
python -m venv .venv
.venv\Scripts\activate
```

### 2. Instalar dependências

```powershell
pip install pandas oracledb openpyxl reportlab weasyprint streamlit flask plotly gspread oauth2client python-dateutil tkcalendar plyer win10toast
```

### 3. Configurar credenciais Oracle

Copie o arquivo de exemplo e preencha com suas credenciais:

```powershell
copy .env.example .env
```

Edite `.env` com os valores reais:

```
DB_HOST=seu_servidor_oracle
DB_PORT=1521
DB_SERVICE=PROD
DB_USER=seu_usuario
DB_PASSWORD=sua_senha
```

As variáveis são lidas automaticamente pelos scripts via `os.getenv()`.

### 4. Executar o script desejado

```powershell
# Script de linha de comando
python analise_produtos_completa.py

# Dashboard web (Streamlit) — acesse http://localhost:8501
streamlit run dashboard_vendedores_streamlit.py

# Dashboard web (Flask) — acesse http://localhost:5000
python dashboard_crescimento.py
```

---

## Competências que demonstro

- Consultas SQL complexas (CTEs, MERGE, window functions) em Oracle e BigQuery
- Transformação e consolidação de dados com Pandas
- Relatórios executivos em Excel (formatação condicional, formulas) e PDF
- Interfaces desktop Tkinter para uso operacional por áreas não técnicas
- Dashboards analíticos com Streamlit e Flask + Plotly
- Integração com Google Sheets para fluxos de comissão e negociação
- Pipeline cloud com BigQuery, particionamento, clustering e ativação CRM
- Boas práticas: credenciais via variáveis de ambiente, tratamento de erros, código comentado

---

## Documentação por categoria

- [`docs/projetos/01-analises-produtos-vendas/README.md`](docs/projetos/01-analises-produtos-vendas/README.md)
- [`docs/projetos/02-vendedores-comissoes/README.md`](docs/projetos/02-vendedores-comissoes/README.md)
- [`docs/projetos/03-compras-departamentos/README.md`](docs/projetos/03-compras-departamentos/README.md)
- [`docs/projetos/04-dashboards-bi/README.md`](docs/projetos/04-dashboards-bi/README.md)
- [`docs/projetos/05-integracoes-automacoes/README.md`](docs/projetos/05-integracoes-automacoes/README.md)
- [`docs/projetos/06-monitoramento-utilitarios/README.md`](docs/projetos/06-monitoramento-utilitarios/README.md)
