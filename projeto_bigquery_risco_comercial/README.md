# Projeto: RiskPulse BQ (Python + BigQuery + Salesforce Marketing Cloud)

Projeto de portfolio com arquitetura enterprise para operacao comercial omnichannel, conectando analytics em BigQuery com ativacao de alertas no Salesforce Marketing Cloud.

## Objetivo

Construir um pipeline que:

- consolida pedidos, devolucoes e custos;
- calcula margem liquida por cliente/produto/departamento/canal;
- aplica validacoes de data quality;
- faz carga incremental com `MERGE`;
- gera alertas de quebra de margem com base estatistica (z-score);
- publica alertas em Data Extension no Salesforce Marketing Cloud para ativacao de jornada;
- registra auditoria de execucao.

## Arquivos

- `pipeline_risco_comercial_bq.py`: pipeline end-to-end.
- `requirements.txt`: dependencias.
- `.env.example`: variaveis de ambiente.
- `sql/01_tabelas_fonte_referencia.sql`: DDL de referencia das tabelas fonte ficticias.

## Arquitetura (alto nivel)

1. `raw_erp.pedidos_cab` + `raw_erp.pedidos_itens` + `raw_erp.devolucoes_itens` + `raw_erp.custos_produtos`
2. Stage particionada: `analytics_comercial.stg_risco_comercial_*`
3. Fato incremental (SCD simplificado via `MERGE`): `analytics_comercial.fact_risco_margem_diario`
4. Tabela de alertas: `analytics_comercial.alertas_quebra_margem`
5. Sync para SFMC Data Extension: `DE_ALERTAS_RISCO_MARGEM`
6. Auditoria: `analytics_comercial.etl_auditoria_execucoes`

## Integracao Salesforce Marketing Cloud

- Autenticacao OAuth2 (`/v2/token`) com credenciais de Installed Package.
- Publicacao em lote na Data Extension via endpoint de `rowset`.
- Payload com chaves compostas de idempotencia (`RunId`, `DataReferencia`, `CodProduto`, `CodCliente`).
- Campos de ativacao para jornada: `AlertLevel`, `PercentualMargem`, `ZScoreMargem`, `GeneratedAtUtc`.

Esse desenho suporta fluxo real de "analytics -> decisao -> ativacao", comum em operacoes maduras de CRM e performance marketing.

## Regras de negocio implementadas

- `faturamento_liquido = faturamento_bruto - valor_devolvido`
- `margem_liquida = faturamento_liquido - custo_total`
- `percentual_margem = margem_liquida / faturamento_liquido`
- alerta quando margem do dia fica abaixo da media do periodo e com z-score negativo relevante

## Como executar

1. Instalar dependencias:

```powershell
pip install -r requirements.txt
```

2. Configurar credenciais GCP:

```powershell
$env:GOOGLE_APPLICATION_CREDENTIALS="C:\\chaves\\service-account.json"
```

3. Definir variaveis (pode copiar `.env.example`):

```powershell
$env:GCP_PROJECT_ID="portfolio-retail-ops"
$env:BQ_LOCATION="US"
$env:BQ_RAW_DATASET="raw_erp"
$env:BQ_ANALYTICS_DATASET="analytics_comercial"
$env:LOOKBACK_DAYS="90"
$env:SFMC_ENABLED="true"
$env:SFMC_AUTH_BASE_URL="https://mc123456789.auth.marketingcloudapis.com"
$env:SFMC_REST_BASE_URL="https://mc123456789.rest.marketingcloudapis.com"
$env:SFMC_CLIENT_ID="SEU_CLIENT_ID"
$env:SFMC_CLIENT_SECRET="SEU_CLIENT_SECRET"
$env:SFMC_ACCOUNT_ID="SEU_MID"
$env:SFMC_DATA_EXTENSION_KEY="DE_ALERTAS_RISCO_MARGEM"
```

4. Rodar em validacao:

```powershell
python pipeline_risco_comercial_bq.py --dry-run
```

5. Rodar carga real:

```powershell
python pipeline_risco_comercial_bq.py
```

## Pontos que valorizam portfolio

- Arquitetura de ponta a ponta: ingestao, modelagem, deteccao de risco e ativacao CRM
- Integracao cloud data stack + martech stack (BigQuery + Salesforce Marketing Cloud)
- SQL analitico com CTEs, particionamento, clustering e `MERGE`
- Operacao rastreavel com auditoria de execucao e contagem de registros enviados para SFMC
