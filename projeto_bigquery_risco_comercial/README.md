# Projeto: RiskPulse BQ (Python + BigQuery)

Mini-projeto de portfolio com arquitetura de pipeline analitico em BigQuery, modelado para parecer um cenario real de operacao comercial.

## Objetivo

Construir um pipeline que:

- consolida pedidos, devolucoes e custos;
- calcula margem liquida por cliente/produto/departamento/canal;
- aplica validacoes de data quality;
- faz carga incremental com `MERGE`;
- gera alertas de quebra de margem com base estatistica (z-score);
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
5. Auditoria: `analytics_comercial.etl_auditoria_execucoes`

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

- Orquestracao com responsabilidades claras (stage, DQ, merge, alertas, auditoria)
- SQL analitico com CTEs, particionamento, clustering e `MERGE`
- Projeto pronto para escalar para Cloud Scheduler + Composer/Workflows
