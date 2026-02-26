from __future__ import annotations

import argparse
import json
import logging
import os
import uuid
from dataclasses import dataclass
from datetime import datetime, timedelta, timezone
from typing import Any

from google.cloud import bigquery
import requests

PIPELINE_NAME = "RiskPulse BQ"
PIPELINE_VERSION = "1.0.1"


@dataclass
class PipelineConfig:
    project_id: str
    location: str
    raw_dataset: str
    analytics_dataset: str
    fact_table: str
    alert_table: str
    audit_table: str
    lookback_days: int
    watermark_ts: datetime
    sfmc_enabled: bool
    sfmc_auth_base_url: str
    sfmc_rest_base_url: str
    sfmc_client_id: str
    sfmc_client_secret: str
    sfmc_account_id: str
    sfmc_data_extension_key: str
    sfmc_batch_size: int

    @classmethod
    def from_env(cls) -> "PipelineConfig":
        project_id = os.getenv("GCP_PROJECT_ID", "portfolio-retail-ops")
        location = os.getenv("BQ_LOCATION", "US")
        raw_dataset = os.getenv("BQ_RAW_DATASET", "raw_erp")
        analytics_dataset = os.getenv("BQ_ANALYTICS_DATASET", "analytics_comercial")
        fact_table = os.getenv("BQ_FACT_TABLE", "fact_risco_margem_diario")
        alert_table = os.getenv("BQ_ALERT_TABLE", "alertas_quebra_margem")
        audit_table = os.getenv("BQ_AUDIT_TABLE", "etl_auditoria_execucoes")
        lookback_days = int(os.getenv("LOOKBACK_DAYS", "90"))
        sfmc_enabled = os.getenv("SFMC_ENABLED", "false").lower() in {
            "1",
            "true",
            "yes",
            "on",
        }
        sfmc_auth_base_url = os.getenv(
            "SFMC_AUTH_BASE_URL", "https://mc123456789.auth.marketingcloudapis.com"
        )
        sfmc_rest_base_url = os.getenv(
            "SFMC_REST_BASE_URL", "https://mc123456789.rest.marketingcloudapis.com"
        )
        sfmc_client_id = os.getenv("SFMC_CLIENT_ID", "")
        sfmc_client_secret = os.getenv("SFMC_CLIENT_SECRET", "")
        sfmc_account_id = os.getenv("SFMC_ACCOUNT_ID", "")
        sfmc_data_extension_key = os.getenv(
            "SFMC_DATA_EXTENSION_KEY", "DE_ALERTAS_RISCO_MARGEM"
        )
        sfmc_batch_size = int(os.getenv("SFMC_BATCH_SIZE", "500"))

        watermark_str = os.getenv("WATERMARK_TS_UTC")
        if watermark_str:
            watermark_ts = datetime.fromisoformat(watermark_str.replace("Z", "+00:00"))
        else:
            watermark_ts = datetime.now(timezone.utc) - timedelta(days=lookback_days)

        return cls(
            project_id=project_id,
            location=location,
            raw_dataset=raw_dataset,
            analytics_dataset=analytics_dataset,
            fact_table=fact_table,
            alert_table=alert_table,
            audit_table=audit_table,
            lookback_days=lookback_days,
            watermark_ts=watermark_ts,
            sfmc_enabled=sfmc_enabled,
            sfmc_auth_base_url=sfmc_auth_base_url,
            sfmc_rest_base_url=sfmc_rest_base_url,
            sfmc_client_id=sfmc_client_id,
            sfmc_client_secret=sfmc_client_secret,
            sfmc_account_id=sfmc_account_id,
            sfmc_data_extension_key=sfmc_data_extension_key,
            sfmc_batch_size=sfmc_batch_size,
        )


class BigQueryRiscoComercialPipeline:
    def __init__(self, cfg: PipelineConfig, dry_run: bool = False) -> None:
        self.cfg = cfg
        self.dry_run = dry_run
        self.client = bigquery.Client(project=cfg.project_id, location=cfg.location)
        self.run_id = uuid.uuid4().hex[:12]
        self.started_at = datetime.now(timezone.utc)

    def run(self) -> dict[str, Any]:
        self._ensure_objects()

        if self.dry_run:
            self._run_query(self._preview_stage_sql())
            return {
                "run_id": self.run_id,
                "status": "DRY_RUN_OK",
                "started_at_utc": self.started_at.isoformat(),
                "finished_at_utc": datetime.now(timezone.utc).isoformat(),
                "staged_rows": 0,
                "merged_rows": 0,
                "alerts_rows": 0,
                "sfmc_rows_sent": 0,
                "dq": {"total_rows": 0, "null_breaks": 0, "duplicate_rows": 0, "critical_errors": 0},
            }

        staging_table = (
            f"{self.cfg.project_id}.{self.cfg.analytics_dataset}."
            f"stg_risco_comercial_{self.started_at.strftime('%Y%m%d_%H%M%S')}"
        )

        staged_rows = self._build_stage_table(staging_table)
        dq = self._validate_stage(staging_table)

        if dq["critical_errors"] > 0:
            self._write_audit(
                status="FAILED_DQ",
                staged_rows=staged_rows,
                merged_rows=0,
                alerts_rows=0,
                sfmc_rows_sent=0,
                dq_summary=dq,
            )
            raise RuntimeError(f"Data quality failed: {dq}")

        merged_rows = self._merge_fact(staging_table)
        alerts_rows = self._refresh_alerts()
        sfmc_rows_sent = self._sync_alerts_to_sfmc()

        self._run_query(f"DROP TABLE `{staging_table}`")
        self._write_audit(
            status="SUCCESS",
            staged_rows=staged_rows,
            merged_rows=merged_rows,
            alerts_rows=alerts_rows,
            sfmc_rows_sent=sfmc_rows_sent,
            dq_summary=dq,
        )

        finished_at = datetime.now(timezone.utc)
        return {
            "run_id": self.run_id,
            "status": "SUCCESS",
            "started_at_utc": self.started_at.isoformat(),
            "finished_at_utc": finished_at.isoformat(),
            "staged_rows": staged_rows,
            "merged_rows": merged_rows,
            "alerts_rows": alerts_rows,
            "sfmc_rows_sent": sfmc_rows_sent,
            "dq": dq,
        }

    def _ensure_objects(self) -> None:
        dataset_ddl = f"""
        CREATE SCHEMA IF NOT EXISTS `{self.cfg.project_id}.{self.cfg.analytics_dataset}`
        OPTIONS (location="{self.cfg.location}")
        """
        self._run_query(dataset_ddl)

        fact_ddl = f"""
        CREATE TABLE IF NOT EXISTS `{self.cfg.project_id}.{self.cfg.analytics_dataset}.{self.cfg.fact_table}` (
          data_referencia DATE,
          canal STRING,
          cod_departamento INT64,
          cod_produto INT64,
          cod_cliente INT64,
          qtd_vendida NUMERIC,
          faturamento_bruto NUMERIC,
          valor_devolvido NUMERIC,
          faturamento_liquido NUMERIC,
          custo_total NUMERIC,
          margem_liquida NUMERIC,
          percentual_margem NUMERIC,
          ticket_medio NUMERIC,
          updated_at TIMESTAMP
        )
        PARTITION BY data_referencia
        CLUSTER BY canal, cod_departamento, cod_produto
        """
        self._run_query(fact_ddl)

        alert_ddl = f"""
        CREATE TABLE IF NOT EXISTS `{self.cfg.project_id}.{self.cfg.analytics_dataset}.{self.cfg.alert_table}` (
          data_referencia DATE,
          canal STRING,
          cod_departamento INT64,
          cod_produto INT64,
          cod_cliente INT64,
          faturamento_liquido NUMERIC,
          margem_liquida NUMERIC,
          percentual_margem NUMERIC,
          media_margem_30d NUMERIC,
          desvio_margem_30d NUMERIC,
          zscore_margem FLOAT64,
          alert_level STRING,
          generated_at TIMESTAMP
        )
        PARTITION BY data_referencia
        CLUSTER BY alert_level, cod_departamento
        """
        self._run_query(alert_ddl)

        audit_ddl = f"""
        CREATE TABLE IF NOT EXISTS `{self.cfg.project_id}.{self.cfg.analytics_dataset}.{self.cfg.audit_table}` (
          run_id STRING,
          status STRING,
          started_at TIMESTAMP,
          finished_at TIMESTAMP,
          staged_rows INT64,
          merged_rows INT64,
          alerts_rows INT64,
          sfmc_rows_sent INT64,
          dq_summary_json STRING
        )
        PARTITION BY DATE(started_at)
        """
        self._run_query(audit_ddl)

    def _build_stage_table(self, staging_table: str) -> int:
        sql = self._materialize_stage_sql(staging_table)
        self._run_query(
            sql,
            params=[
                bigquery.ScalarQueryParameter(
                    "watermark_ts", "TIMESTAMP", self.cfg.watermark_ts
                )
            ],
        )
        return self._count_rows(staging_table)

    def _preview_stage_sql(self) -> str:
        return self._materialize_stage_sql(
            f"{self.cfg.project_id}.{self.cfg.analytics_dataset}.stg_risco_comercial_preview"
        )

    def _materialize_stage_sql(self, staging_table: str) -> str:
        return f"""
        CREATE OR REPLACE TABLE `{staging_table}`
        PARTITION BY data_referencia
        CLUSTER BY canal, cod_departamento, cod_produto
        AS
        WITH pedidos AS (
          SELECT
            DATE(pc.data_faturamento) AS data_referencia,
            pc.canal_venda AS canal,
            pi.cod_departamento,
            pi.cod_produto,
            pc.cod_cliente,
            SUM(pi.qtd) AS qtd_vendida,
            SUM(pi.valor_total) AS faturamento_bruto
          FROM `{self.cfg.project_id}.{self.cfg.raw_dataset}.pedidos_cab` pc
          JOIN `{self.cfg.project_id}.{self.cfg.raw_dataset}.pedidos_itens` pi
            ON pi.id_pedido = pc.id_pedido
          WHERE pc.status_pedido = 'FATURADO'
            AND pc.data_faturamento >= @watermark_ts
          GROUP BY 1,2,3,4,5
        ),
        devolucoes AS (
          SELECT
            DATE(di.data_devolucao) AS data_referencia,
            di.cod_produto,
            di.cod_cliente,
            SUM(di.valor_devolvido) AS valor_devolvido
          FROM `{self.cfg.project_id}.{self.cfg.raw_dataset}.devolucoes_itens` di
          WHERE di.data_devolucao >= @watermark_ts
          GROUP BY 1,2,3
        ),
        custos AS (
          SELECT
            DATE(cp.data_custo) AS data_referencia,
            cp.cod_produto,
            AVG(cp.custo_medio) AS custo_medio
          FROM `{self.cfg.project_id}.{self.cfg.raw_dataset}.custos_produtos` cp
          WHERE cp.data_custo >= @watermark_ts
          GROUP BY 1,2
        ),
        consolidado AS (
          SELECT
            p.data_referencia,
            p.canal,
            p.cod_departamento,
            p.cod_produto,
            p.cod_cliente,
            p.qtd_vendida,
            p.faturamento_bruto,
            COALESCE(d.valor_devolvido, 0) AS valor_devolvido,
            (p.faturamento_bruto - COALESCE(d.valor_devolvido, 0)) AS faturamento_liquido,
            (p.qtd_vendida * COALESCE(c.custo_medio, 0)) AS custo_total
          FROM pedidos p
          LEFT JOIN devolucoes d
            ON d.data_referencia = p.data_referencia
            AND d.cod_produto = p.cod_produto
            AND d.cod_cliente = p.cod_cliente
          LEFT JOIN custos c
            ON c.data_referencia = p.data_referencia
            AND c.cod_produto = p.cod_produto
        )
        SELECT
          data_referencia,
          canal,
          cod_departamento,
          cod_produto,
          cod_cliente,
          qtd_vendida,
          faturamento_bruto,
          valor_devolvido,
          faturamento_liquido,
          custo_total,
          (faturamento_liquido - custo_total) AS margem_liquida,
          SAFE_DIVIDE((faturamento_liquido - custo_total), NULLIF(faturamento_liquido, 0)) AS percentual_margem,
          SAFE_DIVIDE(faturamento_liquido, NULLIF(qtd_vendida, 0)) AS ticket_medio,
          CURRENT_TIMESTAMP() AS updated_at
        FROM consolidado
        """

    def _validate_stage(self, staging_table: str) -> dict[str, int]:
        sql = f"""
        SELECT
          COUNT(*) AS total_rows,
          COUNTIF(data_referencia IS NULL OR canal IS NULL OR cod_produto IS NULL) AS null_breaks,
          COUNT(*) - COUNT(DISTINCT CONCAT(
            CAST(data_referencia AS STRING), '|',
            canal, '|',
            CAST(cod_departamento AS STRING), '|',
            CAST(cod_produto AS STRING), '|',
            CAST(cod_cliente AS STRING)
          )) AS duplicate_rows
        FROM `{staging_table}`
        """
        row = self._run_query(sql).result().to_dataframe().iloc[0]
        summary = {
            "total_rows": int(row["total_rows"]),
            "null_breaks": int(row["null_breaks"]),
            "duplicate_rows": int(row["duplicate_rows"]),
        }
        summary["critical_errors"] = summary["null_breaks"] + max(
            summary["duplicate_rows"], 0
        )
        return summary

    def _merge_fact(self, staging_table: str) -> int:
        target = f"{self.cfg.project_id}.{self.cfg.analytics_dataset}.{self.cfg.fact_table}"
        sql = f"""
        MERGE `{target}` T
        USING `{staging_table}` S
        ON T.data_referencia = S.data_referencia
          AND T.canal = S.canal
          AND T.cod_departamento = S.cod_departamento
          AND T.cod_produto = S.cod_produto
          AND T.cod_cliente = S.cod_cliente
        WHEN MATCHED THEN
          UPDATE SET
            qtd_vendida = S.qtd_vendida,
            faturamento_bruto = S.faturamento_bruto,
            valor_devolvido = S.valor_devolvido,
            faturamento_liquido = S.faturamento_liquido,
            custo_total = S.custo_total,
            margem_liquida = S.margem_liquida,
            percentual_margem = S.percentual_margem,
            ticket_medio = S.ticket_medio,
            updated_at = CURRENT_TIMESTAMP()
        WHEN NOT MATCHED THEN
          INSERT (
            data_referencia, canal, cod_departamento, cod_produto, cod_cliente,
            qtd_vendida, faturamento_bruto, valor_devolvido, faturamento_liquido,
            custo_total, margem_liquida, percentual_margem, ticket_medio, updated_at
          )
          VALUES (
            S.data_referencia, S.canal, S.cod_departamento, S.cod_produto, S.cod_cliente,
            S.qtd_vendida, S.faturamento_bruto, S.valor_devolvido, S.faturamento_liquido,
            S.custo_total, S.margem_liquida, S.percentual_margem, S.ticket_medio, CURRENT_TIMESTAMP()
          )
        """
        self._run_query(sql)
        return self._count_rows(target)

    def _refresh_alerts(self) -> int:
        target = f"{self.cfg.project_id}.{self.cfg.analytics_dataset}.{self.cfg.alert_table}"
        source = f"{self.cfg.project_id}.{self.cfg.analytics_dataset}.{self.cfg.fact_table}"
        sql = f"""
        CREATE OR REPLACE TABLE `{target}`
        PARTITION BY data_referencia
        CLUSTER BY alert_level, cod_departamento
        AS
        WITH base AS (
          SELECT *
          FROM `{source}`
          WHERE data_referencia >= DATE_SUB(CURRENT_DATE(), INTERVAL 120 DAY)
        ),
        stats AS (
          SELECT
            cod_departamento,
            cod_produto,
            AVG(percentual_margem) AS media_margem_30d,
            STDDEV_POP(percentual_margem) AS desvio_margem_30d
          FROM base
          WHERE data_referencia >= DATE_SUB(CURRENT_DATE(), INTERVAL 30 DAY)
          GROUP BY 1,2
        )
        SELECT
          b.data_referencia,
          b.canal,
          b.cod_departamento,
          b.cod_produto,
          b.cod_cliente,
          b.faturamento_liquido,
          b.margem_liquida,
          b.percentual_margem,
          s.media_margem_30d,
          s.desvio_margem_30d,
          SAFE_DIVIDE((b.percentual_margem - s.media_margem_30d), NULLIF(s.desvio_margem_30d, 0)) AS zscore_margem,
          CASE
            WHEN b.percentual_margem < 0 OR SAFE_DIVIDE((b.percentual_margem - s.media_margem_30d), NULLIF(s.desvio_margem_30d, 0)) <= -3 THEN 'CRITICAL'
            WHEN SAFE_DIVIDE((b.percentual_margem - s.media_margem_30d), NULLIF(s.desvio_margem_30d, 0)) <= -2 THEN 'HIGH'
            ELSE 'MEDIUM'
          END AS alert_level,
          CURRENT_TIMESTAMP() AS generated_at
        FROM base b
        JOIN stats s
          ON s.cod_departamento = b.cod_departamento
          AND s.cod_produto = b.cod_produto
        WHERE b.percentual_margem < s.media_margem_30d
          AND s.desvio_margem_30d IS NOT NULL
        """
        self._run_query(sql)
        return self._count_rows(target)

    def _write_audit(
        self,
        status: str,
        staged_rows: int,
        merged_rows: int,
        alerts_rows: int,
        sfmc_rows_sent: int,
        dq_summary: dict[str, int],
    ) -> None:
        target = f"{self.cfg.project_id}.{self.cfg.analytics_dataset}.{self.cfg.audit_table}"
        sql = f"""
        INSERT INTO `{target}` (
          run_id, status, started_at, finished_at, staged_rows, merged_rows, alerts_rows, sfmc_rows_sent, dq_summary_json
        )
        VALUES (
          @run_id, @status, @started_at, CURRENT_TIMESTAMP(), @staged_rows, @merged_rows, @alerts_rows, @sfmc_rows_sent, @dq_summary_json
        )
        """
        self._run_query(
            sql,
            params=[
                bigquery.ScalarQueryParameter("run_id", "STRING", self.run_id),
                bigquery.ScalarQueryParameter("status", "STRING", status),
                bigquery.ScalarQueryParameter("started_at", "TIMESTAMP", self.started_at),
                bigquery.ScalarQueryParameter("staged_rows", "INT64", staged_rows),
                bigquery.ScalarQueryParameter("merged_rows", "INT64", merged_rows),
                bigquery.ScalarQueryParameter("alerts_rows", "INT64", alerts_rows),
                bigquery.ScalarQueryParameter("sfmc_rows_sent", "INT64", sfmc_rows_sent),
                bigquery.ScalarQueryParameter(
                    "dq_summary_json", "STRING", json.dumps(dq_summary, ensure_ascii=True)
                ),
            ],
        )

    def _sync_alerts_to_sfmc(self) -> int:
        if not self.cfg.sfmc_enabled:
            logging.info("SFMC desabilitado; sincronizacao de alertas ignorada.")
            return 0

        required = [
            self.cfg.sfmc_client_id,
            self.cfg.sfmc_client_secret,
            self.cfg.sfmc_account_id,
            self.cfg.sfmc_data_extension_key,
        ]
        if any(not value for value in required):
            raise ValueError(
                "SFMC habilitado, mas variaveis obrigatorias nao foram preenchidas."
            )

        rows = self._fetch_sfmc_payload()
        if not rows:
            logging.info("Sem alertas para sincronizar com SFMC.")
            return 0

        token = self._sfmc_get_token()
        endpoint = (
            f"{self.cfg.sfmc_rest_base_url}/hub/v1/dataevents/key:"
            f"{self.cfg.sfmc_data_extension_key}/rowset"
        )
        headers = {
            "Authorization": f"Bearer {token}",
            "Content-Type": "application/json",
        }

        sent = 0
        for i in range(0, len(rows), self.cfg.sfmc_batch_size):
            batch = rows[i : i + self.cfg.sfmc_batch_size]
            response = requests.post(
                endpoint,
                headers=headers,
                json=batch,
                timeout=30,
            )
            if response.status_code >= 300:
                raise RuntimeError(
                    f"Falha no envio SFMC ({response.status_code}): {response.text[:500]}"
                )
            sent += len(batch)
        return sent

    def _sfmc_get_token(self) -> str:
        url = f"{self.cfg.sfmc_auth_base_url}/v2/token"
        payload = {
            "grant_type": "client_credentials",
            "client_id": self.cfg.sfmc_client_id,
            "client_secret": self.cfg.sfmc_client_secret,
            "account_id": self.cfg.sfmc_account_id,
        }
        response = requests.post(url, json=payload, timeout=30)
        if response.status_code >= 300:
            raise RuntimeError(
                f"Nao foi possivel autenticar no SFMC ({response.status_code}): {response.text[:500]}"
            )
        token = response.json().get("access_token")
        if not token:
            raise RuntimeError("SFMC retornou resposta sem access_token.")
        return token

    def _fetch_sfmc_payload(self) -> list[dict[str, Any]]:
        table = f"{self.cfg.project_id}.{self.cfg.analytics_dataset}.{self.cfg.alert_table}"
        sql = f"""
        SELECT
          CAST(data_referencia AS STRING) AS data_referencia,
          canal,
          CAST(cod_departamento AS STRING) AS cod_departamento,
          CAST(cod_produto AS STRING) AS cod_produto,
          CAST(cod_cliente AS STRING) AS cod_cliente,
          CAST(faturamento_liquido AS STRING) AS faturamento_liquido,
          CAST(margem_liquida AS STRING) AS margem_liquida,
          CAST(percentual_margem AS STRING) AS percentual_margem,
          CAST(zscore_margem AS STRING) AS zscore_margem,
          alert_level,
          FORMAT_TIMESTAMP('%Y-%m-%dT%H:%M:%SZ', generated_at) AS generated_at
        FROM `{table}`
        WHERE data_referencia >= DATE_SUB(CURRENT_DATE(), INTERVAL 7 DAY)
        """
        df = self._run_query(sql).result().to_dataframe()
        payload: list[dict[str, Any]] = []
        for _, row in df.iterrows():
            payload.append(
                {
                    "keys": {
                        "RunId": self.run_id,
                        "DataReferencia": row["data_referencia"],
                        "CodProduto": row["cod_produto"],
                        "CodCliente": row["cod_cliente"],
                    },
                    "values": {
                        "Canal": row["canal"],
                        "CodDepartamento": row["cod_departamento"],
                        "FaturamentoLiquido": row["faturamento_liquido"],
                        "MargemLiquida": row["margem_liquida"],
                        "PercentualMargem": row["percentual_margem"],
                        "ZScoreMargem": row["zscore_margem"],
                        "AlertLevel": row["alert_level"],
                        "GeneratedAtUtc": row["generated_at"],
                    },
                }
            )
        return payload

    def _count_rows(self, table_fqn: str) -> int:
        sql = f"SELECT COUNT(*) AS total FROM `{table_fqn}`"
        row = self._run_query(sql).result().to_dataframe().iloc[0]
        return int(row["total"])

    def _run_query(
        self, sql: str, params: list[bigquery.ScalarQueryParameter] | None = None
    ) -> bigquery.QueryJob:
        job_config = bigquery.QueryJobConfig(
            use_legacy_sql=False,
            query_parameters=params or [],
            dry_run=self.dry_run,
            labels={
                "pipeline": "risco_comercial",
                "camada": "analytics",
                "run_id": self.run_id,
            },
        )
        job = self.client.query(sql, job_config=job_config)
        if not self.dry_run:
            job.result()
        return job


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(
        description="Pipeline BigQuery de risco comercial (margem e alertas)."
    )
    parser.add_argument(
        "--dry-run",
        action="store_true",
        help="Valida sintaxe e custos estimados sem executar alteracoes.",
    )
    return parser.parse_args()


def main() -> None:
    logging.basicConfig(level=logging.INFO, format="%(asctime)s | %(levelname)s | %(message)s")
    logging.info("Inicializando %s v%s", PIPELINE_NAME, PIPELINE_VERSION)
    args = parse_args()
    cfg = PipelineConfig.from_env()
    pipeline = BigQueryRiscoComercialPipeline(cfg, dry_run=args.dry_run)
    result = pipeline.run()
    logging.info("Pipeline finalizada: %s", result)


if __name__ == "__main__":
    main()
