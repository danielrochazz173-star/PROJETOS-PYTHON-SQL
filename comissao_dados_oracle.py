"""
Dados de Comissão = COMISSOESDADOS a partir do Oracle
Estrutura: CODIGOVEND (CODUSUR) | NOMEVEND | CODPROD | zz_PRODDESCRICAO | Faturamento - Devolucoes (BI)
Usa a mesma lógica de faturamento líquido do gerar_dados_bi_vendas.py (vendas - devoluções proporcionais).
"""

import oracledb
import pandas as pd
from datetime import datetime
from pathlib import Path
import os

# Configuração do Oracle Instant Client
try:
    oracledb.init_oracle_client(lib_dir=r"C:\oracle\instantclient_23_9")
except Exception:
    pass

DB_HOST = os.getenv('DB_HOST', '10.0.0.10')
DB_PORT = int(os.getenv('DB_PORT', '1521'))
DB_SERVICE = os.getenv('DB_SERVICE', 'PROD')
DB_USER = os.getenv('DB_USER', 'powerbi')
DB_PASSWORD = os.getenv('DB_PASSWORD', '')


def get_db_connection():
    try:
        return oracledb.connect(
            user=DB_USER,
            password=DB_PASSWORD,
            dsn=f"{DB_HOST}:{DB_PORT}/{DB_SERVICE}",
        )
    except Exception as e:
        print(f"Erro ao conectar: {e}")
        return None


def buscar_comissao_dados(conn, data_inicio, data_fim):
    """
    Retorna dados no formato COMISSOESDADOS:
    CODIGOVEND (CODUSUR) | NOMEVEND | CODPROD | zz_PRODDESCRICAO | Faturamento - Devolucoes (BI)
    Mesma lógica do gerar_dados_bi_vendas: faturamento líquido = vendas - devoluções proporcionais por vendedor/data.
    """
    data_inicio_str = data_inicio.strftime('%d/%m/%Y')
    data_fim_str = data_fim.strftime('%d/%m/%Y')

    query = f"""
    WITH VENDAS AS (
        SELECT 
            C.CODUSUR AS codigovend,
            I.CODPROD AS codprod,
            TO_NUMBER(TO_CHAR(C.DATA, 'YYYYMMDD')) AS data_key,
            CASE 
                WHEN NVL(I.BONIFIC, 'N') = 'N' THEN 
                    ROUND(NVL(I.QT, 0) * NVL(I.PVENDA, 0), 2)
                ELSE 0 
            END AS valor_faturamento,
            CASE 
                WHEN NVL(I.BONIFIC, 'N') = 'N' THEN 
                    ROUND(NVL(I.QT, 0) * NVL(I.VLCUSTOFIN, 0), 2)
                ELSE 0 
            END AS custo_faturamento
        FROM PCPEDC C
        JOIN PCPEDI I ON C.NUMPED = I.NUMPED
        JOIN PCPRODUT P ON I.CODPROD = P.CODPROD
        WHERE C.DATA >= TO_DATE('{data_inicio_str}', 'DD/MM/YYYY')
          AND C.DATA < TO_DATE('{data_fim_str}', 'DD/MM/YYYY')
          AND C.CODFILIAL IN ('1', '98')
          AND C.POSICAO = 'F'
          AND C.DTCANCEL IS NULL
          AND C.CONDVENDA NOT IN (4, 5, 8, 10, 13, 20, 98, 99)
          AND NVL(I.BONIFIC, 'N') = 'N'
    ),
    DEVOLUCOES_AGREGADAS AS (
        SELECT 
            D.CODUSUR AS codigovend,
            TO_NUMBER(TO_CHAR(D.DTENT, 'YYYYMMDD')) AS data_key,
            SUM(NVL(D.VLDEVOLUCAO, 0)) AS valor_devolucao_total,
            SUM(NVL(D.VLCUSTOFIN, 0)) AS custo_devolucao_total
        FROM VIEW_DEVOL_RESUMO_FATURAMENTO D
        WHERE D.DTENT >= TO_DATE('{data_inicio_str}', 'DD/MM/YYYY')
          AND D.DTENT < TO_DATE('{data_fim_str}', 'DD/MM/YYYY')
        GROUP BY D.CODUSUR, TO_NUMBER(TO_CHAR(D.DTENT, 'YYYYMMDD'))
    ),
    VENDAS_AGREGADAS AS (
        SELECT 
            codigovend,
            data_key,
            SUM(valor_faturamento) AS valor_faturamento_total,
            SUM(custo_faturamento) AS custo_faturamento_total
        FROM VENDAS V
        GROUP BY codigovend, data_key
    ),
    FATURAMENTO_LIQUIDO AS (
        SELECT 
            V.codigovend,
            V.codprod,
            V.data_key,
            CASE 
                WHEN NVL(VA.valor_faturamento_total, 0) > 0 THEN
                    V.valor_faturamento - ROUND((V.valor_faturamento / VA.valor_faturamento_total) * NVL(DA.valor_devolucao_total, 0), 2)
                ELSE V.valor_faturamento
            END AS faturamento_liquido
        FROM VENDAS V
        LEFT JOIN VENDAS_AGREGADAS VA ON V.codigovend = VA.codigovend AND V.data_key = VA.data_key
        LEFT JOIN DEVOLUCOES_AGREGADAS DA ON V.codigovend = DA.codigovend AND V.data_key = DA.data_key
    )
    SELECT 
        F.codigovend   AS codigovend,
        U.NOME         AS nomevend,
        F.codprod      AS codprod,
        P.DESCRICAO    AS zz_proddescricao,
        SUM(F.faturamento_liquido) AS faturamento_devolucoes_bi
    FROM FATURAMENTO_LIQUIDO F
    JOIN PCUSUARI U ON F.codigovend = U.CODUSUR
    JOIN PCPRODUT P ON F.codprod = P.CODPROD
    GROUP BY F.codigovend, U.NOME, F.codprod, P.DESCRICAO
    HAVING SUM(F.faturamento_liquido) <> 0
    ORDER BY F.codigovend, SUM(F.faturamento_liquido) DESC
    """

    try:
        df = pd.read_sql_query(query, conn)
        df.columns = [x.lower() for x in df.columns]
        df['faturamento_devolucoes_bi'] = pd.to_numeric(df['faturamento_devolucoes_bi'], errors='coerce').fillna(0).astype(float)
        return df
    except Exception as e:
        print(f"Erro ao buscar dados: {e}")
        import traceback
        traceback.print_exc()
        return pd.DataFrame()


if __name__ == '__main__':
    # Exemplo: mês 12/2025
    mes, ano = 12, 2025
    data_inicio = datetime(ano, mes, 1)
    if mes == 12:
        data_fim = datetime(ano + 1, 1, 1)
    else:
        data_fim = datetime(ano, mes + 1, 1)

    conn = get_db_connection()
    if not conn:
        print("Não foi possível conectar ao Oracle.")
        exit(1)

    df = buscar_comissao_dados(conn, data_inicio, data_fim)
    conn.close()

    print("Colunas:", list(df.columns))
    print("Total de linhas:", len(df))
    if not df.empty:
        print("\nPrimeiras linhas:")
        print(df.head(10).to_string())
        # Salvar exemplo CSV (nomes alinhados ao que você citou)
        out = Path("DADOS_BI_VENDAS")
        out.mkdir(exist_ok=True)
        df_out = df.rename(columns={
            'codigovend': 'CODIGOVEND',
            'nomevend': 'NOMEVEND',
            'codprod': 'CODPROD',
            'zz_proddescricao': 'zz_PRODDESCRICAO',
            'faturamento_devolucoes_bi': 'Faturamento - Devolucoes (BI)'
        })
        df_out.to_csv(out / "comissao_dados_oracle.csv", sep=';', decimal=',', index=False, encoding='utf-8-sig')
        print(f"\nArquivo salvo: {out / 'comissao_dados_oracle.csv'}")
# portfolio-commit-ready: comissao_dados_oracle
