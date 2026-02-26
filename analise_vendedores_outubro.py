"""
Script rápido para análise de vendedores - Outubro 2025
Mostra vendedores ordenados por faturamento líquido e margem
"""

import oracledb
import pandas as pd
from datetime import datetime

# Configuração do Oracle Instant Client
oracledb.init_oracle_client(lib_dir=r"C:\oracle\instantclient_23_9")

# =========================================================================
# CONFIGURAÇÕES
# =========================================================================
DB_HOST = '10.0.0.10'
DB_PORT = 1521
DB_SERVICE = 'PROD'
DB_USER = 'powerbi'
DB_PASSWORD = 'cbjc4xp3nlq6'

# =========================================================================
# QUERY SQL
# =========================================================================
QUERY_VENDEDORES = """
WITH VENDAS_VALIDAS AS (
    SELECT 
        C.CODUSUR,
        SUM(
            CASE 
                WHEN NVL(I.BONIFIC, 'N') = 'N' THEN 
                    DECODE(C.CONDVENDA, 
                        5, 0, 6, 0, 11, 0, 12, 0, 
                        ROUND(NVL(I.QT, 0) * NVL(I.PVENDA, 0), 2)
                    )
                ELSE 0 
            END
        ) AS VALOR_BRUTO,
        SUM(
            CASE 
                WHEN NVL(I.BONIFIC, 'N') = 'N' THEN 
                    DECODE(C.CONDVENDA, 
                        5, 0, 6, 0, 11, 0, 12, 0, 
                        NVL(I.QT, 0) * NVL(I.VLCUSTOFIN, 0)
                    )
                ELSE 0 
            END
        ) AS CUSTO_BRUTO
    FROM PCPEDC C 
    JOIN PCPEDI I ON C.NUMPED = I.NUMPED
    WHERE C.DATA BETWEEN TO_DATE('01/10/2025', 'DD/MM/YYYY') AND TO_DATE('31/10/2025', 'DD/MM/YYYY')
        AND C.CODFILIAL IN ('1', '98')
        AND C.POSICAO = 'F'
        AND C.DTCANCEL IS NULL
        AND C.CONDVENDA NOT IN (4, 8, 10, 13, 20, 98, 99)
    GROUP BY C.CODUSUR
),
DEVOLUCOES AS (
    SELECT 
        D.CODUSUR,
        SUM(NVL(D.VLDEVOLUCAO, 0)) AS VALOR_DEVOLVIDO,
        SUM(NVL(D.VLCUSTOFIN, 0)) AS CUSTO_DEVOLVIDO
    FROM VIEW_DEVOL_RESUMO_FATURAMENTO D
    WHERE D.DTENT BETWEEN TO_DATE('01/10/2025', 'DD/MM/YYYY') AND TO_DATE('31/10/2025', 'DD/MM/YYYY')
    GROUP BY D.CODUSUR
)
SELECT 
    U.CODUSUR AS "CODIGO_VENDEDOR",
    U.NOME AS "NOME_VENDEDOR",
    (NVL(V.VALOR_BRUTO, 0) - NVL(D.VALOR_DEVOLVIDO, 0)) AS FATURAMENTO_LIQUIDO,
    (NVL(V.CUSTO_BRUTO, 0) - NVL(D.CUSTO_DEVOLVIDO, 0)) AS CUSTO_LIQUIDO
FROM PCUSUARI U 
LEFT JOIN VENDAS_VALIDAS V ON U.CODUSUR = V.CODUSUR 
LEFT JOIN DEVOLUCOES D ON U.CODUSUR = D.CODUSUR
WHERE (NVL(V.VALOR_BRUTO, 0) > 0 OR NVL(D.VALOR_DEVOLVIDO, 0) > 0)
ORDER BY FATURAMENTO_LIQUIDO DESC
"""

# =========================================================================
# FUNÇÕES
# =========================================================================
def get_db_connection():
    """Conecta ao banco Oracle"""
    try:
        return oracledb.connect(user=DB_USER, password=DB_PASSWORD, dsn=f"{DB_HOST}:{DB_PORT}/{DB_SERVICE}")
    except oracledb.DatabaseError as e:
        print(f"❌ Erro ao conectar ao Oracle: {e}")
        return None


# =========================================================================
# EXECUÇÃO
# =========================================================================
if __name__ == "__main__":
    print("=" * 80)
    print("📊 ANÁLISE DE VENDEDORES - OUTUBRO 2025")
    print("=" * 80)
    print()
    
    # Conecta ao banco
    print("🔌 Conectando ao Oracle...")
    conn = get_db_connection()
    
    if not conn:
        print("❌ Falha na conexão. Abortando.")
        exit(1)
    
    try:
        # Executa a query
        print("📈 Executando query...")
        df = pd.read_sql_query(QUERY_VENDEDORES, conn)
        
        if df.empty:
            print("⚠️  Nenhum vendedor encontrado.")
            exit(0)
        
        # Converte valores para numéricos (já vêm como números do Oracle)
        df['FATURAMENTO_LIQUIDO'] = pd.to_numeric(df['FATURAMENTO_LIQUIDO'], errors='coerce').fillna(0)
        df['CUSTO_LIQUIDO'] = pd.to_numeric(df['CUSTO_LIQUIDO'], errors='coerce').fillna(0)
        
        # Calcula margem no Python (igual produtos)
        df['MARGEM'] = df.apply(
            lambda row: 0.0 if row['FATURAMENTO_LIQUIDO'] <= 0 
            else round(((row['FATURAMENTO_LIQUIDO'] - row['CUSTO_LIQUIDO']) / row['FATURAMENTO_LIQUIDO']) * 100, 2),
            axis=1
        )
        
        # Ordena por faturamento líquido (já vem ordenado, mas garantimos)
        df = df.sort_values('FATURAMENTO_LIQUIDO', ascending=False)
        
        print(f"✅ {len(df)} vendedores encontrados.\n")
        print("=" * 80)
        print(f"{'CÓDIGO':<10} {'NOME VENDEDOR':<50} {'FAT. LÍQ.':>15} {'MARGEM':>10}")
        print("=" * 80)
        
        # Mostra todos os vendedores
        for idx, row in df.iterrows():
            codigo = str(row['CODIGO_VENDEDOR'])[:10]
            nome = str(row['NOME_VENDEDOR'])[:48]
            fat_liq = f"R$ {row['FATURAMENTO_LIQUIDO']:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
            margem = f"{row['MARGEM']:.2f}%"
            
            print(f"{codigo:<10} {nome:<50} {fat_liq:>15} {margem:>10}")
        
        print("=" * 80)
        print(f"\n📊 RESUMO:")
        print(f"   Total de vendedores: {len(df)}")
        print(f"   Faturamento líquido total: R$ {df['FATURAMENTO_LIQUIDO'].sum():,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
        print(f"   Margem média: {df['MARGEM'].mean():.2f}%")
        print(f"   Maior faturamento: R$ {df['FATURAMENTO_LIQUIDO'].max():,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
        print(f"   Menor faturamento: R$ {df['FATURAMENTO_LIQUIDO'].min():,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
        print()
        
        # Salva em CSV
        arquivo_csv = f"vendedores_outubro_2025_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        df.to_csv(arquivo_csv, index=False, encoding='utf-8-sig')
        print(f"💾 Dados salvos em: {arquivo_csv}")
        
    except Exception as e:
        print(f"❌ Erro ao executar query: {e}")
        import traceback
        traceback.print_exc()
    finally:
        conn.close()
        print("\n🔌 Conexão fechada.")

# portfolio-commit-ready: analise_vendedores_outubro
