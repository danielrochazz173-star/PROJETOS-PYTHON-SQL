"""
Script de análise de produtos - Novembro 2025
Baseado na lógica correta do dashboard Streamlit
"""

import oracledb
import pandas as pd
from datetime import datetime
import os

# Configuração do Oracle Instant Client
oracledb.init_oracle_client(lib_dir=r"C:\oracle\instantclient_23_9")

# =========================================================================
# CONFIGURAÇÕES
# =========================================================================
DB_HOST = os.getenv('DB_HOST', '10.0.0.10')
DB_PORT = int(os.getenv('DB_PORT', '1521'))
DB_SERVICE = os.getenv('DB_SERVICE', 'PROD')
DB_USER = os.getenv('DB_USER', 'powerbi')
DB_PASSWORD = os.getenv('DB_PASSWORD', '')

# =========================================================================
# QUERY SQL - BASEADA NO CÓDIGO CORRETO
# =========================================================================
QUERY_PRODUTOS = """
WITH VENDAS_VALIDAS AS (
    SELECT 
        I.CODPROD,
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
    JOIN PCUSUARI U ON C.CODUSUR = U.CODUSUR
    WHERE C.DATA BETWEEN TO_DATE('01/11/2025', 'DD/MM/YYYY') AND TO_DATE('30/11/2025', 'DD/MM/YYYY')
        AND C.CODFILIAL IN ('1', '98')
        AND C.POSICAO = 'F'
        AND C.DTCANCEL IS NULL
        AND C.CONDVENDA NOT IN (4, 8, 10, 13, 20, 98, 99)
        AND U.CODSUPERVISOR IN (1, 2)
    GROUP BY I.CODPROD
),
DEVOLUCOES AS (
    SELECT 
        D.CODPROD,
        SUM(NVL(D.VLDEVOLUCAO, 0)) AS VALOR_DEVOLVIDO,
        SUM(NVL(D.VLCUSTOFIN, 0)) AS CUSTO_DEVOLVIDO
    FROM VIEW_DEVOL_RESUMO_FATURAMENTO D
    JOIN PCUSUARI U ON D.CODUSUR = U.CODUSUR
    WHERE D.DTENT BETWEEN TO_DATE('01/11/2025', 'DD/MM/YYYY') AND TO_DATE('30/11/2025', 'DD/MM/YYYY')
        AND U.CODSUPERVISOR IN (1, 2)
    GROUP BY D.CODPROD
)
SELECT 
    P.CODPROD,
    P.DESCRICAO,
    (NVL(V.VALOR_BRUTO, 0) - NVL(D.VALOR_DEVOLVIDO, 0)) AS FATURAMENTO_LIQUIDO,
    (NVL(V.CUSTO_BRUTO, 0) - NVL(D.CUSTO_DEVOLVIDO, 0)) AS CUSTO_LIQUIDO
FROM PCPRODUT P 
LEFT JOIN VENDAS_VALIDAS V ON P.CODPROD = V.CODPROD 
LEFT JOIN DEVOLUCOES D ON P.CODPROD = D.CODPROD
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
    print("📊 ANÁLISE DE PRODUTOS - NOVEMBRO 2025")
    print("   (CODSUPERVISOR 1 e 2)")
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
        df = pd.read_sql_query(QUERY_PRODUTOS, conn)
        
        if df.empty:
            print("⚠️  Nenhum produto encontrado.")
            exit(0)
        
        # Converte valores para numéricos
        df['FATURAMENTO_LIQUIDO'] = pd.to_numeric(df['FATURAMENTO_LIQUIDO'], errors='coerce').fillna(0)
        df['CUSTO_LIQUIDO'] = pd.to_numeric(df['CUSTO_LIQUIDO'], errors='coerce').fillna(0)
        
        # Calcula margem no Python
        df['MARGEM'] = df.apply(
            lambda row: 0.0 if row['FATURAMENTO_LIQUIDO'] <= 0 
            else round(((row['FATURAMENTO_LIQUIDO'] - row['CUSTO_LIQUIDO']) / row['FATURAMENTO_LIQUIDO']) * 100, 2),
            axis=1
        )
        
        # Ordena por faturamento líquido
        df = df.sort_values('FATURAMENTO_LIQUIDO', ascending=False)
        
        print(f"✅ {len(df)} produtos encontrados.\n")
        print("=" * 80)
        print(f"{'CÓDIGO':<10} {'PRODUTO':<60} {'FAT. LÍQ.':>15} {'MARGEM':>10}")
        print("=" * 80)
        
        # Mostra todos os produtos
        for idx, row in df.iterrows():
            codigo = str(row['CODPROD'])[:10]
            produto = str(row['DESCRICAO'])[:58]
            fat_liq = f"R$ {row['FATURAMENTO_LIQUIDO']:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
            margem = f"{row['MARGEM']:.2f}%"
            
            print(f"{codigo:<10} {produto:<60} {fat_liq:>15} {margem:>10}")
        
        print("=" * 80)
        print(f"\n📊 RESUMO:")
        print(f"   Total de produtos: {len(df)}")
        print(f"   Faturamento líquido total: R$ {df['FATURAMENTO_LIQUIDO'].sum():,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
        print(f"   Margem média: {df['MARGEM'].mean():.2f}%")
        print()
        
        # Salva em CSV
        arquivo_csv = f"produtos_novembro_2025_corrigido_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        df[['CODPROD', 'DESCRICAO', 'FATURAMENTO_LIQUIDO', 'CUSTO_LIQUIDO', 'MARGEM']].to_csv(
            arquivo_csv, index=False, encoding='utf-8-sig'
        )
        print(f"💾 Dados salvos em: {arquivo_csv}")
        
    except Exception as e:
        print(f"❌ Erro ao executar query: {e}")
        import traceback
        traceback.print_exc()
    finally:
        conn.close()
        print("\n🔌 Conexão fechada.")

# portfolio-commit-ready: analise_produtos_novembro_corrigido
