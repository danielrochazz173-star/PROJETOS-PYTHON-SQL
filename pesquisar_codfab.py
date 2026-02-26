"""
Script rápido para pesquisar produto por CODFAB
"""

import oracledb

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

CODFAB_BUSCA = '00030080602'

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
    print(f"🔍 PESQUISANDO CODFAB: {CODFAB_BUSCA}")
    print("=" * 80)
    print()
    
    # Conecta ao banco
    print("🔌 Conectando ao Oracle...")
    conn = get_db_connection()
    
    if not conn:
        print("❌ Falha na conexão. Abortando.")
        exit(1)
    
    try:
        # Query de busca
        query = """
        SELECT 
            P.CODPROD,
            P.DESCRICAO,
            P.CODFAB,
            P.CODAUXILIAR AS EAN,
            P.EMBALAGEM,
            P.PESOLIQ,
            P.QTUNITCX,
            D.DESCRICAO AS DEPARTAMENTO,
            D.CODEPTO,
            COALESCE((E.QTESTGER - E.QTRESERV - E.QTBLOQUEADA) / NULLIF(P.QTUNITCX, 1), 
                     (E.QTESTGER - E.QTRESERV - E.QTBLOQUEADA), 0) AS ESTOQUE_CX
        FROM PCPRODUT P
        LEFT JOIN PCDEPTO D ON P.CODEPTO = D.CODEPTO
        LEFT JOIN PCEST E ON P.CODPROD = E.CODPROD AND E.CODFILIAL = '1'
        WHERE P.CODFAB = :codfab
          AND P.DTEXCLUSAO IS NULL
          AND (P.OBS2 <> 'FL' OR P.OBS2 IS NULL)
          AND P.REVENDA = 'S'
        ORDER BY P.DESCRICAO
        """
        
        print("📈 Executando query...")
        cursor = conn.cursor()
        cursor.execute(query, codfab=CODFAB_BUSCA)
        
        resultados = cursor.fetchall()
        
        if not resultados:
            print(f"⚠️  Nenhum produto encontrado com CODFAB = {CODFAB_BUSCA}")
        else:
            print(f"✅ {len(resultados)} produto(s) encontrado(s):\n")
            print("=" * 120)
            print(f"{'CODPROD':<10} {'DESCRIÇÃO':<50} {'CODFAB':<15} {'EAN':<15} {'DEPT.':<8} {'ESTOQUE':<10}")
            print("=" * 120)
            
            for row in resultados:
                codprod = str(row[0])[:10]
                descricao = str(row[1])[:48] if row[1] else 'N/A'
                codfab = str(row[2])[:15] if row[2] else 'N/A'
                ean = str(row[3])[:15] if row[3] else 'N/A'
                depto = str(row[8])[:6] if row[8] else 'N/A'
                estoque = f"{row[10]:.2f}" if row[10] else '0.00'
                
                print(f"{codprod:<10} {descricao:<50} {codfab:<15} {ean:<15} {depto:<8} {estoque:<10}")
            
            print("=" * 120)
            print()
            
            # Mostra detalhes do primeiro resultado
            if len(resultados) > 0:
                primeiro = resultados[0]
                print("📋 DETALHES DO PRIMEIRO PRODUTO:")
                print(f"   CODPROD: {primeiro[0]}")
                print(f"   DESCRIÇÃO: {primeiro[1]}")
                print(f"   CODFAB: {primeiro[2]}")
                print(f"   EAN: {primeiro[3]}")
                print(f"   EMBALAGEM: {primeiro[4]}")
                print(f"   PESOLIQ: {primeiro[5]}")
                print(f"   QTUNITCX: {primeiro[6]}")
                print(f"   DEPARTAMENTO: {primeiro[8]} (Código: {primeiro[9]})")
                print(f"   ESTOQUE: {primeiro[10]:.2f} caixas")
        
    except Exception as e:
        print(f"❌ Erro ao executar query: {e}")
        import traceback
        traceback.print_exc()
    finally:
        conn.close()
        print("\n🔌 Conexão fechada.")




