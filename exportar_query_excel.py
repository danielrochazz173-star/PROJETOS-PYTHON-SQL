"""
Script para exportar qualquer query SQL do Oracle para Excel
Aceita SELECT * FROM ou qualquer query SQL e gera Excel com colunas separadas
"""

import oracledb
import pandas as pd
from datetime import datetime
from pathlib import Path
import sys
import os

# Configuração do Oracle Instant Client
try:
    oracledb.init_oracle_client(lib_dir=r"C:\oracle\instantclient_23_9")
except Exception:
    pass  # Pode já estar inicializado

# =========================================================================
# CONFIGURAÇÕES
# =========================================================================
DB_HOST = '10.0.0.10'
DB_PORT = 1521
DB_SERVICE = 'PROD'
DB_USER = 'powerbi'
DB_PASSWORD = 'cbjc4xp3nlq6'

# =========================================================================
# FUNÇÕES
# =========================================================================
def get_db_connection():
    """Conecta ao banco Oracle"""
    try:
        return oracledb.connect(
            user=DB_USER,
            password=DB_PASSWORD,
            dsn=f"{DB_HOST}:{DB_PORT}/{DB_SERVICE}",
        )
    except oracledb.DatabaseError as e:
        print(f"❌ Erro ao conectar ao Oracle: {e}")
        return None

def executar_query_para_excel(query_sql, nome_arquivo=None):
    """
    Executa uma query SQL e exporta para Excel
    
    Args:
        query_sql: Query SQL a ser executada (ex: "SELECT * FROM pcclient")
        nome_arquivo: Nome do arquivo Excel (opcional, será gerado automaticamente se não informado)
    """
    print("🔌 Conectando ao Oracle...")
    conn = get_db_connection()
    
    if not conn:
        print("❌ Não foi possível conectar ao banco de dados!")
        return None
    
    try:
        print(f"📊 Executando query...")
        print(f"   Query: {query_sql[:100]}..." if len(query_sql) > 100 else f"   Query: {query_sql}")
        
        # Executa a query e carrega em DataFrame
        df = pd.read_sql(query_sql, conn)
        
        if df.empty:
            print("⚠️  A query não retornou nenhum resultado!")
            return None
        
        print(f"✅ Query executada com sucesso! {len(df)} registros encontrados.")
        print(f"📋 Colunas: {', '.join(df.columns.tolist())}")
        
        # Gera nome do arquivo se não foi informado
        if not nome_arquivo:
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            # Tenta extrair nome da tabela da query
            query_upper = query_sql.upper()
            if "FROM" in query_upper:
                try:
                    tabela = query_upper.split("FROM")[1].strip().split()[0]
                    nome_arquivo = f"export_{tabela}_{timestamp}.xlsx"
                except:
                    nome_arquivo = f"export_query_{timestamp}.xlsx"
            else:
                nome_arquivo = f"export_query_{timestamp}.xlsx"
        
        # Garante extensão .xlsx
        if not nome_arquivo.endswith('.xlsx'):
            nome_arquivo += '.xlsx'
        
        # Salva no diretório atual
        caminho_arquivo = Path(nome_arquivo)
        
        print(f"💾 Exportando para Excel: {caminho_arquivo}")
        
        # Exporta para Excel (sem formatação, apenas dados puros)
        df.to_excel(caminho_arquivo, index=False, engine='openpyxl')
        
        print(f"✅ Arquivo Excel criado com sucesso!")
        print(f"   📁 Local: {caminho_arquivo.absolute()}")
        print(f"   📊 Registros: {len(df)}")
        print(f"   📋 Colunas: {len(df.columns)}")
        
        return str(caminho_arquivo.absolute())
        
    except Exception as e:
        print(f"❌ Erro ao executar query ou exportar Excel: {e}")
        import traceback
        traceback.print_exc()
        return None
    finally:
        conn.close()
        print("🔌 Conexão fechada.")

def main():
    """Função principal"""
    print("=" * 60)
    print("📊 EXPORTADOR DE QUERIES SQL PARA EXCEL")
    print("=" * 60)
    print()
    
    # Verifica se a query foi passada como argumento
    if len(sys.argv) > 1:
        # Query passada como argumento (remove aspas se houver)
        query_sql = " ".join(sys.argv[1:]).strip('"').strip("'")
        nome_arquivo = None
    else:
        # Solicita query do usuário
        print("Digite a query SQL que deseja executar:")
        print("Exemplo: SELECT * FROM pcclient")
        print("Exemplo: SELECT * FROM pcprodut WHERE codprod = 123")
        print()
        query_sql = input("Query SQL: ").strip()
        
        if not query_sql:
            print("❌ Query não informada!")
            return
        
        # Pergunta se quer nomear o arquivo
        nome_arquivo = input("Nome do arquivo Excel (Enter para automático): ").strip()
        if not nome_arquivo:
            nome_arquivo = None
    
    # Executa e exporta
    resultado = executar_query_para_excel(query_sql, nome_arquivo)
    
    if resultado:
        print()
        print("=" * 60)
        print("✅ Processo concluído com sucesso!")
        print("=" * 60)
    else:
        print()
        print("=" * 60)
        print("❌ Processo falhou!")
        print("=" * 60)

if __name__ == "__main__":
    main()

# portfolio-commit-ready: exportar_query_excel
