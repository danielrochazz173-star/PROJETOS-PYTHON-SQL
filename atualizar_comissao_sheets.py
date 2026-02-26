"""
Script para atualizar dados de comissão no Google Sheets
Busca dados do Oracle e joga na aba COMISSOESDADOS

Requer:
- pip install gspread oauth2client oracledb pandas
"""

import gspread
from oauth2client.service_account import ServiceAccountCredentials
import oracledb
import pandas as pd
from datetime import datetime
import sys

# ========== CONFIGURAÇÃO ORACLE ==========
try:
    oracledb.init_oracle_client(lib_dir=r"C:\oracle\instantclient_23_9")
except:
    pass

DB_CONFIG = {
    'user': 'powerbi',
    'password': 'cbjc4xp3nlq6',
    'dsn': '10.0.0.10:1521/PROD'
}

# ========== CONFIGURAÇÃO GOOGLE SHEETS ==========
SPREADSHEET_ID = '1Tvv9ERrCW_ilvP-j28f5ciyJebSxPDos060DFXbol6k'  # COMISSÕES VENDEDORES
CREDENTIALS_FILE = 'tensile-proxy-468412-t2-dd576fcb8c22.json'  # Seu arquivo de credenciais
ABA_DADOS = 'COMISSOESDADOS'

def conectar_oracle():
    """Conecta no banco Oracle"""
    try:
        conn = oracledb.connect(**DB_CONFIG)
        print("✅ Conectado ao Oracle")
        return conn
    except Exception as e:
        print(f"❌ Erro ao conectar no Oracle: {e}")
        sys.exit(1)

def conectar_google_sheets():
    """Conecta no Google Sheets"""
    try:
        scope = ['https://spreadsheets.google.com/feeds',
                 'https://www.googleapis.com/auth/drive']
        creds = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, scope)
        client = gspread.authorize(creds)
        print("✅ Conectado ao Google Sheets")
        return client
    except Exception as e:
        print(f"❌ Erro ao conectar no Google Sheets: {e}")
        sys.exit(1)

def buscar_dados_comissao(conn, data_inicio, data_fim):
    """
    Busca dados SIMPLES: RCA, NOME, CODPROD, DESCRICAO, FATURAMENTO, MARGEM
    O Apps Script vai juntar e calcular a margem geral
    """
    print(f"📊 Buscando dados de {data_inicio.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')}...")
    
    data_inicio_str = data_inicio.strftime('%d/%m/%Y')
    data_fim_str = data_fim.strftime('%d/%m/%Y')
    
    query = f"""
    WITH VENDAS AS (
        SELECT 
            P.CODUSUR,
            U.NOME AS NOME_VENDEDOR,
            I.CODPROD,
            PR.DESCRICAO AS PRODUTO_DESCRICAO,
            ROUND(SUM(I.QT * I.PVENDA), 2) AS VALOR_FATURAMENTO,
            ROUND(SUM(I.QT * I.VLCUSTOFIN), 2) AS CUSTO_FATURAMENTO
        FROM PCPEDC P
        INNER JOIN PCPEDI I ON P.NUMPED = I.NUMPED
        INNER JOIN PCUSUARI U ON P.CODUSUR = U.CODUSUR
        INNER JOIN PCPRODUT PR ON I.CODPROD = PR.CODPROD
        WHERE P.DATA >= TO_DATE('{data_inicio_str}', 'DD/MM/YYYY')
          AND P.DATA < TO_DATE('{data_fim_str}', 'DD/MM/YYYY')
          AND P.CODFILIAL IN ('1', '98')
          AND P.POSICAO = 'F'
          AND P.DTCANCEL IS NULL
          AND NVL(I.BONIFIC, 'N') = 'N'
        GROUP BY P.CODUSUR, U.NOME, I.CODPROD, PR.DESCRICAO
    ),
    DEVOLUCOES AS (
        SELECT
            D.CODUSUR,
            D.CODPROD,
            ROUND(SUM(D.QT * D.PUNIT), 2) AS VALOR_DEVOLUCAO,
            ROUND(SUM(D.QT * D.CUSTOFIN), 2) AS CUSTO_DEVOLUCAO
        FROM PCNFSAID N
        INNER JOIN PCMOV D ON N.NUMTRANSVENDA = D.NUMTRANSVENDA
        WHERE N.DTCANCEL IS NULL
          AND N.CONDVENDA NOT IN (7, 8, 10, 13, 14, 98, 99)
          AND N.TIPOVENDA NOT IN ('SR', 'OR')
          AND D.DTMOV >= TO_DATE('{data_inicio_str}', 'DD/MM/YYYY')
          AND D.DTMOV < TO_DATE('{data_fim_str}', 'DD/MM/YYYY')
          AND D.STATUS = 'EF'
        GROUP BY D.CODUSUR, D.CODPROD
    )
    SELECT 
        V.CODUSUR,
        V.NOME_VENDEDOR,
        V.CODPROD,
        V.PRODUTO_DESCRICAO,
        ROUND(V.VALOR_FATURAMENTO - NVL(D.VALOR_DEVOLUCAO, 0), 2) AS FATURAMENTO_LIQUIDO,
        CASE 
            WHEN (V.VALOR_FATURAMENTO - NVL(D.VALOR_DEVOLUCAO, 0)) > 0 THEN
                ROUND(
                    ((V.VALOR_FATURAMENTO - NVL(D.VALOR_DEVOLUCAO, 0)) - 
                     (V.CUSTO_FATURAMENTO - NVL(D.CUSTO_DEVOLUCAO, 0))) / 
                    (V.VALOR_FATURAMENTO - NVL(D.VALOR_DEVOLUCAO, 0)) * 100,
                    2
                )
            ELSE 0
        END AS MARGEM_PERCENTUAL
    FROM VENDAS V
    LEFT JOIN DEVOLUCOES D ON V.CODUSUR = D.CODUSUR AND V.CODPROD = D.CODPROD
    WHERE (V.VALOR_FATURAMENTO - NVL(D.VALOR_DEVOLUCAO, 0)) > 0
    ORDER BY V.CODUSUR, V.CODPROD
    """
    
    df = pd.read_sql(query, conn)
    print(f"   ✅ {len(df)} registros encontrados")
    
    return df

def atualizar_planilha(gc, df):
    """Atualiza a aba COMISSOESDADOS no Google Sheets - SIMPLES"""
    try:
        print(f"📝 Atualizando planilha...")
        
        # Abre a planilha
        planilha = gc.open_by_key(SPREADSHEET_ID)
        
        # Tenta encontrar a aba, se não existir, cria
        try:
            aba = planilha.worksheet(ABA_DADOS)
        except gspread.exceptions.WorksheetNotFound:
            print(f"   ⚠️ Aba '{ABA_DADOS}' não encontrada. Criando...")
            aba = planilha.add_worksheet(title=ABA_DADOS, rows=5000, cols=10)
        
        # Limpa a aba
        aba.clear()
        
        # SIMPLES: RCA, NOME, CODPROD, DESCRICAO, FATURAMENTO, MARGEM
        header = [
            'CODUSUR', 
            'NOME_VENDEDOR', 
            'CODPROD', 
            'PRODUTO_DESCRICAO', 
            'FATURAMENTO_LIQUIDO',
            'MARGEM_PERCENTUAL'
        ]
        
        # Seleciona as colunas
        df_final = df[[
            'CODUSUR', 
            'NOME_VENDEDOR', 
            'CODPROD', 
            'PRODUTO_DESCRICAO', 
            'FATURAMENTO_LIQUIDO',
            'MARGEM_PERCENTUAL'
        ]]
        
        # Converte para lista de listas
        dados = [header] + df_final.values.tolist()
        
        # Atualiza a planilha em lote (muito mais rápido)
        aba.update(values=dados, range_name='A1', value_input_option='USER_ENTERED')
        
        print(f"   ✅ {len(df_final)} linhas atualizadas na aba '{ABA_DADOS}'")
        print(f"\n🎯 Agora você pode rodar o script do Apps Script para calcular as comissões!")
        
    except Exception as e:
        print(f"❌ Erro ao atualizar planilha: {e}")
        raise

def main():
    """Função principal"""
    print("=" * 60)
    print("🚀 ATUALIZAÇÃO DE DADOS DE COMISSÃO")
    print("=" * 60)
    
    # Pede mês e ano
    try:
        mes = int(input("Digite o mês (1-12): "))
        ano = int(input("Digite o ano (ex: 2025): "))
        
        if mes < 1 or mes > 12:
            print("❌ Mês inválido!")
            sys.exit(1)
        
        # Calcula data início e fim
        data_inicio = datetime(ano, mes, 1)
        if mes == 12:
            data_fim = datetime(ano + 1, 1, 1)
        else:
            data_fim = datetime(ano, mes + 1, 1)
        
    except ValueError:
        print("❌ Valores inválidos!")
        sys.exit(1)
    
    print(f"\n📅 Período: {data_inicio.strftime('%d/%m/%Y')} a {data_fim.strftime('%d/%m/%Y')}")
    print()
    
    # Conecta no Oracle
    conn = conectar_oracle()
    
    # Busca dados
    df = buscar_dados_comissao(conn, data_inicio, data_fim)
    conn.close()
    
    if len(df) == 0:
        print("⚠️ Nenhum dado encontrado para o período!")
        sys.exit(0)
    
    # Mostra resumo
    print(f"\n📊 RESUMO:")
    print(f"   Vendedores: {df['CODUSUR'].nunique()}")
    print(f"   Produtos: {df['CODPROD'].nunique()}")
    print(f"   Faturamento Total: R$ {df['FATURAMENTO_LIQUIDO'].sum():,.2f}")
    print()
    
    # Conecta no Google Sheets
    gc = conectar_google_sheets()
    
    # Atualiza planilha
    atualizar_planilha(gc, df)
    
    print("\n" + "=" * 60)
    print("✅ CONCLUÍDO!")
    print("=" * 60)

if __name__ == "__main__":
    main()
# portfolio-commit-ready: atualizar_comissao_sheets
