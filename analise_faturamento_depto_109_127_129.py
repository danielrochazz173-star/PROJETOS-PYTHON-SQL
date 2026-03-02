"""
Script de análise de faturamento por departamentos 109, 127 e 129
Busca todas as notas fiscais faturadas desses departamentos nos dias 24 e 25
Permite escolher o mês/ano para análise
"""

import oracledb
import pandas as pd
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
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

# Departamentos a buscar
CODEPTO_LIST = [109, 127, 129]

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

def get_query_faturamento(mes, ano, todos_supervisores=True):
    """Gera a query SQL para buscar faturamento por departamento nos dias 24 e 25
    Usa a lógica correta de faturamento com DECODE para condições de venda e devoluções"""
    data_24 = f"24/{mes:02d}/{ano}"
    data_25 = f"25/{mes:02d}/{ano}"
    
    codepto_str = ', '.join(map(str, CODEPTO_LIST))
    
    # Filtro de supervisor - se todos_supervisores for False, filtra por supervisor 1 e 2
    if todos_supervisores:
        filtro_supervisor = ""
        filtro_supervisor_devol = ""
    else:
        filtro_supervisor = "AND U.CODSUPERVISOR IN (1, 2)"
        filtro_supervisor_devol = "AND U.CODSUPERVISOR IN (1, 2)"
    
    query = f"""
    WITH VENDAS_VALIDAS AS (
        SELECT 
            CL.CODCLI,
            CL.CLIENTE,
            C.NUMNOTA,
            C.NUMPED,
            C.DATA,
            I.CODPROD,
            P.DESCRICAO AS PRODUTO,
            D.CODEPTO,
            D.DESCRICAO AS DEPARTAMENTO,
            I.PVENDA,
            I.QT,
            C.CODUSUR,
            U.NOME AS VENDEDOR,
            C.CONDVENDA,
            -- Calcula valor bruto usando DECODE para condições de venda
            CASE 
                WHEN NVL(I.BONIFIC, 'N') = 'N' THEN 
                    DECODE(C.CONDVENDA, 
                        5, 0, 6, 0, 11, 0, 12, 0, 
                        ROUND(NVL(I.QT, 0) * NVL(I.PVENDA, 0), 2)
                    )
                ELSE 0 
            END AS VALOR_BRUTO
        FROM PCPEDC C
        JOIN PCPEDI I ON C.NUMPED = I.NUMPED
        JOIN PCCLIENT CL ON C.CODCLI = CL.CODCLI
        JOIN PCPRODUT P ON I.CODPROD = P.CODPROD
        JOIN PCDEPTO D ON P.CODEPTO = D.CODEPTO
        LEFT JOIN PCUSUARI U ON C.CODUSUR = U.CODUSUR
        WHERE (C.DATA = TO_DATE('{data_24}', 'DD/MM/YYYY') 
            OR C.DATA = TO_DATE('{data_25}', 'DD/MM/YYYY'))
            AND P.CODEPTO IN ({codepto_str})
            AND C.CODFILIAL IN ('1', '98')
            AND C.POSICAO = 'F'
            AND C.DTCANCEL IS NULL
            AND C.CONDVENDA NOT IN (4, 8, 10, 13, 20, 98, 99)
            AND NVL(I.BONIFIC, 'N') = 'N'
            {filtro_supervisor}
    ),
    DEVOLUCOES AS (
        SELECT 
            D.CODPROD,
            D.CODUSUR,
            SUM(NVL(D.VLDEVOLUCAO, 0)) AS VALOR_DEVOLVIDO
        FROM VIEW_DEVOL_RESUMO_FATURAMENTO D
        JOIN PCUSUARI U ON D.CODUSUR = U.CODUSUR
        JOIN PCPRODUT P ON D.CODPROD = P.CODPROD
        WHERE (D.DTENT = TO_DATE('{data_24}', 'DD/MM/YYYY')
           OR D.DTENT = TO_DATE('{data_25}', 'DD/MM/YYYY'))
           AND P.CODEPTO IN ({codepto_str})
           {filtro_supervisor_devol}
        GROUP BY D.CODPROD, D.CODUSUR
    )
    SELECT 
        V.CODCLI,
        V.CLIENTE,
        V.NUMNOTA,
        V.DATA,
        V.CODPROD,
        V.PRODUTO,
        V.PVENDA,
        V.QT,
        V.VENDEDOR,
        (V.VALOR_BRUTO - NVL(DEV.VALOR_DEVOLVIDO, 0)) AS VALOR_TOTAL
    FROM VENDAS_VALIDAS V
    LEFT JOIN DEVOLUCOES DEV ON V.CODPROD = DEV.CODPROD AND V.CODUSUR = DEV.CODUSUR
    WHERE V.VALOR_BRUTO > 0 OR NVL(DEV.VALOR_DEVOLVIDO, 0) > 0
    ORDER BY V.DATA DESC, V.CODCLI, V.CODPROD
    """
    
    return query

# =========================================================================
# EXECUÇÃO
# =========================================================================
if __name__ == "__main__":
    print("=" * 100)
    print("📊 ANÁLISE DE FATURAMENTO - DEPARTAMENTOS 109, 127 E 129")
    print("=" * 100)
    print()
    
    # Solicita mês e ano
    try:
        ano = int(input("Digite o ano (ex: 2025): "))
        mes = int(input("Digite o mês (1-12): "))
        
        if mes < 1 or mes > 12:
            print("❌ Mês inválido. Deve ser entre 1 e 12.")
            exit(1)
        
        print()
        opcao_supervisor = input("Buscar de todos os supervisores? (S/N) [S]: ").strip().upper()
        todos_supervisores = opcao_supervisor != 'N'
        
        print()
        print(f"🔍 Buscando faturamento dos departamentos {CODEPTO_LIST}")
        print(f"   Período: Dias 24 e 25 de {mes:02d}/{ano}")
        if todos_supervisores:
            print(f"   Supervisores: Todos")
        else:
            print(f"   Supervisores: 1 e 2")
        print()
        
    except ValueError:
        print("❌ Entrada inválida. Use apenas números.")
        exit(1)
    
    # Conecta ao banco
    print("🔌 Conectando ao Oracle...")
    conn = get_db_connection()
    
    if not conn:
        print("❌ Falha na conexão. Abortando.")
        exit(1)
    
    try:
        # Executa a query
        print("📈 Executando query...")
        query = get_query_faturamento(mes, ano, todos_supervisores)
        df = pd.read_sql_query(query, conn)
        
        if df.empty:
            print("⚠️  Nenhum registro encontrado com os critérios especificados.")
            exit(0)
        
        print(f"✅ {len(df)} registro(s) encontrado(s).\n")
        print("=" * 140)
        print(f"{'CÓD.':<8} {'CLIENTE':<35} {'NF':<12} {'DATA':<12} {'PROD.':<8} {'PRODUTO':<30} {'PVENDA':<12} {'QTD':<8} {'VENDEDOR':<20} {'VALOR':<12}")
        print("=" * 140)
        
        # Mostra todos os registros
        for idx, row in df.iterrows():
            codcli = str(row['CODCLI'])[:8]
            cliente = str(row['CLIENTE'])[:33]
            numnota = str(row['NUMNOTA'])[:12] if pd.notna(row['NUMNOTA']) else 'N/A'
            data = row['DATA'].strftime('%d/%m/%Y') if pd.notna(row['DATA']) else 'N/A'
            codprod = str(row['CODPROD'])[:8]
            produto = str(row['PRODUTO'])[:28]
            pvenda = f"R$ {row['PVENDA']:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
            qt = f"{row['QT']:.0f}" if pd.notna(row['QT']) else '0'
            vendedor = str(row['VENDEDOR'])[:18] if pd.notna(row['VENDEDOR']) else 'N/A'
            valor = f"R$ {row['VALOR_TOTAL']:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
            
            print(f"{codcli:<8} {cliente:<35} {numnota:<12} {data:<12} {codprod:<8} {produto:<30} {pvenda:<12} {qt:<8} {vendedor:<20} {valor:<12}")
        
        print("=" * 140)
        print(f"\n📊 RESUMO GERAL:")
        print(f"   Total de registros: {len(df)}")
        print(f"   Clientes únicos: {df['CODCLI'].nunique()}")
        print(f"   Notas fiscais únicas: {df['NUMNOTA'].nunique()}")
        print(f"   Valor total: R$ {df['VALOR_TOTAL'].sum():,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
        print()
        
        # Salva em Excel formatado
        supervisor_sufixo = "todos" if todos_supervisores else "sup1_2"
        arquivo_excel = f"faturamento_depto_109_127_129_dias24_25_{mes:02d}_{ano}_{supervisor_sufixo}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        # Seleciona apenas as colunas solicitadas para o Excel
        df_excel = df[['CODCLI', 'CLIENTE', 'NUMNOTA', 'DATA', 'CODPROD', 'PRODUTO', 'PVENDA', 'QT', 'VENDEDOR', 'VALOR_TOTAL']].copy()
        df_excel = df_excel.rename(columns={'NUMNOTA': 'NF'})
        
        with pd.ExcelWriter(arquivo_excel, engine='openpyxl') as writer:
            # Aba 1: Detalhado
            df_excel.to_excel(writer, sheet_name='Detalhado', index=False)
            ws_detalhado = writer.sheets['Detalhado']
            
            # Formata cabeçalho
            header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
            header_font = Font(bold=True, color="FFFFFF", size=11)
            border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            
            for cell in ws_detalhado[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = border
            
            # Ajusta largura das colunas
            ws_detalhado.column_dimensions['A'].width = 10  # CODCLI
            ws_detalhado.column_dimensions['B'].width = 40  # CLIENTE
            ws_detalhado.column_dimensions['C'].width = 12  # NF
            ws_detalhado.column_dimensions['D'].width = 12  # DATA
            ws_detalhado.column_dimensions['E'].width = 10  # CODPROD
            ws_detalhado.column_dimensions['F'].width = 40  # PRODUTO
            ws_detalhado.column_dimensions['G'].width = 12  # PVENDA
            ws_detalhado.column_dimensions['H'].width = 10  # QT
            ws_detalhado.column_dimensions['I'].width = 25  # VENDEDOR
            ws_detalhado.column_dimensions['J'].width = 15  # VALOR_TOTAL
            
            # Formata dados
            for row in ws_detalhado.iter_rows(min_row=2, max_row=ws_detalhado.max_row):
                for cell in row:
                    cell.border = border
                    if cell.column == 10:  # VALOR_TOTAL
                        cell.number_format = 'R$ #,##0.00'
                    elif cell.column == 7:  # PVENDA
                        cell.number_format = 'R$ #,##0.00'
                    elif cell.column == 4:  # DATA
                        cell.number_format = 'DD/MM/YYYY'
            
            # Congela primeira linha
            ws_detalhado.freeze_panes = 'A2'
        
        print(f"💾 Dados salvos em Excel: {arquivo_excel}")
        
    except Exception as e:
        print(f"❌ Erro ao executar query: {e}")
        import traceback
        traceback.print_exc()
    finally:
        conn.close()
        print("\n🔌 Conexão fechada.")

# portfolio-commit-ready: analise_faturamento_depto_109_127_129
