"""
Script universal de análise de vendas de produtos
- Aceita múltiplos códigos de produtos separados por vírgula
- Aceita múltiplos preços separados por vírgula (opcional)
- Mostra todas as vendas do mês escolhido com NF, data e preço
"""

import oracledb
import pandas as pd
from datetime import datetime
from calendar import monthrange
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

def get_query_vendas(mes, ano, codprod_list, precos_list=None):
    """Gera a query SQL para buscar vendas dos produtos informados"""
    data_inicio = f"01/{mes:02d}/{ano}"
    # Calcula o último dia do mês
    ultimo_dia = monthrange(ano, mes)[1]
    data_fim = f"{ultimo_dia:02d}/{mes:02d}/{ano}"
    
    codprod_str = ', '.join(map(str, codprod_list))
    
    # Monta filtro de preços se fornecido
    filtro_preco = ""
    if precos_list and len(precos_list) > 0:
        precos_str = ', '.join([f"{preco:.2f}" for preco in precos_list])
        filtro_preco = f"AND ROUND(I.PVENDA, 2) IN ({precos_str})"
    
    query = f"""
    SELECT 
        C.NUMNOTA AS NF,
        C.DATA,
        I.CODPROD,
        P.DESCRICAO AS PRODUTO,
        I.PVENDA AS PRECO,
        I.QT AS QUANTIDADE,
        ROUND(I.QT * I.PVENDA, 2) AS VALOR_TOTAL,
        CL.CODCLI,
        CL.CLIENTE,
        C.CODUSUR,
        U.NOME AS VENDEDOR
    FROM PCPEDC C
    JOIN PCPEDI I ON C.NUMPED = I.NUMPED
    JOIN PCCLIENT CL ON C.CODCLI = CL.CODCLI
    JOIN PCPRODUT P ON I.CODPROD = P.CODPROD
    LEFT JOIN PCUSUARI U ON C.CODUSUR = U.CODUSUR
    WHERE C.DATA BETWEEN TO_DATE('{data_inicio}', 'DD/MM/YYYY') 
        AND TO_DATE('{data_fim}', 'DD/MM/YYYY')
        AND I.CODPROD IN ({codprod_str})
        {filtro_preco}
        AND C.CODFILIAL IN ('1', '98')
        AND C.POSICAO = 'F'
        AND C.DTCANCEL IS NULL
        AND C.CONDVENDA NOT IN (4, 8, 10, 13, 20, 98, 99)
        AND NVL(I.BONIFIC, 'N') = 'N'
    ORDER BY C.DATA DESC, C.NUMNOTA, I.CODPROD
    """
    
    return query

# =========================================================================
# EXECUÇÃO
# =========================================================================
if __name__ == "__main__":
    print("=" * 100)
    print("📊 ANÁLISE DE VENDAS - PRODUTOS (UNIVERSAL)")
    print("=" * 100)
    print()
    
    # Solicita códigos dos produtos
    try:
        codigos_input = input("Digite os códigos dos produtos separados por vírgula (ex: 2608 ou 2608,875,879): ").strip()
        
        if not codigos_input:
            print("❌ Nenhum código informado.")
            exit(1)
        
        # Processa códigos
        codprod_list = []
        for cod in codigos_input.split(','):
            cod_limpo = cod.strip()
            if cod_limpo.isdigit():
                codprod_list.append(int(cod_limpo))
        
        if not codprod_list:
            print("❌ Nenhum código válido encontrado.")
            exit(1)
        
        # Solicita preços (opcional)
        precos_input = input("Digite os preços separados por vírgula (ex: 9.90 ou 9.90,8.50) - ou ENTER para todos: ").strip()
        
        precos_list = None
        if precos_input:
            precos_list = []
            for preco_str in precos_input.split(','):
                preco_limpo = preco_str.strip()
                try:
                    # Converte ponto para vírgula se necessário e depois para float
                    preco_float = float(preco_limpo.replace(',', '.'))
                    precos_list.append(preco_float)
                except ValueError:
                    print(f"⚠️  Preço inválido ignorado: {preco_limpo}")
            
            if len(precos_list) == 0:
                precos_list = None
                print("⚠️  Nenhum preço válido informado. Buscando todos os preços.")
        
        # Solicita mês e ano
        ano = int(input("Digite o ano (ex: 2025): "))
        mes = int(input("Digite o mês (1-12): "))
        
        if mes < 1 or mes > 12:
            print("❌ Mês inválido. Deve ser entre 1 e 12.")
            exit(1)
        
        print()
        print(f"🔍 Buscando vendas dos produtos: {codprod_list}")
        if precos_list:
            precos_formatados = [f"R$ {p:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.') for p in precos_list]
            print(f"   Preços filtrados: {', '.join(precos_formatados)}")
        else:
            print(f"   Preços: Todos")
        print(f"   Período: {mes:02d}/{ano}")
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
        query = get_query_vendas(mes, ano, codprod_list, precos_list)
        df = pd.read_sql_query(query, conn)
        
        if df.empty:
            print("⚠️  Nenhuma venda encontrada com os critérios especificados.")
            exit(0)
        
        print(f"✅ {len(df)} registro(s) encontrado(s).\n")
        print("=" * 140)
        print(f"{'NF':<12} {'DATA':<12} {'PRODUTO':<40} {'PREÇO':<12} {'QTD':<10} {'VALOR TOTAL':<15} {'CLIENTE':<40} {'VENDEDOR':<25}")
        print("=" * 140)
        
        # Mostra todos os registros
        for idx, row in df.iterrows():
            nf = str(row['NF'])[:12] if pd.notna(row['NF']) else 'N/A'
            data = row['DATA'].strftime('%d/%m/%Y') if pd.notna(row['DATA']) else 'N/A'
            produto = str(row['PRODUTO'])[:38]
            preco = f"R$ {row['PRECO']:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
            qt = f"{row['QUANTIDADE']:.0f}" if pd.notna(row['QUANTIDADE']) else '0'
            valor = f"R$ {row['VALOR_TOTAL']:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
            cliente = str(row['CLIENTE'])[:38]
            vendedor = str(row['VENDEDOR'])[:23] if pd.notna(row['VENDEDOR']) else 'N/A'
            
            print(f"{nf:<12} {data:<12} {produto:<40} {preco:<12} {qt:<10} {valor:<15} {cliente:<40} {vendedor:<25}")
        
        print("=" * 140)
        print(f"\n📊 RESUMO:")
        print(f"   Total de registros: {len(df)}")
        print(f"   Notas fiscais únicas: {df['NF'].nunique()}")
        print(f"   Clientes únicos: {df['CODCLI'].nunique()}")
        print(f"   Valor total: R$ {df['VALOR_TOTAL'].sum():,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
        print(f"   Quantidade total: {df['QUANTIDADE'].sum():,.0f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
        print()
        
        # Agrupa por NF
        print("=" * 120)
        print("📋 RESUMO POR NOTA FISCAL:")
        print("=" * 120)
        print(f"{'NF':<12} {'DATA':<12} {'QTD ITENS':<12} {'VALOR TOTAL':<15} {'CLIENTE':<50}")
        print("=" * 120)
        
        resumo_nf = df.groupby(['NF', 'DATA', 'CLIENTE']).agg({
            'QUANTIDADE': 'sum',
            'VALOR_TOTAL': 'sum'
        }).reset_index()
        resumo_nf = resumo_nf.sort_values('VALOR_TOTAL', ascending=False)
        
        for idx, row in resumo_nf.iterrows():
            nf = str(row['NF'])[:12] if pd.notna(row['NF']) else 'N/A'
            data = row['DATA'].strftime('%d/%m/%Y') if pd.notna(row['DATA']) else 'N/A'
            qtd_itens = f"{row['QUANTIDADE']:.0f}".replace(',', 'X').replace('.', ',').replace('X', '.')
            valor_total = f"R$ {row['VALOR_TOTAL']:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
            cliente = str(row['CLIENTE'])[:48]
            
            print(f"{nf:<12} {data:<12} {qtd_itens:<12} {valor_total:<15} {cliente:<50}")
        
        print("=" * 120)
        print()
        
        # Salva em Excel formatado
        codprod_nome = '_'.join(map(str, codprod_list))
        if precos_list:
            precos_nome = '_'.join([f"{p:.2f}".replace('.', 'p') for p in precos_list])
            arquivo_excel = f"vendas_produtos_{codprod_nome}_precos_{precos_nome}_{mes:02d}_{ano}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        else:
            arquivo_excel = f"vendas_produtos_{codprod_nome}_{mes:02d}_{ano}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
        
        with pd.ExcelWriter(arquivo_excel, engine='openpyxl') as writer:
            # Aba 1: Detalhado
            df.to_excel(writer, sheet_name='Detalhado', index=False)
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
            ws_detalhado.column_dimensions['A'].width = 12  # NF
            ws_detalhado.column_dimensions['B'].width = 12  # DATA
            ws_detalhado.column_dimensions['C'].width = 10  # CODPROD
            ws_detalhado.column_dimensions['D'].width = 35  # PRODUTO
            ws_detalhado.column_dimensions['E'].width = 12  # PRECO
            ws_detalhado.column_dimensions['F'].width = 10  # QUANTIDADE
            ws_detalhado.column_dimensions['G'].width = 15  # VALOR_TOTAL
            ws_detalhado.column_dimensions['H'].width = 10  # CODCLI
            ws_detalhado.column_dimensions['I'].width = 40  # CLIENTE
            ws_detalhado.column_dimensions['J'].width = 10  # CODUSUR
            ws_detalhado.column_dimensions['K'].width = 25  # VENDEDOR
            
            # Formata dados
            for row in ws_detalhado.iter_rows(min_row=2, max_row=ws_detalhado.max_row):
                for cell in row:
                    cell.border = border
                    if cell.column == 5:  # PRECO
                        cell.number_format = 'R$ #,##0.00'
                    elif cell.column == 7:  # VALOR_TOTAL
                        cell.number_format = 'R$ #,##0.00'
                    elif cell.column == 2:  # DATA
                        cell.number_format = 'DD/MM/YYYY'
            
            # Aba 2: Resumo por NF
            resumo_nf_excel = resumo_nf.rename(columns={'QUANTIDADE': 'QTD_TOTAL'})
            resumo_nf_excel.to_excel(writer, sheet_name='Resumo por NF', index=False)
            ws_resumo = writer.sheets['Resumo por NF']
            
            # Formata cabeçalho do resumo
            for cell in ws_resumo[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.border = border
            
            # Ajusta largura das colunas do resumo
            ws_resumo.column_dimensions['A'].width = 12  # NF
            ws_resumo.column_dimensions['B'].width = 12  # DATA
            ws_resumo.column_dimensions['C'].width = 12  # QTD_TOTAL
            ws_resumo.column_dimensions['D'].width = 15  # VALOR_TOTAL
            ws_resumo.column_dimensions['E'].width = 40  # CLIENTE
            
            # Formata dados do resumo
            for row in ws_resumo.iter_rows(min_row=2, max_row=ws_resumo.max_row):
                for cell in row:
                    cell.border = border
                    if cell.column == 4:  # VALOR_TOTAL
                        cell.number_format = 'R$ #,##0.00'
                    elif cell.column == 2:  # DATA
                        cell.number_format = 'DD/MM/YYYY'
            
            # Congela primeira linha
            ws_detalhado.freeze_panes = 'A2'
            ws_resumo.freeze_panes = 'A2'
        
        print(f"💾 Dados salvos em Excel: {arquivo_excel}")
        
    except Exception as e:
        print(f"❌ Erro ao executar query: {e}")
        import traceback
        traceback.print_exc()
    finally:
        conn.close()
        print("\n🔌 Conexão fechada.")




# portfolio-commit-ready: analise_vendas_produtos_universal
