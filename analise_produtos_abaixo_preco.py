"""
Sistema de Análise de Produtos Vendidos Abaixo do Custo
Gera Excel com duas abas interligadas:
- Aba 1: Produtos vendidos abaixo do custo no mês selecionado
- Aba 2: Histórico completo de produtos vendidos abaixo do custo
"""

import tkinter as tk
from tkinter import ttk, messagebox
import oracledb
import pandas as pd
from datetime import datetime
from calendar import monthrange
from pathlib import Path
import os
import threading
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.formatting.rule import ColorScaleRule, FormulaRule

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
# FUNÇÕES DE BANCO
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
        return None

def buscar_produtos_abaixo_preco_mes(conn, mes, ano):
    """Busca produtos vendidos abaixo do custo no mês especificado
    Compara preço de venda com custo do produto
    """
    data_inicio = datetime(ano, mes, 1)
    if mes == 12:
        data_fim = datetime(ano + 1, 1, 1)
    else:
        data_fim = datetime(ano, mes + 1, 1)
    
    data_inicio_str = data_inicio.strftime('%d/%m/%Y')
    data_fim_str = data_fim.strftime('%d/%m/%Y')
    
    query = f"""
    WITH VENDAS_MES AS (
        SELECT 
            I.CODPROD,
            P.DESCRICAO AS PRODUTO,
            P.EMBALAGEM,
            P.UNIDADE,
            P.PVENDA AS PRECO_TABELA,
            E.CUSTOFIN AS CUSTO,
            MIN(I.PVENDA) AS MENOR_PRECO_VENDA,
            MAX(I.PVENDA) AS MAIOR_PRECO_VENDA,
            AVG(I.PVENDA) AS PRECO_MEDIO_VENDA,
            SUM(
                CASE 
                    WHEN NVL(I.BONIFIC, 'N') = 'N' THEN 
                        DECODE(C.CONDVENDA, 
                            5, 0, 6, 0, 11, 0, 12, 0, 
                            NVL(I.QT, 0)
                        )
                    ELSE 0 
                END
            ) AS QUANTIDADE_VENDIDA,
            COUNT(DISTINCT C.NUMPED) AS NUM_PEDIDOS,
            COUNT(DISTINCT C.CODCLI) AS NUM_CLIENTES
        FROM PCPEDC C 
        JOIN PCPEDI I ON C.NUMPED = I.NUMPED
        JOIN PCPRODUT P ON I.CODPROD = P.CODPROD
        LEFT JOIN PCEST E ON P.CODPROD = E.CODPROD AND E.CODFILIAL = '1'
        WHERE C.DATA >= TO_DATE('{data_inicio_str}', 'DD/MM/YYYY')
          AND C.DATA < TO_DATE('{data_fim_str}', 'DD/MM/YYYY')
          AND C.CODFILIAL IN ('1', '98')
          AND C.POSICAO = 'F'
          AND C.DTCANCEL IS NULL
          AND C.CONDVENDA NOT IN (4, 5, 6, 8, 10, 11, 12, 13, 20, 98, 99)
          AND NVL(I.BONIFIC, 'N') = 'N'
          AND I.PVENDA < NVL(E.CUSTOFIN, 0)  -- Vendido abaixo do custo
        GROUP BY I.CODPROD, P.DESCRICAO, P.EMBALAGEM, P.UNIDADE, P.PVENDA, E.CUSTOFIN
    ),
    HISTORICO_MESES AS (
        SELECT 
            I.CODPROD,
            COUNT(DISTINCT TO_CHAR(C.DATA, 'MM/YYYY')) AS MESES_VENDIDO_ABAIXO
        FROM PCPEDC C 
        JOIN PCPEDI I ON C.NUMPED = I.NUMPED
        JOIN PCPRODUT P ON I.CODPROD = P.CODPROD
        LEFT JOIN PCEST E ON P.CODPROD = E.CODPROD AND E.CODFILIAL = '1'
        WHERE C.CODFILIAL IN ('1', '98')
          AND C.POSICAO = 'F'
          AND C.DTCANCEL IS NULL
          AND C.CONDVENDA NOT IN (4, 5, 6, 8, 10, 11, 12, 13, 20, 98, 99)
          AND NVL(I.BONIFIC, 'N') = 'N'
          AND I.PVENDA < NVL(E.CUSTOFIN, 0)  -- Vendido abaixo do custo
        GROUP BY I.CODPROD
    )
    SELECT 
        V.CODPROD,
        V.PRODUTO,
        V.CUSTO,
        V.PRECO_MEDIO_VENDA AS PRECO_VENDA,
        V.QUANTIDADE_VENDIDA,
        (V.PRECO_MEDIO_VENDA * V.QUANTIDADE_VENDIDA) AS VALOR_TOTAL_VENDIDO,
        ((V.CUSTO - V.PRECO_MEDIO_VENDA) * V.QUANTIDADE_VENDIDA) AS PREJUIZO_TOTAL,
        NVL(H.MESES_VENDIDO_ABAIXO, 0) AS MESES_HISTORICO
    FROM VENDAS_MES V
    LEFT JOIN HISTORICO_MESES H ON V.CODPROD = H.CODPROD
    WHERE V.QUANTIDADE_VENDIDA > 0
    ORDER BY MESES_HISTORICO DESC, PREJUIZO_TOTAL DESC
    """
    
    try:
        df = pd.read_sql_query(query, conn)
        df.columns = [x.lower() for x in df.columns]
        return df
    except Exception as e:
        print(f"Erro ao buscar produtos abaixo do custo: {e}")
        import traceback
        traceback.print_exc()
        return pd.DataFrame()

def buscar_historico_completo(conn):
    """Busca histórico completo de todos os produtos vendidos abaixo do custo"""
    query = """
    WITH VENDAS_HISTORICO AS (
        SELECT 
            I.CODPROD,
            P.DESCRICAO AS PRODUTO,
            E.CUSTOFIN AS CUSTO_ATUAL,
            MIN(I.PVENDA) AS MENOR_PRECO_VENDA,
            MAX(I.PVENDA) AS MAIOR_PRECO_VENDA,
            AVG(I.PVENDA) AS PRECO_MEDIO_VENDA,
            COUNT(DISTINCT TO_CHAR(C.DATA, 'MM/YYYY')) AS MESES_VENDIDO_ABAIXO,
            MIN(C.DATA) AS PRIMEIRA_VENDA_ABAIXO,
            MAX(C.DATA) AS ULTIMA_VENDA_ABAIXO,
            SUM(
                CASE 
                    WHEN NVL(I.BONIFIC, 'N') = 'N' THEN 
                        DECODE(C.CONDVENDA, 
                            5, 0, 6, 0, 11, 0, 12, 0, 
                            NVL(I.QT, 0)
                        )
                    ELSE 0 
                END
            ) AS QUANTIDADE_TOTAL_VENDIDA,
            SUM(
                CASE 
                    WHEN NVL(I.BONIFIC, 'N') = 'N' THEN 
                        DECODE(C.CONDVENDA, 
                            5, 0, 6, 0, 11, 0, 12, 0, 
                            ROUND(NVL(I.QT, 0) * NVL(I.PVENDA, 0), 2)
                        )
                    ELSE 0 
                END
            ) AS VALOR_TOTAL_VENDIDO,
            SUM(
                CASE 
                    WHEN NVL(I.BONIFIC, 'N') = 'N' THEN 
                        DECODE(C.CONDVENDA, 
                            5, 0, 6, 0, 11, 0, 12, 0, 
                            (NVL(E.CUSTOFIN, 0) - NVL(I.PVENDA, 0)) * NVL(I.QT, 0)
                        )
                    ELSE 0 
                END
            ) AS PREJUIZO_TOTAL
        FROM PCPEDC C 
        JOIN PCPEDI I ON C.NUMPED = I.NUMPED
        JOIN PCPRODUT P ON I.CODPROD = P.CODPROD
        LEFT JOIN PCEST E ON P.CODPROD = E.CODPROD AND E.CODFILIAL = '1'
        WHERE C.CODFILIAL IN ('1', '98')
          AND C.POSICAO = 'F'
          AND C.DTCANCEL IS NULL
          AND C.CONDVENDA NOT IN (4, 5, 6, 8, 10, 11, 12, 13, 20, 98, 99)
          AND NVL(I.BONIFIC, 'N') = 'N'
          AND I.PVENDA < NVL(E.CUSTOFIN, 0)  -- Vendido abaixo do custo
        GROUP BY I.CODPROD, P.DESCRICAO, E.CUSTOFIN
    )
    SELECT 
        CODPROD,
        PRODUTO,
        CUSTO_ATUAL,
        MENOR_PRECO_VENDA,
        MAIOR_PRECO_VENDA,
        PRECO_MEDIO_VENDA,
        MESES_VENDIDO_ABAIXO,
        PRIMEIRA_VENDA_ABAIXO,
        ULTIMA_VENDA_ABAIXO,
        QUANTIDADE_TOTAL_VENDIDA,
        VALOR_TOTAL_VENDIDO,
        PREJUIZO_TOTAL
    FROM VENDAS_HISTORICO
    WHERE QUANTIDADE_TOTAL_VENDIDA > 0
    ORDER BY MESES_VENDIDO_ABAIXO DESC, PREJUIZO_TOTAL DESC
    """
    
    try:
        df = pd.read_sql_query(query, conn)
        df.columns = [x.lower() for x in df.columns]
        return df
    except Exception as e:
        print(f"Erro ao buscar histórico completo: {e}")
        import traceback
        traceback.print_exc()
        return pd.DataFrame()

# =========================================================================
# FUNÇÕES DE EXCEL
# =========================================================================
def formatar_aba_mes(ws, mes, ano):
    """Formata a aba do mês selecionado"""
    # Estilos
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    border_style = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Adicionar título
    ws.insert_rows(1)
    meses_nomes = {
        1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril',
        5: 'Maio', 6: 'Junho', 7: 'Julho', 8: 'Agosto',
        9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'
    }
    mes_nome = meses_nomes[mes]
    num_cols = ws.max_column
    ws.merge_cells(f'A1:{get_column_letter(num_cols)}1')
    titulo_cell = ws['A1']
    titulo_cell.value = f"Produtos Vendidos Abaixo do Custo - {mes_nome}/{ano}"
    titulo_cell.font = Font(bold=True, size=14, color="366092")
    titulo_cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # Formatar cabeçalho (linha 2 após título)
    linha_cabecalho = 2
    for cell in ws[linha_cabecalho]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border_style
    
    # Formatar dados - SEM CORES, apenas formatação de números
    no_fill = PatternFill()
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row):
        for cell in row:
            cell.fill = no_fill  # SEM COR
            cell.border = border_style
            
            if cell.column == 1:  # Código
                cell.alignment = Alignment(horizontal='center', vertical='center')
            elif cell.column == 2:  # Produto
                cell.alignment = Alignment(horizontal='left', vertical='center')
            elif cell.column in [3, 4, 6, 7]:  # Custo, Preço Venda, Valor Total, Prejuízo Total
                cell.alignment = Alignment(horizontal='right', vertical='center')
                cell.number_format = '#,##0.00'
            elif cell.column == 5:  # Quantidade
                cell.alignment = Alignment(horizontal='right', vertical='center')
                cell.number_format = '#,##0'
            elif cell.column == 8:  # Meses Histórico
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.number_format = '#,##0'
    
    # Ajustar larguras
    ws.column_dimensions['A'].width = 12  # Código
    ws.column_dimensions['B'].width = 50  # Produto
    ws.column_dimensions['C'].width = 15  # Custo
    ws.column_dimensions['D'].width = 15  # Preço Venda
    ws.column_dimensions['E'].width = 15  # Quantidade
    ws.column_dimensions['F'].width = 18  # Valor Total Vendido
    ws.column_dimensions['G'].width = 18  # Prejuízo Total
    ws.column_dimensions['H'].width = 15  # Meses Histórico
    
    # Adicionar linha de totais
    if ws.max_row > 2:
        linha_total = ws.max_row + 1
        ws[f'A{linha_total}'] = 'TOTAL'
        ws[f'A{linha_total}'].font = Font(bold=True)
        ws[f'A{linha_total}'].alignment = Alignment(horizontal='right', vertical='center')
        
        # Fórmulas de total
        ws[f'F{linha_total}'] = f'=SUM(F3:F{ws.max_row})'  # Valor Total Vendido
        ws[f'G{linha_total}'] = f'=SUM(G3:G{ws.max_row})'  # Prejuízo Total
        # Meses Histórico não tem total (é informação histórica)
        
        # Formatar células de totais
        for col in ['F', 'G']:
            cell = ws[f'{col}{linha_total}']
            cell.font = Font(bold=True)
            cell.alignment = Alignment(horizontal='right', vertical='center')
            cell.number_format = '#,##0.00'
            cell.border = border_style
            cell.fill = PatternFill(start_color="E7E6E6", end_color="E7E6E6", fill_type="solid")  # Cinza claro apenas no total

def adicionar_hyperlinks_meses_historico(ws_mes, ws_historico):
    """Adiciona hyperlinks na coluna Meses Histórico que levam para a aba de detalhes"""
    # Coluna H = Meses Histórico (coluna 8)
    col_meses = 8
    
    # Criar dicionário com código do produto -> linha na aba histórico
    codprod_linha = {}
    for row_num in range(3, ws_historico.max_row + 1):
        codprod = ws_historico.cell(row=row_num, column=1).value
        if codprod:
            codprod_linha[codprod] = row_num
    
    # Adicionar hyperlinks na coluna Meses Histórico (exceto linha de totais)
    max_row = ws_mes.max_row
    if ws_mes[f'A{max_row}'].value == 'TOTAL':
        max_row = max_row - 1
    
    for row_num in range(3, max_row + 1):
        codprod = ws_mes.cell(row=row_num, column=1).value  # Coluna A = Código
        cell_meses = ws_mes.cell(row=row_num, column=col_meses)
        
        if codprod and codprod in codprod_linha:
            linha_historico = codprod_linha[codprod]
            # Criar hyperlink para a aba de detalhes
            hyperlink = f"#'Detalhes Histórico'!A{linha_historico}"
            cell_meses.hyperlink = hyperlink
            cell_meses.font = Font(color="0000FF", underline="single")  # Azul e sublinhado
            # Manter o valor original (número de meses)
            if cell_meses.value is None:
                cell_meses.value = 0

def formatar_aba_detalhes_historico(ws):
    """Formata a aba de detalhes histórico"""
    # Estilos
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    border_style = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Adicionar título
    ws.insert_rows(1)
    num_cols = ws.max_column
    ws.merge_cells(f'A1:{get_column_letter(num_cols)}1')
    titulo_cell = ws['A1']
    titulo_cell.value = "Detalhes Histórico - Produtos Vendidos Abaixo do Custo"
    titulo_cell.font = Font(bold=True, size=14, color="366092")
    titulo_cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # Formatar cabeçalho (linha 2 após título)
    linha_cabecalho = 2
    for cell in ws[linha_cabecalho]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border_style
    
    # Formatar dados - SEM CORES
    no_fill = PatternFill()
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row):
        for cell in row:
            cell.fill = no_fill
            cell.border = border_style
            
            if cell.column == 1:  # Código
                cell.alignment = Alignment(horizontal='center', vertical='center')
            elif cell.column == 2:  # Produto
                cell.alignment = Alignment(horizontal='left', vertical='center')
            elif cell.column in [3, 4, 5, 6, 10, 11, 12]:  # Valores monetários
                cell.alignment = Alignment(horizontal='right', vertical='center')
                cell.number_format = '#,##0.00'
            elif cell.column == 7:  # Meses Vendido Abaixo
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.number_format = '#,##0'
            elif cell.column in [8, 9]:  # Datas
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.number_format = 'DD/MM/YYYY'
    
    # Ajustar larguras
    ws.column_dimensions['A'].width = 12  # Código
    ws.column_dimensions['B'].width = 50  # Produto
    ws.column_dimensions['C'].width = 15  # Custo Atual
    ws.column_dimensions['D'].width = 15  # Menor Preço
    ws.column_dimensions['E'].width = 15  # Maior Preço
    ws.column_dimensions['F'].width = 15  # Preço Médio
    ws.column_dimensions['G'].width = 15  # Meses Vendido Abaixo
    ws.column_dimensions['H'].width = 15  # Primeira Venda
    ws.column_dimensions['I'].width = 15  # Última Venda
    ws.column_dimensions['J'].width = 15  # Quantidade Total
    ws.column_dimensions['K'].width = 18  # Valor Total Vendido
    ws.column_dimensions['L'].width = 18  # Prejuízo Total

def formatar_aba_historico(ws):
    """Formata a aba de histórico completo"""
    # Estilos
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    border_style = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin')
    )
    
    # Adicionar título
    ws.insert_rows(1)
    num_cols = ws.max_column
    ws.merge_cells(f'A1:{get_column_letter(num_cols)}1')
    titulo_cell = ws['A1']
    titulo_cell.value = "Histórico Completo - Produtos Vendidos Abaixo do Custo"
    titulo_cell.font = Font(bold=True, size=14, color="366092")
    titulo_cell.alignment = Alignment(horizontal='center', vertical='center')
    
    # Formatar cabeçalho (linha 2 após título)
    linha_cabecalho = 2
    for cell in ws[linha_cabecalho]:
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal='center', vertical='center', wrap_text=True)
        cell.border = border_style
    
    # Formatar dados - SEM CORES, apenas formatação de números
    no_fill = PatternFill()
    for row in ws.iter_rows(min_row=3, max_row=ws.max_row):
        for cell in row:
            cell.fill = no_fill  # SEM COR
            cell.border = border_style
            
            if cell.column == 1:  # Código
                cell.alignment = Alignment(horizontal='center', vertical='center')
            elif cell.column == 2:  # Produto
                cell.alignment = Alignment(horizontal='left', vertical='center')
            elif cell.column in [3, 4, 5, 6, 7, 8, 9, 10, 11, 17]:  # Valores monetários
                cell.alignment = Alignment(horizontal='right', vertical='center')
                cell.number_format = '#,##0.00'
            elif cell.column == 12:  # Prejuízo Máximo Percentual
                cell.alignment = Alignment(horizontal='right', vertical='center')
                cell.number_format = '0.00%'
            elif cell.column == 13:  # Meses Vendido Abaixo
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.number_format = '#,##0'
            elif cell.column in [14, 15]:  # Datas
                cell.alignment = Alignment(horizontal='center', vertical='center')
                cell.number_format = 'DD/MM/YYYY'
            elif cell.column in [16, 18, 19, 20]:  # Quantidades e números
                cell.alignment = Alignment(horizontal='right', vertical='center')
                cell.number_format = '#,##0'
            elif cell.column == 21:  # Margem Percentual Mín Histórica
                cell.alignment = Alignment(horizontal='right', vertical='center')
                cell.number_format = '0.00%'
    
    # Ajustar larguras
    ws.column_dimensions['A'].width = 12  # Código
    ws.column_dimensions['B'].width = 40  # Produto
    ws.column_dimensions['C'].width = 12  # Embalagem
    ws.column_dimensions['D'].width = 10  # Unidade
    ws.column_dimensions['E'].width = 15  # Preço Tabela Atual
    ws.column_dimensions['F'].width = 15  # Custo Atual
    ws.column_dimensions['G'].width = 15  # Menor Preço Histórico
    ws.column_dimensions['H'].width = 15  # Maior Preço Histórico
    ws.column_dimensions['I'].width = 15  # Preço Médio Histórico
    ws.column_dimensions['J'].width = 15  # Prejuízo Máx Absoluto
    ws.column_dimensions['K'].width = 15  # Prejuízo Máx %
    ws.column_dimensions['L'].width = 12  # Meses Vendido Abaixo
    ws.column_dimensions['M'].width = 15  # Primeira Venda
    ws.column_dimensions['N'].width = 15  # Última Venda
    ws.column_dimensions['O'].width = 15  # Quantidade Total
    ws.column_dimensions['P'].width = 12  # Total Pedidos
    ws.column_dimensions['Q'].width = 12  # Total Clientes
    ws.column_dimensions['R'].width = 12  # Total Vendedores
    ws.column_dimensions['S'].width = 15  # Margem Mín Histórica
    ws.column_dimensions['T'].width = 15  # Margem % Mín Histórica


# =========================================================================
# CLASSE PRINCIPAL
# =========================================================================
class AnaliseProdutosAbaixoPreco:
    def __init__(self, root):
        self.root = root
        self.root.title("📊 Análise: Produtos Vendidos Abaixo do Custo")
        self.root.geometry("500x400")
        
        # Variáveis
        self.conn = None
        
        # Criar interface
        self.criar_interface()
        
        # Conectar ao banco
        self.conectar_banco()
    
    def criar_interface(self):
        """Cria a interface gráfica"""
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Título
        titulo = ttk.Label(main_frame, text="📊 Análise: Produtos Abaixo do Custo", 
                          font=("Arial", 16, "bold"))
        titulo.pack(pady=(0, 20))
        
        # Frame de mês
        frame_mes = ttk.LabelFrame(main_frame, text="Selecionar Mês para Análise", padding="15")
        frame_mes.pack(fill=tk.X, pady=10)
        
        # Mês
        ttk.Label(frame_mes, text="Mês:").pack(anchor=tk.W, pady=(0, 5))
        frame_mes_input = ttk.Frame(frame_mes)
        frame_mes_input.pack(fill=tk.X, pady=(0, 10))
        
        self.mes_var = tk.IntVar(value=datetime.now().month)
        meses = list(range(1, 13))
        meses_nomes = ['Janeiro', 'Fevereiro', 'Março', 'Abril', 'Maio', 'Junho',
                      'Julho', 'Agosto', 'Setembro', 'Outubro', 'Novembro', 'Dezembro']
        self.combo_mes = ttk.Combobox(frame_mes_input, textvariable=self.mes_var, 
                                     values=[f"{m:02d} - {n}" for m, n in zip(meses, meses_nomes)],
                                     state="readonly", width=30)
        self.combo_mes.current(datetime.now().month - 1)
        self.combo_mes.pack(side=tk.LEFT, padx=2)
        
        # Ano
        ttk.Label(frame_mes, text="Ano:").pack(anchor=tk.W, pady=(0, 5))
        frame_ano_input = ttk.Frame(frame_mes)
        frame_ano_input.pack(fill=tk.X)
        
        self.ano_var = tk.IntVar(value=datetime.now().year)
        anos = list(range(2020, datetime.now().year + 2))
        self.combo_ano = ttk.Combobox(frame_ano_input, textvariable=self.ano_var,
                                     values=anos, state="readonly", width=30)
        self.combo_ano.current(len(anos) - 1)
        self.combo_ano.pack(side=tk.LEFT, padx=2)
        
        # Botão gerar Excel
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=20)
        
        self.btn_gerar = ttk.Button(btn_frame, text="📄 Gerar Excel", 
                                   command=self.gerar_excel, state=tk.DISABLED)
        self.btn_gerar.pack(side=tk.LEFT, padx=(0, 10))
        
        self.status_label = ttk.Label(btn_frame, text="Conectando ao banco...")
        self.status_label.pack(side=tk.LEFT)
    
    def conectar_banco(self):
        """Conecta ao banco"""
        def conectar():
            try:
                self.status_label.config(text="Conectando ao banco...")
                self.conn = get_db_connection()
                
                if not self.conn:
                    self.status_label.config(text="❌ Erro ao conectar")
                    messagebox.showerror("Erro", "Não foi possível conectar ao banco de dados")
                    return
                
                self.status_label.config(text="✅ Pronto para gerar relatório")
                self.btn_gerar.config(state=tk.NORMAL)
                
            except Exception as e:
                self.status_label.config(text=f"❌ Erro: {str(e)}")
                messagebox.showerror("Erro", f"Erro ao conectar: {e}")
        
        threading.Thread(target=conectar, daemon=True).start()
    
    def gerar_excel(self):
        """Gera o Excel com análise de produtos abaixo do custo"""
        if not self.conn:
            messagebox.showerror("Erro", "Não há conexão com o banco de dados")
            return
        
        # Obter mês e ano selecionados
        mes_ano_str = self.combo_mes.get()
        mes = int(mes_ano_str.split(' - ')[0])
        ano = int(self.combo_ano.get())
        
        def executar():
            try:
                self.status_label.config(text="Buscando dados do mês...")
                self.root.update()
                
                # Buscar produtos abaixo do custo do mês
                df_mes = buscar_produtos_abaixo_preco_mes(self.conn, mes, ano)
                
                if df_mes.empty:
                    messagebox.showinfo("Informação", "Nenhum produto vendido abaixo do custo encontrado.")
                    self.status_label.config(text="Nenhum produto encontrado")
                    return
                
                self.status_label.config(text="Buscando histórico completo...")
                self.root.update()
                
                # Buscar histórico completo
                df_historico = buscar_historico_completo(self.conn)
                
                self.status_label.config(text="Gerando Excel...")
                self.root.update()
                
                # Preparar dados para Excel - Aba Mês
                if not df_mes.empty:
                    df_mes_excel = df_mes[[
                        'codprod', 'produto', 'custo', 'preco_venda',
                        'quantidade_vendida', 'valor_total_vendido', 'prejuizo_total', 'meses_historico'
                    ]].copy()
                    
                    df_mes_excel.columns = [
                        'Código', 'Produto', 'Custo', 'Preço Venda',
                        'Quantidade', 'Valor Total Vendido', 'Prejuízo Total', 'Meses Histórico'
                    ]
                else:
                    df_mes_excel = pd.DataFrame(columns=[
                        'Código', 'Produto', 'Custo', 'Preço Venda',
                        'Quantidade', 'Valor Total Vendido', 'Prejuízo Total', 'Meses Histórico'
                    ])
                
                # Nome do arquivo
                meses_nomes = {
                    1: 'Jan', 2: 'Fev', 3: 'Mar', 4: 'Abr',
                    5: 'Mai', 6: 'Jun', 7: 'Jul', 8: 'Ago',
                    9: 'Set', 10: 'Out', 11: 'Nov', 12: 'Dez'
                }
                mes_abrev = meses_nomes[mes]
                nome_arquivo = f"produtos_abaixo_preco_{mes_abrev}{str(ano)[-2:]}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                caminho_arquivo = Path(nome_arquivo).resolve()
                
                # Preparar dados para Excel - Aba Histórico
                if not df_historico.empty:
                    df_historico_excel = df_historico[[
                        'codprod', 'produto', 'custo_atual',
                        'menor_preco_venda', 'maior_preco_venda', 'preco_medio_venda',
                        'meses_vendido_abaixo', 'primeira_venda_abaixo', 'ultima_venda_abaixo',
                        'quantidade_total_vendida', 'valor_total_vendido', 'prejuizo_total'
                    ]].copy()
                    
                    df_historico_excel.columns = [
                        'Código', 'Produto', 'Custo Atual',
                        'Menor Preço', 'Maior Preço', 'Preço Médio',
                        'Meses Vendido Abaixo', 'Primeira Venda', 'Última Venda',
                        'Quantidade Total', 'Valor Total Vendido', 'Prejuízo Total'
                    ]
                else:
                    df_historico_excel = pd.DataFrame(columns=[
                        'Código', 'Produto', 'Custo Atual',
                        'Menor Preço', 'Maior Preço', 'Preço Médio',
                        'Meses Vendido Abaixo', 'Primeira Venda', 'Última Venda',
                        'Quantidade Total', 'Valor Total Vendido', 'Prejuízo Total'
                    ])
                
                # Criar Excel com duas abas
                with pd.ExcelWriter(caminho_arquivo, engine='openpyxl') as writer:
                    df_mes_excel.to_excel(writer, sheet_name=f'{mes_abrev}-{str(ano)[-2:]}', index=False)
                    df_historico_excel.to_excel(writer, sheet_name='Detalhes Histórico', index=False)
                
                # Formatar Excel
                wb = openpyxl.load_workbook(caminho_arquivo)
                ws_mes = wb[f'{mes_abrev}-{str(ano)[-2:]}']
                ws_historico = wb['Detalhes Histórico']
                
                formatar_aba_mes(ws_mes, mes, ano)
                formatar_aba_detalhes_historico(ws_historico)
                
                # Adicionar hyperlinks na coluna Meses Histórico
                adicionar_hyperlinks_meses_historico(ws_mes, ws_historico)
                
                wb.save(caminho_arquivo)
                
                self.status_label.config(text="✅ Excel gerado com sucesso!")
                messagebox.showinfo("Sucesso", 
                                  f"Excel gerado com sucesso!\n\nArquivo: {caminho_arquivo}\n\n"
                                  f"Produtos no mês: {len(df_mes_excel)}")
                
                # Abrir Excel
                try:
                    os.startfile(str(caminho_arquivo))
                except:
                    pass
                    
            except Exception as e:
                messagebox.showerror("Erro", f"Erro ao gerar Excel:\n{e}")
                self.status_label.config(text="❌ Erro ao gerar")
                import traceback
                traceback.print_exc()
        
        threading.Thread(target=executar, daemon=True).start()

# =========================================================================
# EXECUÇÃO
# =========================================================================
if __name__ == "__main__":
    root = tk.Tk()
    app = AnaliseProdutosAbaixoPreco(root)
    root.mainloop()

# portfolio-commit-ready: analise_produtos_abaixo_preco
