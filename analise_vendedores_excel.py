"""
Sistema de Análise de Vendedores - Geração de Excel Simplificado
Interface Tkinter para gerar análise simplificada em Excel
"""

import tkinter as tk
from tkinter import ttk, messagebox
import oracledb
import pandas as pd
from datetime import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from pathlib import Path
import os
import threading

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
    except Exception as e:
        return None

def buscar_vendedores_mes(_conn, mes, ano, coddepto=None, rcas_list=None, supervisores_list=None):
    """Busca dados de vendedores para o mês especificado"""
    data_inicio = datetime(ano, mes, 1)
    if mes == 12:
        data_fim = datetime(ano + 1, 1, 1)
    else:
        data_fim = datetime(ano, mes + 1, 1)
    
    data_inicio_str = data_inicio.strftime('%d/%m/%Y')
    data_fim_str = data_fim.strftime('%d/%m/%Y')
    
    # Filtro de departamento
    filtro_depto = ""
    if coddepto:
        filtro_depto = f"AND P.CODEPTO = {coddepto}"
    
    # Filtro de supervisores
    filtro_supervisor = ""
    if supervisores_list:
        supervisores_str = ', '.join(map(str, supervisores_list))
        filtro_supervisor = f"AND U.CODSUPERVISOR IN ({supervisores_str})"
    
    # Filtro de RCAs
    filtro_rca = ""
    if rcas_list:
        chunks = [rcas_list[i:i+1000] for i in range(0, len(rcas_list), 1000)]
        if len(chunks) == 1:
            filtro_rca = f"AND C.CODUSUR IN ({', '.join(map(str, rcas_list))})"
        else:
            condicoes = []
            for chunk in chunks:
                condicoes.append(f"C.CODUSUR IN ({', '.join(map(str, chunk))})")
            filtro_rca = f"AND ({' OR '.join(condicoes)})"
    
    filtro_rca_devol = ""
    if rcas_list:
        chunks = [rcas_list[i:i+1000] for i in range(0, len(rcas_list), 1000)]
        if len(chunks) == 1:
            filtro_rca_devol = f"AND D.CODUSUR IN ({', '.join(map(str, rcas_list))})"
        else:
            condicoes = []
            for chunk in chunks:
                condicoes.append(f"D.CODUSUR IN ({', '.join(map(str, chunk))})")
            filtro_rca_devol = f"AND ({' OR '.join(condicoes)})"
    
    query = f"""
    WITH VENDAS_VALIDAS AS (
        SELECT 
            C.CODUSUR,
            U.NOME AS NOME_VENDEDOR,
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
            ) AS CUSTO_BRUTO,
            COUNT(DISTINCT I.CODPROD) AS QTD_PRODUTOS_DIFERENTES
        FROM PCPEDC C 
        JOIN PCPEDI I ON C.NUMPED = I.NUMPED
        JOIN PCPRODUT P ON I.CODPROD = P.CODPROD
        JOIN PCUSUARI U ON C.CODUSUR = U.CODUSUR
        WHERE C.DATA >= TO_DATE('{data_inicio_str}', 'DD/MM/YYYY')
          AND C.DATA < TO_DATE('{data_fim_str}', 'DD/MM/YYYY')
          AND C.CODFILIAL IN ('1', '98')
          AND C.POSICAO = 'F'
          AND C.DTCANCEL IS NULL
          AND C.CONDVENDA NOT IN (4, 8, 10, 13, 20, 98, 99)
          {filtro_depto}
          {filtro_rca}
          {filtro_supervisor}
        GROUP BY C.CODUSUR, U.NOME
    ),
    DEVOLUCOES AS (
        SELECT 
            D.CODUSUR,
            SUM(NVL(D.VLDEVOLUCAO, 0)) AS VALOR_DEVOLVIDO,
            SUM(NVL(D.VLCUSTOFIN, 0)) AS CUSTO_DEVOLVIDO
        FROM VIEW_DEVOL_RESUMO_FATURAMENTO D
        JOIN PCPRODUT P ON D.CODPROD = P.CODPROD
        WHERE D.DTENT >= TO_DATE('{data_inicio_str}', 'DD/MM/YYYY')
          AND D.DTENT < TO_DATE('{data_fim_str}', 'DD/MM/YYYY')
          {filtro_depto}
          {filtro_rca_devol}
        GROUP BY D.CODUSUR
    )
    SELECT 
        V.CODUSUR AS RCA,
        V.NOME_VENDEDOR,
        (NVL(V.VALOR_BRUTO, 0) - NVL(D.VALOR_DEVOLVIDO, 0)) AS FATURAMENTO,
        (NVL(V.VALOR_BRUTO, 0) - NVL(D.VALOR_DEVOLVIDO, 0)) - (NVL(V.CUSTO_BRUTO, 0) - NVL(D.CUSTO_DEVOLVIDO, 0)) AS MARGEM_VALOR,
        NVL(V.QTD_PRODUTOS_DIFERENTES, 0) AS MIX
    FROM VENDAS_VALIDAS V
    LEFT JOIN DEVOLUCOES D ON V.CODUSUR = D.CODUSUR
    WHERE V.VALOR_BRUTO > 0
    ORDER BY FATURAMENTO DESC
    """
    
    try:
        df = pd.read_sql_query(query, _conn)
        df.columns = [x.lower() for x in df.columns]
        
        # Calcular margem percentual
        if not df.empty:
            df['margem_percentual'] = df.apply(
                lambda row: 0.0 if row['faturamento'] <= 0 
                else round(((row['margem_valor'] / row['faturamento']) * 100), 2),
                axis=1
            )
            df['mes'] = mes
            df['ano'] = ano
        
        return df
    except Exception as e:
        print(f"Erro ao buscar dados: {e}")
        return pd.DataFrame()

def buscar_dados_ano_completo(_conn, ano, coddepto=None, rcas_list=None, supervisores_list=None):
    """Busca dados de todos os meses do ano até o mês anterior ao atual"""
    dados_mensais = []
    
    # Obter mês atual
    hoje = datetime.now()
    mes_atual = hoje.month
    ano_atual = hoje.year
    
    # Determinar até qual mês buscar
    if ano == ano_atual:
        if mes_atual == 1:
            mes_limite = 0
        else:
            mes_limite = mes_atual - 1
    else:
        mes_limite = 12
    
    if mes_limite == 0:
        return pd.DataFrame()
    
    for mes in range(1, mes_limite + 1):
        df_mes = buscar_vendedores_mes(_conn, mes, ano, coddepto, rcas_list, supervisores_list)
        if not df_mes.empty:
            dados_mensais.append(df_mes)
    
    if not dados_mensais:
        return pd.DataFrame()
    
    df_completo = pd.concat(dados_mensais, ignore_index=True)
    return df_completo

def calcular_resumo_vendedores(df_ano, df_mes_atual, ano):
    """Calcula resumo simplificado por vendedor"""
    if df_ano.empty:
        return pd.DataFrame()
    
    hoje = datetime.now()
    mes_atual = hoje.month
    ano_atual = hoje.year
    
    # Determinar mês passado e atual
    if ano == ano_atual:
        mes_passado = mes_atual - 1 if mes_atual > 1 else 12
        ano_mes_passado = ano if mes_atual > 1 else ano - 1
        mes_atual_num = mes_atual
    else:
        mes_passado = 12
        ano_mes_passado = ano
        mes_atual_num = None
    
    # Combinar dados do ano com mês atual se disponível
    if not df_mes_atual.empty:
        df_completo = pd.concat([df_ano, df_mes_atual], ignore_index=True)
    else:
        df_completo = df_ano.copy()
    
    vendedores_resumo = []
    
    for rca in df_completo['rca'].unique():
        df_vendedor = df_completo[df_completo['rca'] == rca].copy()
        
        if df_vendedor.empty:
            continue
        
        nome_vendedor = df_vendedor.iloc[0]['nome_vendedor']
        
        # Faturamento Anual (soma de todos os meses)
        faturamento_anual = df_vendedor['faturamento'].sum()
        margem_anual = df_vendedor['margem_valor'].sum()
        margem_anual_pct = (margem_anual / faturamento_anual * 100) if faturamento_anual > 0 else 0
        mix_anual = df_vendedor['mix'].max()  # Maior mix do ano
        
        # Média Ano (média dos meses)
        meses_com_venda = df_vendedor[df_vendedor['faturamento'] > 0]
        if not meses_com_venda.empty:
            faturamento_medio_ano = meses_com_venda['faturamento'].mean()
            margem_media_ano = meses_com_venda['margem_percentual'].mean()
            mix_medio_ano = meses_com_venda['mix'].mean()
        else:
            faturamento_medio_ano = 0
            margem_media_ano = 0
            mix_medio_ano = 0
        
        # Média Trimestre (últimos 3 meses)
        df_vendedor_sorted = df_vendedor.sort_values(['ano', 'mes'], ascending=False)
        ultimos_3_meses = df_vendedor_sorted.head(3)
        if not ultimos_3_meses.empty:
            faturamento_medio_trimestre = ultimos_3_meses['faturamento'].mean()
            margem_media_trimestre = ultimos_3_meses['margem_percentual'].mean()
            mix_medio_trimestre = ultimos_3_meses['mix'].mean()
        else:
            faturamento_medio_trimestre = 0
            margem_media_trimestre = 0
            mix_medio_trimestre = 0
        
        # Mês Passado
        df_mes_passado = df_vendedor[(df_vendedor['mes'] == mes_passado) & (df_vendedor['ano'] == ano_mes_passado)]
        if not df_mes_passado.empty:
            faturamento_mes_passado = df_mes_passado.iloc[0]['faturamento']
            margem_mes_passado = df_mes_passado.iloc[0]['margem_percentual']
            mix_mes_passado = df_mes_passado.iloc[0]['mix']
        else:
            faturamento_mes_passado = 0
            margem_mes_passado = 0
            mix_mes_passado = 0
        
        # Mês Atual
        if mes_atual_num and ano == ano_atual:
            df_mes_atual = df_vendedor[(df_vendedor['mes'] == mes_atual_num) & (df_vendedor['ano'] == ano_atual)]
            if not df_mes_atual.empty:
                faturamento_mes_atual = df_mes_atual.iloc[0]['faturamento']
                margem_mes_atual = df_mes_atual.iloc[0]['margem_percentual']
                mix_mes_atual = df_mes_atual.iloc[0]['mix']
            else:
                faturamento_mes_atual = 0
                margem_mes_atual = 0
                mix_mes_atual = 0
        else:
            faturamento_mes_atual = 0
            margem_mes_atual = 0
            mix_mes_atual = 0
        
        vendedores_resumo.append({
            'rca': rca,
            'nome_vendedor': nome_vendedor,
            'faturamento_anual': faturamento_anual,
            'margem_anual_pct': margem_anual_pct,
            'mix_anual': mix_anual,
            'faturamento_medio_ano': faturamento_medio_ano,
            'margem_media_ano': margem_media_ano,
            'mix_medio_ano': mix_medio_ano,
            'faturamento_medio_trimestre': faturamento_medio_trimestre,
            'margem_media_trimestre': margem_media_trimestre,
            'mix_medio_trimestre': mix_medio_trimestre,
            'faturamento_mes_passado': faturamento_mes_passado,
            'margem_mes_passado': margem_mes_passado,
            'mix_mes_passado': mix_mes_passado,
            'faturamento_mes_atual': faturamento_mes_atual,
            'margem_mes_atual': margem_mes_atual,
            'mix_mes_atual': mix_mes_atual
        })
    
    df_resumo = pd.DataFrame(vendedores_resumo)
    # Ordenar por faturamento anual (decrescente)
    df_resumo = df_resumo.sort_values('faturamento_anual', ascending=False).reset_index(drop=True)
    return df_resumo

def buscar_supervisores(_conn):
    """Busca lista de supervisores"""
    try:
        query = """
        SELECT 
            CODSUPERVISOR,
            NOME
        FROM PCSUPERV
        ORDER BY CODSUPERVISOR
        """
        df = pd.read_sql_query(query, _conn)
        df.columns = [x.lower() for x in df.columns]
        return df
    except:
        return pd.DataFrame()

def gerar_excel_simplificado(df_resumo, ano, caminho_arquivo):
    """Gera Excel simplificado com 3 abas: Faturamento, Margem e Mix"""
    wb = Workbook()
    
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])
    
    # Estilos
    header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=11)
    title_font = Font(bold=True, size=14, color="366092")
    border = Border(
        left=Side(style='thin'),
        right=Side(style='thin'),
        top=Side(style='thin'),
        bottom=Side(style='thin'),
    )
    center_align = Alignment(horizontal='center', vertical='center')
    right_align = Alignment(horizontal='right', vertical='center')
    
    def criar_aba(ws, titulo, colunas_dados, formato_numero):
        """Cria uma aba com os dados especificados"""
        ws.merge_cells('A1:G1')
        ws['A1'] = titulo
        ws['A1'].font = title_font
        ws['A1'].alignment = center_align
        
        headers = ['RCA', 'Nome Vendedor', 'Anual', 'Média Ano', 'Média Trimestre', 'Mês Passado', 'Mês Atual']
        for col, header in enumerate(headers, 1):
            cell = ws.cell(row=3, column=col)
            cell.value = header
            cell.fill = header_fill
            cell.font = header_font
            cell.alignment = center_align
            cell.border = border
        
        for idx, row in df_resumo.iterrows():
            r = idx + 4
            ws.cell(r, 1, row['rca']).border = border
            ws.cell(r, 2, row['nome_vendedor']).border = border
            
            # Dados das 5 colunas
            for col_idx, col_data in enumerate(colunas_dados, 3):
                valor = row[col_data]
                if formato_numero == 'percentual':
                    valor = valor / 100
                elif formato_numero == 'inteiro':
                    valor = int(valor)
                
                cell = ws.cell(r, col_idx)
                cell.value = valor
                if formato_numero == 'monetario':
                    cell.number_format = '#,##0.00'
                elif formato_numero == 'percentual':
                    cell.number_format = '0.00%'
                elif formato_numero == 'inteiro':
                    cell.number_format = '#,##0'
                cell.alignment = right_align
                cell.border = border
        
        ws.column_dimensions['A'].width = 10
        ws.column_dimensions['B'].width = 35
        for col in range(3, 8):
            ws.column_dimensions[get_column_letter(col)].width = 18
        ws.freeze_panes = 'A4'
    
    # ABA 1: FATURAMENTO
    ws_faturamento = wb.create_sheet("Faturamento", 0)
    criar_aba(ws_faturamento, f"FATURAMENTO - {ano}", 
              ['faturamento_anual', 'faturamento_medio_ano', 'faturamento_medio_trimestre', 
               'faturamento_mes_passado', 'faturamento_mes_atual'],
              'monetario')
    
    # ABA 2: MARGEM
    ws_margem = wb.create_sheet("Margem")
    criar_aba(ws_margem, f"MARGEM % - {ano}", 
              ['margem_anual_pct', 'margem_media_ano', 'margem_media_trimestre', 
               'margem_mes_passado', 'margem_mes_atual'],
              'percentual')
    
    # ABA 3: MIX
    ws_mix = wb.create_sheet("Mix")
    criar_aba(ws_mix, f"MIX - {ano}", 
              ['mix_anual', 'mix_medio_ano', 'mix_medio_trimestre', 
               'mix_mes_passado', 'mix_mes_atual'],
              'inteiro')
    
    # ABA 4: ANÁLISE MÉDIA ANUAL (ACIMA/ABAIXO)
    ws_analise = wb.create_sheet("Análise Média Anual")
    ws_analise.merge_cells('A1:E1')
    ws_analise['A1'] = f"ANÁLISE MÉDIA ANUAL - {ano}"
    ws_analise['A1'].font = title_font
    ws_analise['A1'].alignment = center_align
    
    # Legenda
    legenda_font = Font(size=10, italic=True)
    ws_analise.merge_cells('A2:E2')
    ws_analise['A2'] = "LEGENDA:"
    ws_analise['A2'].font = Font(size=10, bold=True)
    ws_analise['A2'].alignment = Alignment(horizontal='left', vertical='center')
    
    ws_analise.merge_cells('A3:E3')
    ws_analise['A3'] = "FATURAMENTO: ACIMA = Média anual ≥ R$ 200.000 | ABAIXO = Média anual < R$ 200.000"
    ws_analise['A3'].font = legenda_font
    ws_analise['A3'].alignment = Alignment(horizontal='left', vertical='center')
    
    ws_analise.merge_cells('A4:E4')
    ws_analise['A4'] = "MARGEM: ACIMA = Média anual ≥ 10% | ABAIXO = Média anual < 10%"
    ws_analise['A4'].font = legenda_font
    ws_analise['A4'].alignment = Alignment(horizontal='left', vertical='center')
    
    ws_analise.merge_cells('A5:E5')
    ws_analise['A5'] = "MIX: ACIMA = Média anual ≥ 20 produtos diferentes | ABAIXO = Média anual < 20 produtos diferentes"
    ws_analise['A5'].font = legenda_font
    ws_analise['A5'].alignment = Alignment(horizontal='left', vertical='center')
    
    headers_analise = ['RCA', 'NOME', 'FAT', 'MARGEM', 'MIX']
    for col, header in enumerate(headers_analise, 1):
        cell = ws_analise.cell(row=6, column=col)
        cell.value = header
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = border
    
    # Critérios
    limite_faturamento = 200000  # 200 mil
    limite_mix = 20
    limite_margem = 10  # 10%
    
    # Contadores para o resumo
    contadores = {
        'fat_acima': 0,
        'fat_abaixo': 0,
        'margem_acima': 0,
        'margem_abaixo': 0,
        'mix_acima': 0,
        'mix_abaixo': 0
    }
    
    for idx, row in df_resumo.iterrows():
        r = idx + 7
        ws_analise.cell(r, 1, row['rca']).border = border
        ws_analise.cell(r, 2, row['nome_vendedor']).border = border
        
        # FAT - baseado na média anual
        fat_medio = row['faturamento_medio_ano']
        fat_status = "ACIMA" if fat_medio >= limite_faturamento else "ABAIXO"
        if fat_status == "ACIMA":
            contadores['fat_acima'] += 1
        else:
            contadores['fat_abaixo'] += 1
        cell_fat = ws_analise.cell(r, 3, fat_status)
        cell_fat.border = border
        cell_fat.alignment = center_align
        
        # MARGEM - baseado na média anual
        margem_media = row['margem_media_ano']
        margem_status = "ACIMA" if margem_media >= limite_margem else "ABAIXO"
        if margem_status == "ACIMA":
            contadores['margem_acima'] += 1
        else:
            contadores['margem_abaixo'] += 1
        cell_margem = ws_analise.cell(r, 4, margem_status)
        cell_margem.border = border
        cell_margem.alignment = center_align
        
        # MIX - baseado na média anual
        mix_medio = row['mix_medio_ano']
        mix_status = "ACIMA" if mix_medio >= limite_mix else "ABAIXO"
        if mix_status == "ACIMA":
            contadores['mix_acima'] += 1
        else:
            contadores['mix_abaixo'] += 1
        cell_mix = ws_analise.cell(r, 5, mix_status)
        cell_mix.border = border
        cell_mix.alignment = center_align
    
    # Resumo no final
    linha_inicio_resumo = len(df_resumo) + 9
    linha_resumo = linha_inicio_resumo
    
    # Título do resumo
    ws_analise.merge_cells(f'A{linha_resumo}:D{linha_resumo}')
    ws_analise[f'A{linha_resumo}'] = "RESUMO ESTATÍSTICO"
    ws_analise[f'A{linha_resumo}'].font = Font(size=12, bold=True, color="366092")
    ws_analise[f'A{linha_resumo}'].alignment = center_align
    linha_resumo += 1
    
    # Cabeçalho da tabela de resumo
    ws_analise.merge_cells(f'A{linha_resumo}:B{linha_resumo}')
    cell_metrica = ws_analise[f'A{linha_resumo}']
    cell_metrica.value = "MÉTRICA"
    cell_metrica.fill = header_fill
    cell_metrica.font = header_font
    cell_metrica.alignment = center_align
    cell_metrica.border = border
    
    cell_acima = ws_analise[f'C{linha_resumo}']
    cell_acima.value = "ACIMA"
    cell_acima.fill = header_fill
    cell_acima.font = header_font
    cell_acima.alignment = center_align
    cell_acima.border = border
    
    cell_abaixo = ws_analise[f'D{linha_resumo}']
    cell_abaixo.value = "ABAIXO"
    cell_abaixo.fill = header_fill
    cell_abaixo.font = header_font
    cell_abaixo.alignment = center_align
    cell_abaixo.border = border
    linha_resumo += 1
    
    # Linha FATURAMENTO
    cell_fat_metrica = ws_analise[f'A{linha_resumo}']
    ws_analise.merge_cells(f'A{linha_resumo}:B{linha_resumo}')
    cell_fat_metrica.value = "FATURAMENTO"
    cell_fat_metrica.font = Font(bold=True)
    cell_fat_metrica.border = border
    
    cell_fat_acima = ws_analise[f'C{linha_resumo}']
    cell_fat_acima.value = contadores['fat_acima']
    cell_fat_acima.number_format = '#,##0'
    cell_fat_acima.alignment = center_align
    cell_fat_acima.border = border
    
    cell_fat_abaixo = ws_analise[f'D{linha_resumo}']
    cell_fat_abaixo.value = contadores['fat_abaixo']
    cell_fat_abaixo.number_format = '#,##0'
    cell_fat_abaixo.alignment = center_align
    cell_fat_abaixo.border = border
    linha_resumo += 1
    
    # Linha MARGEM
    cell_margem_metrica = ws_analise[f'A{linha_resumo}']
    ws_analise.merge_cells(f'A{linha_resumo}:B{linha_resumo}')
    cell_margem_metrica.value = "MARGEM"
    cell_margem_metrica.font = Font(bold=True)
    cell_margem_metrica.border = border
    
    cell_margem_acima = ws_analise[f'C{linha_resumo}']
    cell_margem_acima.value = contadores['margem_acima']
    cell_margem_acima.number_format = '#,##0'
    cell_margem_acima.alignment = center_align
    cell_margem_acima.border = border
    
    cell_margem_abaixo = ws_analise[f'D{linha_resumo}']
    cell_margem_abaixo.value = contadores['margem_abaixo']
    cell_margem_abaixo.number_format = '#,##0'
    cell_margem_abaixo.alignment = center_align
    cell_margem_abaixo.border = border
    linha_resumo += 1
    
    # Linha MIX
    cell_mix_metrica = ws_analise[f'A{linha_resumo}']
    ws_analise.merge_cells(f'A{linha_resumo}:B{linha_resumo}')
    cell_mix_metrica.value = "MIX"
    cell_mix_metrica.font = Font(bold=True)
    cell_mix_metrica.border = border
    
    cell_mix_acima = ws_analise[f'C{linha_resumo}']
    cell_mix_acima.value = contadores['mix_acima']
    cell_mix_acima.number_format = '#,##0'
    cell_mix_acima.alignment = center_align
    cell_mix_acima.border = border
    
    cell_mix_abaixo = ws_analise[f'D{linha_resumo}']
    cell_mix_abaixo.value = contadores['mix_abaixo']
    cell_mix_abaixo.number_format = '#,##0'
    cell_mix_abaixo.alignment = center_align
    cell_mix_abaixo.border = border
    
    ws_analise.column_dimensions['A'].width = 10
    ws_analise.column_dimensions['B'].width = 35
    ws_analise.column_dimensions['C'].width = 15
    ws_analise.column_dimensions['D'].width = 15
    ws_analise.column_dimensions['E'].width = 15
    ws_analise.freeze_panes = 'A7'
    
    # Salvar
    wb.save(caminho_arquivo)

# =========================================================================
# INTERFACE TKINTER
# =========================================================================
class AnaliseVendedoresApp:
    def __init__(self, root):
        self.root = root
        self.root.title("📊 Análise de Vendedores - Gerador de Excel")
        self.root.geometry("600x700")
        
        self.conn = None
        self.df_departamentos = None
        self.df_supervisores = None
        
        self.criar_interface()
        self.conectar_banco()
    
    def criar_interface(self):
        """Cria a interface gráfica"""
        main_frame = ttk.Frame(self.root, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # Título
        titulo = ttk.Label(main_frame, text="📊 Análise de Vendedores", 
                          font=("Arial", 16, "bold"))
        titulo.pack(pady=(0, 20))
        
        # Frame de filtros
        frame_filtros = ttk.LabelFrame(main_frame, text="Filtros", padding="15")
        frame_filtros.pack(fill=tk.BOTH, expand=True, pady=10)
        
        # Ano
        frame_ano = ttk.Frame(frame_filtros)
        frame_ano.pack(fill=tk.X, pady=5)
        ttk.Label(frame_ano, text="Ano:", width=15).pack(side=tk.LEFT)
        self.ano_var = tk.IntVar(value=datetime.now().year)
        self.combo_ano = ttk.Combobox(frame_ano, textvariable=self.ano_var, 
                                     values=list(range(datetime.now().year, datetime.now().year - 5, -1)),
                                     state="readonly", width=37)
        self.combo_ano.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Departamento
        frame_depto = ttk.Frame(frame_filtros)
        frame_depto.pack(fill=tk.X, pady=5)
        ttk.Label(frame_depto, text="Departamento:", width=15).pack(side=tk.LEFT)
        self.combo_depto = ttk.Combobox(frame_depto, width=40, state="readonly")
        self.combo_depto.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        
        # Supervisores
        frame_sup = ttk.Frame(frame_filtros)
        frame_sup.pack(fill=tk.BOTH, expand=True, pady=5)
        ttk.Label(frame_sup, text="Supervisores:", width=15).pack(anchor=tk.W)
        
        # Frame com scroll para supervisores
        canvas_sup = tk.Canvas(frame_sup, height=150)
        scrollbar_sup = ttk.Scrollbar(frame_sup, orient="vertical", command=canvas_sup.yview)
        scrollable_sup = ttk.Frame(canvas_sup)
        
        scrollable_sup.bind(
            "<Configure>",
            lambda e: canvas_sup.configure(scrollregion=canvas_sup.bbox("all"))
        )
        
        canvas_sup.create_window((0, 0), window=scrollable_sup, anchor="nw")
        canvas_sup.configure(yscrollcommand=scrollbar_sup.set)
        
        self.supervisores_vars = {}
        self.frame_supervisores = scrollable_sup
        
        canvas_sup.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar_sup.pack(side=tk.RIGHT, fill=tk.Y)
        
        # Botões de seleção rápida supervisores
        frame_botoes_sup = ttk.Frame(frame_filtros)
        frame_botoes_sup.pack(fill=tk.X, pady=5)
        ttk.Button(frame_botoes_sup, text="Selecionar Todos", 
                  command=self.selecionar_todos_supervisores).pack(side=tk.LEFT, padx=5)
        ttk.Button(frame_botoes_sup, text="Apenas 1 e 2", 
                  command=self.selecionar_supervisores_1_2).pack(side=tk.LEFT, padx=5)
        ttk.Button(frame_botoes_sup, text="Limpar", 
                  command=self.limpar_supervisores).pack(side=tk.LEFT, padx=5)
        
        # RCAs
        frame_rca = ttk.Frame(frame_filtros)
        frame_rca.pack(fill=tk.X, pady=5)
        ttk.Label(frame_rca, text="RCAs (vírgula):", width=15).pack(side=tk.LEFT)
        self.entry_rca = ttk.Entry(frame_rca, width=50)
        self.entry_rca.pack(side=tk.LEFT, padx=5, fill=tk.X, expand=True)
        ttk.Label(frame_rca, text="(deixe vazio para todos)").pack(side=tk.LEFT, padx=5)
        
        # Botão de gerar
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(fill=tk.X, pady=(20, 0))
        
        self.btn_gerar = ttk.Button(btn_frame, text="📊 Gerar Excel", 
                                   command=self.gerar_excel, state=tk.DISABLED)
        self.btn_gerar.pack(side=tk.LEFT, padx=(0, 10))
        
        self.status_label = ttk.Label(btn_frame, text="Conectando ao banco...")
        self.status_label.pack(side=tk.LEFT)
    
    def conectar_banco(self):
        """Conecta ao banco e carrega dados"""
        def carregar():
            try:
                self.status_label.config(text="Conectando ao banco...")
                self.conn = get_db_connection()
                
                if not self.conn:
                    self.status_label.config(text="❌ Erro ao conectar")
                    messagebox.showerror("Erro", "Não foi possível conectar ao banco de dados")
                    return
                
                # Carregar departamentos
                self.status_label.config(text="Carregando departamentos...")
                query_deptos = "SELECT CODEPTO, DESCRICAO FROM PCDEPTO ORDER BY DESCRICAO"
                self.df_departamentos = pd.read_sql_query(query_deptos, self.conn)
                self.df_departamentos.columns = [x.lower() for x in self.df_departamentos.columns]
                
                if not self.df_departamentos.empty:
                    deptos = ["Todos"] + [f"{int(row['codepto'])} - {row['descricao']}" 
                                         for _, row in self.df_departamentos.iterrows()]
                    self.combo_depto['values'] = deptos
                    self.combo_depto.current(0)
                
                # Carregar supervisores
                self.status_label.config(text="Carregando supervisores...")
                self.df_supervisores = buscar_supervisores(self.conn)
                
                if not self.df_supervisores.empty:
                    for _, row in self.df_supervisores.iterrows():
                        var = tk.BooleanVar()
                        cod = int(row['codsupervisor'])
                        # Marcar 1 e 2 por padrão
                        if cod in [1, 2]:
                            var.set(True)
                        self.supervisores_vars[cod] = var
                        cb = ttk.Checkbutton(
                            self.frame_supervisores,
                            text=f"{cod} - {row['nome']}",
                            variable=var
                        )
                        cb.pack(anchor=tk.W, pady=2)
                
                self.status_label.config(text="✅ Pronto para gerar relatório")
                self.btn_gerar.config(state=tk.NORMAL)
                
            except Exception as e:
                self.status_label.config(text=f"❌ Erro: {str(e)}")
                messagebox.showerror("Erro", f"Erro ao conectar: {e}")
        
        threading.Thread(target=carregar, daemon=True).start()
    
    def selecionar_todos_supervisores(self):
        """Seleciona todos os supervisores"""
        for var in self.supervisores_vars.values():
            var.set(True)
    
    def selecionar_supervisores_1_2(self):
        """Seleciona apenas supervisores 1 e 2"""
        for cod, var in self.supervisores_vars.items():
            var.set(cod in [1, 2])
    
    def limpar_supervisores(self):
        """Limpa seleção de supervisores"""
        for var in self.supervisores_vars.values():
            var.set(False)
    
    def gerar_excel(self):
        """Gera o arquivo Excel"""
        if not self.conn:
            messagebox.showerror("Erro", "Não há conexão com o banco de dados")
            return
        
        # Obter filtros
        ano = self.ano_var.get()
        
        depto_sel = self.combo_depto.get()
        coddepto = None
        if depto_sel and depto_sel != "Todos":
            coddepto = int(depto_sel.split(" - ")[0])
        
        supervisores_list = [cod for cod, var in self.supervisores_vars.items() if var.get()]
        if not supervisores_list:
            messagebox.showwarning("Aviso", "Selecione pelo menos um supervisor!")
            return
        
        rcas_str = self.entry_rca.get().strip()
        rcas_list = None
        if rcas_str:
            try:
                rcas_list = [int(r.strip()) for r in rcas_str.split(',') if r.strip().isdigit()]
                if not rcas_list:
                    rcas_list = None
            except:
                rcas_list = None
        
        def executar():
            try:
                self.status_label.config(text="Buscando dados do ano...")
                self.root.update()
                
                # Buscar dados
                df_ano = buscar_dados_ano_completo(self.conn, ano, coddepto, rcas_list, supervisores_list)
                
                # Buscar mês atual se for o ano atual
                hoje = datetime.now()
                df_mes_atual = pd.DataFrame()
                if ano == hoje.year:
                    self.status_label.config(text="Buscando mês atual...")
                    self.root.update()
                    df_mes_atual = buscar_vendedores_mes(self.conn, hoje.month, hoje.year, coddepto, rcas_list, supervisores_list)
                
                if df_ano.empty and df_mes_atual.empty:
                    messagebox.showwarning("Aviso", "Nenhum dado encontrado para os filtros selecionados")
                    self.status_label.config(text="❌ Nenhum dado encontrado")
                    return
                
                self.status_label.config(text="Calculando resumo...")
                self.root.update()
                
                df_resumo = calcular_resumo_vendedores(df_ano, df_mes_atual, ano)
                
                if df_resumo.empty:
                    messagebox.showwarning("Aviso", "Não foi possível calcular resumo")
                    self.status_label.config(text="❌ Erro ao calcular resumo")
                    return
                
                # Salvar arquivo
                self.status_label.config(text="Gerando Excel...")
                self.root.update()
                
                nome_arquivo = f"analise_completa_vendedores_{ano}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
                caminho_arquivo = Path(nome_arquivo).resolve()
                
                gerar_excel_simplificado(df_resumo, ano, caminho_arquivo)
                
                self.status_label.config(text="✅ Excel gerado com sucesso!")
                messagebox.showinfo("Sucesso", f"Excel gerado com sucesso!\n\nArquivo: {caminho_arquivo}")
                
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
    app = AnaliseVendedoresApp(root)
    root.mainloop()
# portfolio-commit-ready: analise_vendedores_excel
