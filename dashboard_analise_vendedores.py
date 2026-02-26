"""
Dashboard Streamlit - Análise Anual de Vendedores
Análise completa de desempenho dos vendedores ao longo do ano
Mostra tendências, comparações e métricas detalhadas
"""

import streamlit as st
import pandas as pd
import numpy as np
import oracledb
from datetime import datetime
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from io import BytesIO

# Configuração do Oracle Instant Client
try:
    oracledb.init_oracle_client(lib_dir=r"C:\oracle\instantclient_23_9")
except:
    pass

# =========================================================================
# CONFIGURAÇÕES
# =========================================================================
DB_HOST = '10.0.0.10'
DB_PORT = 1521
DB_SERVICE = 'PROD'
DB_USER = 'powerbi'
DB_PASSWORD = 'cbjc4xp3nlq6'

# =========================================================================
# CONFIGURAÇÃO STREAMLIT
# =========================================================================
st.set_page_config(
    page_title="Análise Anual de Vendedores",
    page_icon="📊",
    layout="wide",
    initial_sidebar_state="expanded"
)

# =========================================================================
# FUNÇÕES DE BANCO
# =========================================================================
@st.cache_resource
def get_db_connection():
    """Conecta ao banco Oracle"""
    try:
        return oracledb.connect(
            user=DB_USER,
            password=DB_PASSWORD,
            dsn=f"{DB_HOST}:{DB_PORT}/{DB_SERVICE}",
        )
    except Exception as e:
        st.error(f"Erro ao conectar ao banco: {e}")
        return None

@st.cache_data(ttl=3600)
def buscar_supervisores(_conn):
    """Busca lista de supervisores do banco"""
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
    except Exception as e:
        st.error(f"Erro ao buscar supervisores: {e}")
        return pd.DataFrame()

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
    WITH CLIENTES_POSITIVADOS AS (
        SELECT DISTINCT
            C.CODUSUR,
            C.CODCLI
        FROM PCPEDC C 
        JOIN PCPEDI I ON C.NUMPED = I.NUMPED
        JOIN PCPRODUT P ON I.CODPROD = P.CODPROD
        JOIN PCUSUARI U ON C.CODUSUR = U.CODUSUR
        WHERE TRUNC(C.DATA, 'MM') = TO_DATE('{data_inicio_str}', 'DD/MM/YYYY')
          AND C.CODFILIAL IN ('1', '98')
          AND C.POSICAO = 'F'
          AND C.DTCANCEL IS NULL
          AND C.CONDVENDA NOT IN (4, 8, 10, 13, 20, 98, 99)
          AND NVL(I.BONIFIC, 'N') = 'N'
          {filtro_depto}
          {filtro_rca}
          {filtro_supervisor}
    ),
    VENDAS_VALIDAS AS (
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
            SUM(
                CASE 
                    WHEN NVL(I.BONIFIC, 'N') = 'N' THEN 
                        DECODE(C.CONDVENDA, 
                            5, 0, 6, 0, 11, 0, 12, 0, 
                            CASE 
                                WHEN NVL(P.QTUNITCX, 1) > 0 THEN 
                                    ROUND(NVL(I.QT, 0) / NVL(P.QTUNITCX, 1), 2)
                                ELSE NVL(I.QT, 0)
                            END
                        )
                    ELSE 0 
                END
            ) AS QTD_CAIXAS,
            SUM(
                CASE 
                    WHEN NVL(I.BONIFIC, 'N') = 'N' THEN 
                        DECODE(C.CONDVENDA, 
                            5, 0, 6, 0, 11, 0, 12, 0, 
                            NVL(I.QT, 0) * NVL(P.PESOBRUTO, 0)
                        )
                    ELSE 0 
                END
            ) AS PESO_TOTAL,
            COUNT(DISTINCT I.CODPROD) AS QTD_PRODUTOS_DIFERENTES,
            SUM(
                CASE 
                    WHEN NVL(I.BONIFIC, 'N') = 'N' THEN 
                        DECODE(C.CONDVENDA, 
                            5, 0, 6, 0, 11, 0, 12, 0, 
                            NVL(I.QT, 0)
                        )
                    ELSE 0 
                END
            ) AS QTD_ITENS_TOTAL
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
    POSITIVACAO_POR_VENDEDOR AS (
        SELECT 
            CODUSUR,
            COUNT(*) AS POSITIVACAO
        FROM CLIENTES_POSITIVADOS
        GROUP BY CODUSUR
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
        NVL(P.POSITIVACAO, 0) AS POSITIVACAO,
        (NVL(V.VALOR_BRUTO, 0) - NVL(D.VALOR_DEVOLVIDO, 0)) - (NVL(V.CUSTO_BRUTO, 0) - NVL(D.CUSTO_DEVOLVIDO, 0)) AS MARGEM_VALOR,
        NVL(V.QTD_CAIXAS, 0) AS QTD_CAIXAS,
        NVL(V.PESO_TOTAL, 0) AS PESO_VENDIDO,
        NVL(V.QTD_PRODUTOS_DIFERENTES, 0) AS QTD_PRODUTOS_DIFERENTES,
        NVL(V.QTD_ITENS_TOTAL, 0) AS QTD_ITENS_TOTAL
    FROM VENDAS_VALIDAS V
    LEFT JOIN DEVOLUCOES D ON V.CODUSUR = D.CODUSUR
    LEFT JOIN POSITIVACAO_POR_VENDEDOR P ON V.CODUSUR = P.CODUSUR
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
        st.error(f"Erro ao buscar dados: {e}")
        return pd.DataFrame()

@st.cache_data(ttl=3600)
def buscar_dados_ano_completo(_conn, ano, coddepto=None, rcas_list=None, supervisores_list=None):
    """Busca dados de todos os meses do ano até o mês anterior ao atual"""
    dados_mensais = []
    
    # Obter mês atual
    hoje = datetime.now()
    mes_atual = hoje.month
    ano_atual = hoje.year
    
    # Determinar até qual mês buscar
    # Se o ano selecionado for o ano atual, buscar apenas até o mês anterior
    # Se for um ano anterior, buscar todos os 12 meses
    if ano == ano_atual:
        if mes_atual == 1:
            # Se estamos em janeiro, não buscar nada do ano atual (ainda não fechou)
            mes_limite = 0
        else:
            mes_limite = mes_atual - 1  # Até o mês anterior
    else:
        mes_limite = 12  # Ano completo se for ano anterior
    
    # Se mes_limite for 0, não há dados para buscar
    if mes_limite == 0:
        return pd.DataFrame()
    
    # Buscar dados de cada mês
    for mes in range(1, mes_limite + 1):
        df_mes = buscar_vendedores_mes(_conn, mes, ano, coddepto, rcas_list, supervisores_list)
        if not df_mes.empty:
            dados_mensais.append(df_mes)
    
    if not dados_mensais:
        return pd.DataFrame()
    
    df_completo = pd.concat(dados_mensais, ignore_index=True)
    return df_completo

def calcular_tendencias(df):
    """Calcula tendências de crescimento/diminuição por vendedor"""
    if df.empty:
        return pd.DataFrame()
    
    # Agrupar por vendedor e calcular métricas
    vendedores_agg = []
    
    for rca in df['rca'].unique():
        df_vendedor = df[df['rca'] == rca].sort_values(['ano', 'mes'])
        
        if len(df_vendedor) < 2:
            continue
        
        # Primeiro e último mês
        primeiro_mes = df_vendedor.iloc[0]
        ultimo_mes = df_vendedor.iloc[-1]
        
        # Identificar melhor e pior mês (por faturamento)
        melhor_mes_idx = df_vendedor['faturamento'].idxmax()
        pior_mes_idx = df_vendedor['faturamento'].idxmin()
        melhor_mes = df_vendedor.loc[melhor_mes_idx]
        pior_mes = df_vendedor.loc[pior_mes_idx]
        
        # Calcular variações (primeiro vs último)
        var_faturamento = ultimo_mes['faturamento'] - primeiro_mes['faturamento']
        var_faturamento_pct = ((ultimo_mes['faturamento'] - primeiro_mes['faturamento']) / primeiro_mes['faturamento'] * 100) if primeiro_mes['faturamento'] > 0 else 0
        
        var_margem_pct = ultimo_mes['margem_percentual'] - primeiro_mes['margem_percentual']
        var_positivacao = ultimo_mes['positivacao'] - primeiro_mes['positivacao']
        var_qtd_caixas = ultimo_mes['qtd_caixas'] - primeiro_mes['qtd_caixas']
        var_qtd_itens = ultimo_mes['qtd_itens_total'] - primeiro_mes['qtd_itens_total']
        var_qtd_produtos = ultimo_mes['qtd_produtos_diferentes'] - primeiro_mes['qtd_produtos_diferentes']
        
        # Comparações: Último mês vs Melhor mês
        var_vs_melhor_fat = ultimo_mes['faturamento'] - melhor_mes['faturamento']
        var_vs_melhor_fat_pct = ((ultimo_mes['faturamento'] - melhor_mes['faturamento']) / melhor_mes['faturamento'] * 100) if melhor_mes['faturamento'] > 0 else 0
        var_vs_melhor_margem = ultimo_mes['margem_percentual'] - melhor_mes['margem_percentual']
        var_vs_melhor_posit = ultimo_mes['positivacao'] - melhor_mes['positivacao']
        var_vs_melhor_itens = ultimo_mes['qtd_itens_total'] - melhor_mes['qtd_itens_total']
        var_vs_melhor_mix = ultimo_mes['qtd_produtos_diferentes'] - melhor_mes['qtd_produtos_diferentes']
        
        # Comparações: Último mês vs Pior mês
        var_vs_pior_fat = ultimo_mes['faturamento'] - pior_mes['faturamento']
        var_vs_pior_fat_pct = ((ultimo_mes['faturamento'] - pior_mes['faturamento']) / pior_mes['faturamento'] * 100) if pior_mes['faturamento'] > 0 else 0
        var_vs_pior_margem = ultimo_mes['margem_percentual'] - pior_mes['margem_percentual']
        var_vs_pior_posit = ultimo_mes['positivacao'] - pior_mes['positivacao']
        var_vs_pior_itens = ultimo_mes['qtd_itens_total'] - pior_mes['qtd_itens_total']
        var_vs_pior_mix = ultimo_mes['qtd_produtos_diferentes'] - pior_mes['qtd_produtos_diferentes']
        
        # Médias do ano
        faturamento_medio = df_vendedor['faturamento'].mean()
        margem_media = df_vendedor['margem_percentual'].mean()
        positivacao_media = df_vendedor['positivacao'].mean()
        qtd_caixas_media = df_vendedor['qtd_caixas'].mean()
        qtd_itens_media = df_vendedor['qtd_itens_total'].mean()
        qtd_produtos_media = df_vendedor['qtd_produtos_diferentes'].mean()
        
        # Total do ano
        faturamento_total = df_vendedor['faturamento'].sum()
        margem_total = df_vendedor['margem_valor'].sum()
        margem_total_pct = (margem_total / faturamento_total * 100) if faturamento_total > 0 else 0
        
        # Tendência (crescente, decrescente, estável)
        meses_com_venda = len(df_vendedor[df_vendedor['faturamento'] > 0])
        
        # Calcular tendência baseada na regressão linear simples
        if len(df_vendedor) >= 3:
            x = np.arange(len(df_vendedor))
            y = df_vendedor['faturamento'].values
            coef = np.polyfit(x, y, 1)[0]
            if coef > 0:
                tendencia = "Crescente"
            elif coef < 0:
                tendencia = "Decrescente"
            else:
                tendencia = "Estável"
        else:
            if var_faturamento_pct > 5:
                tendencia = "Crescente"
            elif var_faturamento_pct < -5:
                tendencia = "Decrescente"
            else:
                tendencia = "Estável"
        
        vendedores_agg.append({
            'rca': rca,
            'nome_vendedor': primeiro_mes['nome_vendedor'],
            'faturamento_total_ano': faturamento_total,
            'faturamento_medio_mes': faturamento_medio,
            'faturamento_primeiro_mes': primeiro_mes['faturamento'],
            'faturamento_ultimo_mes': ultimo_mes['faturamento'],
            'faturamento_melhor_mes': melhor_mes['faturamento'],
            'faturamento_pior_mes': pior_mes['faturamento'],
            'mes_melhor': melhor_mes['mes'],
            'mes_pior': pior_mes['mes'],
            'var_faturamento': var_faturamento,
            'var_faturamento_pct': var_faturamento_pct,
            'var_vs_melhor_fat': var_vs_melhor_fat,
            'var_vs_melhor_fat_pct': var_vs_melhor_fat_pct,
            'var_vs_pior_fat': var_vs_pior_fat,
            'var_vs_pior_fat_pct': var_vs_pior_fat_pct,
            'margem_media_ano': margem_media,
            'margem_total_pct': margem_total_pct,
            'margem_melhor_mes': melhor_mes['margem_percentual'],
            'margem_pior_mes': pior_mes['margem_percentual'],
            'margem_ultimo_mes': ultimo_mes['margem_percentual'],
            'var_margem_pct': var_margem_pct,
            'var_vs_melhor_margem': var_vs_melhor_margem,
            'var_vs_pior_margem': var_vs_pior_margem,
            'positivacao_media': positivacao_media,
            'positivacao_melhor_mes': melhor_mes['positivacao'],
            'positivacao_pior_mes': pior_mes['positivacao'],
            'positivacao_ultimo_mes': ultimo_mes['positivacao'],
            'var_positivacao': var_positivacao,
            'var_vs_melhor_posit': var_vs_melhor_posit,
            'var_vs_pior_posit': var_vs_pior_posit,
            'qtd_itens_melhor_mes': melhor_mes['qtd_itens_total'],
            'qtd_itens_pior_mes': pior_mes['qtd_itens_total'],
            'qtd_itens_ultimo_mes': ultimo_mes['qtd_itens_total'],
            'mix_melhor_mes': melhor_mes['qtd_produtos_diferentes'],
            'mix_pior_mes': pior_mes['qtd_produtos_diferentes'],
            'mix_ultimo_mes': ultimo_mes['qtd_produtos_diferentes'],
            'qtd_caixas_media': qtd_caixas_media,
            'var_qtd_caixas': var_qtd_caixas,
            'qtd_itens_media': qtd_itens_media,
            'var_qtd_itens': var_qtd_itens,
            'var_vs_melhor_itens': var_vs_melhor_itens,
            'var_vs_pior_itens': var_vs_pior_itens,
            'var_vs_melhor_mix': var_vs_melhor_mix,
            'var_vs_pior_mix': var_vs_pior_mix,
            'qtd_produtos_media': qtd_produtos_media,
            'var_qtd_produtos': var_qtd_produtos,
            'meses_com_venda': meses_com_venda,
            'tendencia': tendencia
        })
    
    df_tendencias = pd.DataFrame(vendedores_agg)
    return df_tendencias

# =========================================================================
# FUNÇÃO DE EXPORTAÇÃO EXCEL
# =========================================================================
def gerar_excel_analise_completa(df_tendencias, df_ano, ano):
    """Gera Excel completo com múltiplas abas de análise"""
    wb = Workbook()
    
    # Remover aba padrão
    if 'Sheet' in wb.sheetnames:
        wb.remove(wb['Sheet'])
    
    meses_nomes = {
        1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril',
        5: 'Maio', 6: 'Junho', 7: 'Julho', 8: 'Agosto',
        9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'
    }
    
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
    
    # =====================================================================
    # ABA 1: RESUMO EXECUTIVO
    # =====================================================================
    ws_resumo = wb.create_sheet("Resumo Executivo", 0)
    
    # Título
    ws_resumo.merge_cells('A1:J1')
    ws_resumo['A1'] = f"RESUMO EXECUTIVO - ANÁLISE DE VENDEDORES {ano}"
    ws_resumo['A1'].font = title_font
    ws_resumo['A1'].alignment = center_align
    
    # Cabeçalho
    headers_resumo = [
        'RCA', 'Nome Vendedor', 'Faturamento Total', 'Faturamento Médio/Mês',
        'Melhor Mês', 'Pior Mês', 'Último Mês', 'Variação % (1º vs Último)',
        'Tendência', 'Meses com Venda'
    ]
    for col, header in enumerate(headers_resumo, 1):
        cell = ws_resumo.cell(row=3, column=col)
        cell.value = header
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = border
    
    # Dados
    for idx, row in df_tendencias.iterrows():
        r = idx + 4
        ws_resumo.cell(r, 1, row['rca']).border = border
        ws_resumo.cell(r, 2, row['nome_vendedor']).border = border
        ws_resumo.cell(r, 3, row['faturamento_total_ano']).number_format = '#,##0.00'
        ws_resumo.cell(r, 3).alignment = right_align
        ws_resumo.cell(r, 3).border = border
        ws_resumo.cell(r, 4, row['faturamento_medio_mes']).number_format = '#,##0.00'
        ws_resumo.cell(r, 4).alignment = right_align
        ws_resumo.cell(r, 4).border = border
        ws_resumo.cell(r, 5, f"{meses_nomes.get(int(row['mes_melhor']), row['mes_melhor'])} - R$ {row['faturamento_melhor_mes']:,.2f}")
        ws_resumo.cell(r, 5).border = border
        ws_resumo.cell(r, 6, f"{meses_nomes.get(int(row['mes_pior']), row['mes_pior'])} - R$ {row['faturamento_pior_mes']:,.2f}")
        ws_resumo.cell(r, 6).border = border
        ws_resumo.cell(r, 7, row['faturamento_ultimo_mes']).number_format = '#,##0.00'
        ws_resumo.cell(r, 7).alignment = right_align
        ws_resumo.cell(r, 7).border = border
        ws_resumo.cell(r, 8, row['var_faturamento_pct'] / 100).number_format = '0.00%'
        ws_resumo.cell(r, 8).alignment = right_align
        ws_resumo.cell(r, 8).border = border
        ws_resumo.cell(r, 9, row['tendencia']).border = border
        ws_resumo.cell(r, 10, row['meses_com_venda']).border = border
        ws_resumo.cell(r, 10).alignment = center_align
    
    # Ajustar larguras
    ws_resumo.column_dimensions['A'].width = 10
    ws_resumo.column_dimensions['B'].width = 35
    for col in range(3, 11):
        ws_resumo.column_dimensions[get_column_letter(col)].width = 18
    
    # =====================================================================
    # ABA 2: ANÁLISE DE FATURAMENTO
    # =====================================================================
    ws_fat = wb.create_sheet("Análise Faturamento")
    
    # Título
    ws_fat.merge_cells('A1:M1')
    ws_fat['A1'] = f"ANÁLISE DE FATURAMENTO - {ano}"
    ws_fat['A1'].font = title_font
    ws_fat['A1'].alignment = center_align
    
    # Cabeçalho
    headers_fat = [
        'RCA', 'Nome', 'Total Ano', 'Média/Mês', 'Melhor Mês', 'Faturamento Melhor',
        'Pior Mês', 'Faturamento Pior', 'Último Mês', 'Faturamento Último',
        'Var vs Melhor', 'Var vs Melhor %', 'Var vs Pior', 'Var vs Pior %'
    ]
    for col, header in enumerate(headers_fat, 1):
        cell = ws_fat.cell(row=3, column=col)
        cell.value = header
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = border
    
    # Dados
    for idx, row in df_tendencias.iterrows():
        r = idx + 4
        ws_fat.cell(r, 1, row['rca']).border = border
        ws_fat.cell(r, 2, row['nome_vendedor']).border = border
        ws_fat.cell(r, 3, row['faturamento_total_ano']).number_format = '#,##0.00'
        ws_fat.cell(r, 3).alignment = right_align
        ws_fat.cell(r, 3).border = border
        ws_fat.cell(r, 4, row['faturamento_medio_mes']).number_format = '#,##0.00'
        ws_fat.cell(r, 4).alignment = right_align
        ws_fat.cell(r, 4).border = border
        ws_fat.cell(r, 5, meses_nomes.get(int(row['mes_melhor']), row['mes_melhor'])).border = border
        ws_fat.cell(r, 6, row['faturamento_melhor_mes']).number_format = '#,##0.00'
        ws_fat.cell(r, 6).alignment = right_align
        ws_fat.cell(r, 6).border = border
        ws_fat.cell(r, 7, meses_nomes.get(int(row['mes_pior']), row['mes_pior'])).border = border
        ws_fat.cell(r, 8, row['faturamento_pior_mes']).number_format = '#,##0.00'
        ws_fat.cell(r, 8).alignment = right_align
        ws_fat.cell(r, 8).border = border
        ws_fat.cell(r, 9, "Último").border = border
        ws_fat.cell(r, 10, row['faturamento_ultimo_mes']).number_format = '#,##0.00'
        ws_fat.cell(r, 10).alignment = right_align
        ws_fat.cell(r, 10).border = border
        ws_fat.cell(r, 11, row['var_vs_melhor_fat']).number_format = '#,##0.00'
        ws_fat.cell(r, 11).alignment = right_align
        ws_fat.cell(r, 11).border = border
        ws_fat.cell(r, 12, row['var_vs_melhor_fat_pct'] / 100).number_format = '0.00%'
        ws_fat.cell(r, 12).alignment = right_align
        ws_fat.cell(r, 12).border = border
        ws_fat.cell(r, 13, row['var_vs_pior_fat']).number_format = '#,##0.00'
        ws_fat.cell(r, 13).alignment = right_align
        ws_fat.cell(r, 13).border = border
        ws_fat.cell(r, 14, row['var_vs_pior_fat_pct'] / 100).number_format = '0.00%'
        ws_fat.cell(r, 14).alignment = right_align
        ws_fat.cell(r, 14).border = border
    
    # Ajustar larguras
    ws_fat.column_dimensions['A'].width = 10
    ws_fat.column_dimensions['B'].width = 30
    for col in range(3, 15):
        ws_fat.column_dimensions[get_column_letter(col)].width = 16
    
    # =====================================================================
    # ABA 3: ANÁLISE DE MARGEM
    # =====================================================================
    ws_margem = wb.create_sheet("Análise Margem")
    
    # Título
    ws_margem.merge_cells('A1:J1')
    ws_margem['A1'] = f"ANÁLISE DE MARGEM - {ano}"
    ws_margem['A1'].font = title_font
    ws_margem['A1'].alignment = center_align
    
    # Cabeçalho
    headers_margem = [
        'RCA', 'Nome', 'Margem Média Ano', 'Margem Total %', 'Melhor Mês', 'Margem Melhor',
        'Pior Mês', 'Margem Pior', 'Último Mês', 'Margem Último', 'Var vs Melhor', 'Var vs Pior'
    ]
    for col, header in enumerate(headers_margem, 1):
        cell = ws_margem.cell(row=3, column=col)
        cell.value = header
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = border
    
    # Dados
    for idx, row in df_tendencias.iterrows():
        r = idx + 4
        ws_margem.cell(r, 1, row['rca']).border = border
        ws_margem.cell(r, 2, row['nome_vendedor']).border = border
        ws_margem.cell(r, 3, row['margem_media_ano'] / 100).number_format = '0.00%'
        ws_margem.cell(r, 3).alignment = right_align
        ws_margem.cell(r, 3).border = border
        ws_margem.cell(r, 4, row['margem_total_pct'] / 100).number_format = '0.00%'
        ws_margem.cell(r, 4).alignment = right_align
        ws_margem.cell(r, 4).border = border
        ws_margem.cell(r, 5, meses_nomes.get(int(row['mes_melhor']), row['mes_melhor'])).border = border
        ws_margem.cell(r, 6, row['margem_melhor_mes'] / 100).number_format = '0.00%'
        ws_margem.cell(r, 6).alignment = right_align
        ws_margem.cell(r, 6).border = border
        ws_margem.cell(r, 7, meses_nomes.get(int(row['mes_pior']), row['mes_pior'])).border = border
        ws_margem.cell(r, 8, row['margem_pior_mes'] / 100).number_format = '0.00%'
        ws_margem.cell(r, 8).alignment = right_align
        ws_margem.cell(r, 8).border = border
        ws_margem.cell(r, 9, "Último").border = border
        ws_margem.cell(r, 10, row['margem_ultimo_mes'] / 100).number_format = '0.00%'
        ws_margem.cell(r, 10).alignment = right_align
        ws_margem.cell(r, 10).border = border
        ws_margem.cell(r, 11, row['var_vs_melhor_margem'] / 100).number_format = '0.00%'
        ws_margem.cell(r, 11).alignment = right_align
        ws_margem.cell(r, 11).border = border
        ws_margem.cell(r, 12, row['var_vs_pior_margem'] / 100).number_format = '0.00%'
        ws_margem.cell(r, 12).alignment = right_align
        ws_margem.cell(r, 12).border = border
    
    # Ajustar larguras
    ws_margem.column_dimensions['A'].width = 10
    ws_margem.column_dimensions['B'].width = 30
    for col in range(3, 13):
        ws_margem.column_dimensions[get_column_letter(col)].width = 16
    
    # =====================================================================
    # ABA 4: ANÁLISE DE MIX
    # =====================================================================
    ws_mix = wb.create_sheet("Análise MIX")
    
    # Título
    ws_mix.merge_cells('A1:I1')
    ws_mix['A1'] = f"ANÁLISE DE MIX (PRODUTOS DIFERENTES) - {ano}"
    ws_mix['A1'].font = title_font
    ws_mix['A1'].alignment = center_align
    
    # Cabeçalho
    headers_mix = [
        'RCA', 'Nome', 'MIX Médio Ano', 'Melhor Mês', 'MIX Melhor',
        'Pior Mês', 'MIX Pior', 'Último Mês', 'MIX Último', 'Var vs Melhor', 'Var vs Pior'
    ]
    for col, header in enumerate(headers_mix, 1):
        cell = ws_mix.cell(row=3, column=col)
        cell.value = header
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = border
    
    # Dados
    for idx, row in df_tendencias.iterrows():
        r = idx + 4
        ws_mix.cell(r, 1, row['rca']).border = border
        ws_mix.cell(r, 2, row['nome_vendedor']).border = border
        ws_mix.cell(r, 3, row['qtd_produtos_media']).number_format = '#,##0'
        ws_mix.cell(r, 3).alignment = right_align
        ws_mix.cell(r, 3).border = border
        ws_mix.cell(r, 4, meses_nomes.get(int(row['mes_melhor']), row['mes_melhor'])).border = border
        ws_mix.cell(r, 5, int(row['mix_melhor_mes'])).number_format = '#,##0'
        ws_mix.cell(r, 5).alignment = right_align
        ws_mix.cell(r, 5).border = border
        ws_mix.cell(r, 6, meses_nomes.get(int(row['mes_pior']), row['mes_pior'])).border = border
        ws_mix.cell(r, 7, int(row['mix_pior_mes'])).number_format = '#,##0'
        ws_mix.cell(r, 7).alignment = right_align
        ws_mix.cell(r, 7).border = border
        ws_mix.cell(r, 8, "Último").border = border
        ws_mix.cell(r, 9, int(row['mix_ultimo_mes'])).number_format = '#,##0'
        ws_mix.cell(r, 9).alignment = right_align
        ws_mix.cell(r, 9).border = border
        ws_mix.cell(r, 10, int(row['var_vs_melhor_mix'])).number_format = '#,##0'
        ws_mix.cell(r, 10).alignment = right_align
        ws_mix.cell(r, 10).border = border
        ws_mix.cell(r, 11, int(row['var_vs_pior_mix'])).number_format = '#,##0'
        ws_mix.cell(r, 11).alignment = right_align
        ws_mix.cell(r, 11).border = border
    
    # Ajustar larguras
    ws_mix.column_dimensions['A'].width = 10
    ws_mix.column_dimensions['B'].width = 30
    for col in range(3, 12):
        ws_mix.column_dimensions[get_column_letter(col)].width = 16
    
    # =====================================================================
    # ABA 5: DADOS MENSAIS COMPLETOS
    # =====================================================================
    ws_mensal = wb.create_sheet("Dados Mensais")
    
    # Título
    ws_mensal.merge_cells('A1:J1')
    ws_mensal['A1'] = f"DADOS MENSAIS COMPLETOS - {ano}"
    ws_mensal['A1'].font = title_font
    ws_mensal['A1'].alignment = center_align
    
    # Cabeçalho
    headers_mensal = [
        'RCA', 'Nome', 'Mês', 'Ano', 'Faturamento', 'Margem %', 'Margem Valor',
        'Positivação', 'Qtd Caixas', 'Qtd Itens', 'MIX'
    ]
    for col, header in enumerate(headers_mensal, 1):
        cell = ws_mensal.cell(row=3, column=col)
        cell.value = header
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = center_align
        cell.border = border
    
    # Dados
    df_mensal_sorted = df_ano.sort_values(['rca', 'ano', 'mes'])
    for idx, row in df_mensal_sorted.iterrows():
        r = idx + 4
        ws_mensal.cell(r, 1, row['rca']).border = border
        ws_mensal.cell(r, 2, row['nome_vendedor']).border = border
        ws_mensal.cell(r, 3, meses_nomes.get(int(row['mes']), row['mes'])).border = border
        ws_mensal.cell(r, 4, int(row['ano'])).border = border
        ws_mensal.cell(r, 5, row['faturamento']).number_format = '#,##0.00'
        ws_mensal.cell(r, 5).alignment = right_align
        ws_mensal.cell(r, 5).border = border
        ws_mensal.cell(r, 6, row['margem_percentual'] / 100).number_format = '0.00%'
        ws_mensal.cell(r, 6).alignment = right_align
        ws_mensal.cell(r, 6).border = border
        ws_mensal.cell(r, 7, row['margem_valor']).number_format = '#,##0.00'
        ws_mensal.cell(r, 7).alignment = right_align
        ws_mensal.cell(r, 7).border = border
        ws_mensal.cell(r, 8, int(row['positivacao'])).number_format = '#,##0'
        ws_mensal.cell(r, 8).alignment = right_align
        ws_mensal.cell(r, 8).border = border
        ws_mensal.cell(r, 9, row['qtd_caixas']).number_format = '#,##0.00'
        ws_mensal.cell(r, 9).alignment = right_align
        ws_mensal.cell(r, 9).border = border
        ws_mensal.cell(r, 10, row['qtd_itens_total']).number_format = '#,##0'
        ws_mensal.cell(r, 10).alignment = right_align
        ws_mensal.cell(r, 10).border = border
        ws_mensal.cell(r, 11, int(row['qtd_produtos_diferentes'])).number_format = '#,##0'
        ws_mensal.cell(r, 11).alignment = right_align
        ws_mensal.cell(r, 11).border = border
    
    # Ajustar larguras
    ws_mensal.column_dimensions['A'].width = 10
    ws_mensal.column_dimensions['B'].width = 30
    ws_mensal.column_dimensions['C'].width = 12
    ws_mensal.column_dimensions['D'].width = 8
    for col in range(5, 12):
        ws_mensal.column_dimensions[get_column_letter(col)].width = 16
    
    # Congelar painéis
    ws_resumo.freeze_panes = 'A4'
    ws_fat.freeze_panes = 'A4'
    ws_margem.freeze_panes = 'A4'
    ws_mix.freeze_panes = 'A4'
    ws_mensal.freeze_panes = 'A4'
    
    # Salvar em buffer
    buffer = BytesIO()
    wb.save(buffer)
    buffer.seek(0)
    
    return buffer

# =========================================================================
# INTERFACE STREAMLIT
# =========================================================================
def main():
    st.title("📊 Análise Anual de Vendedores")
    st.markdown("---")
    
    # Sidebar - Filtros
    with st.sidebar:
        st.header("⚙️ Filtros")
        
        # Ano
        hoje = datetime.now()
        ano_atual = hoje.year
        mes_atual = hoje.month
        
        # Mostrar aviso sobre o período de análise
        if mes_atual == 1:
            st.info(f"ℹ️ Analisando até 12/{ano_atual - 1} (janeiro ainda não fechou)")
        else:
            meses_nomes = {
                1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril',
                5: 'Maio', 6: 'Junho', 7: 'Julho', 8: 'Agosto',
                9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'
            }
            mes_anterior_nome = meses_nomes.get(mes_atual - 1, f"{mes_atual - 1}")
            st.info(f"ℹ️ Analisando até {mes_anterior_nome}/{ano_atual} (mês atual ainda não fechou)")
        
        ano = st.selectbox(
            "Ano",
            options=list(range(ano_atual, ano_atual - 5, -1)),
            index=0
        )
        
        # Departamento (opcional - buscar do banco)
        conn = get_db_connection()
        coddepto = None
        if conn:
            try:
                query_deptos = "SELECT CODEPTO, DESCRICAO FROM PCDEPTO ORDER BY DESCRICAO"
                df_deptos = pd.read_sql_query(query_deptos, conn)
                df_deptos.columns = [x.lower() for x in df_deptos.columns]
                
                deptos_opcoes = ["Todos"] + [f"{int(row['codepto'])} - {row['descricao']}" 
                                           for _, row in df_deptos.iterrows()]
                depto_sel = st.selectbox("Departamento", deptos_opcoes)
                
                if depto_sel != "Todos":
                    coddepto = int(depto_sel.split(" - ")[0])
            except:
                pass
        
        # RCAs (opcional)
        rcas_str = st.text_input("RCAs (separados por vírgula)", value="", help="Deixe vazio para todos")
        rcas_list = None
        if rcas_str.strip():
            try:
                rcas_list = [int(r.strip()) for r in rcas_str.split(',') if r.strip().isdigit()]
                if not rcas_list:
                    rcas_list = None
            except:
                rcas_list = None
        
        # Supervisores
        supervisores_list = None
        if conn:
            try:
                df_supervisores = buscar_supervisores(conn)
                if not df_supervisores.empty:
                    # Criar lista de opções
                    opcoes_supervisores = [
                        f"{int(row['codsupervisor'])} - {row['nome']}" 
                        for _, row in df_supervisores.iterrows()
                    ]
                    
                    # Valores padrão: supervisores 1 e 2
                    valores_padrao = []
                    for _, row in df_supervisores.iterrows():
                        cod = int(row['codsupervisor'])
                        if cod in [1, 2]:
                            valores_padrao.append(f"{cod} - {row['nome']}")
                    
                    # Se não encontrou 1 e 2, usar os dois primeiros
                    if len(valores_padrao) < 2 and len(opcoes_supervisores) >= 2:
                        valores_padrao = opcoes_supervisores[:2]
                    elif len(valores_padrao) == 0 and len(opcoes_supervisores) > 0:
                        valores_padrao = [opcoes_supervisores[0]]
                    
                    supervisores_selecionados = st.multiselect(
                        "Supervisores",
                        options=opcoes_supervisores,
                        default=valores_padrao,
                        help="Selecione os supervisores para filtrar"
                    )
                    
                    if supervisores_selecionados:
                        supervisores_list = []
                        for sel in supervisores_selecionados:
                            cod = int(sel.split(' - ')[0])
                            supervisores_list.append(cod)
            except Exception as e:
                st.warning(f"Não foi possível carregar supervisores: {e}")
        
        st.markdown("---")
        st.markdown("### 📈 Métricas Analisadas")
        st.markdown("""
        - **Faturamento**: Valor líquido de vendas
        - **Margem**: Percentual e valor de margem
        - **Positivação**: Quantidade de clientes diferentes
        - **Quantidade de Itens**: Total de itens vendidos
        - **Quantidade de Produtos**: Produtos diferentes vendidos
        - **Caixas**: Quantidade em caixas
        """)
    
    # Carregar dados
    if not conn:
        st.error("Não foi possível conectar ao banco de dados")
        return
    
    try:
        with st.spinner("Carregando dados do ano completo..."):
            df_ano = buscar_dados_ano_completo(conn, ano, coddepto, rcas_list, supervisores_list)
        
        if df_ano.empty:
            st.warning("Nenhum dado encontrado para os filtros selecionados")
            return
        
        # Calcular tendências
        with st.spinner("Calculando tendências..."):
            df_tendencias = calcular_tendencias(df_ano)
        
        if df_tendencias.empty:
            st.warning("Não foi possível calcular tendências. Verifique se há dados suficientes (pelo menos 2 meses).")
            return
    except Exception as e:
        st.error(f"Erro ao processar dados: {e}")
        import traceback
        st.code(traceback.format_exc())
        return
    
    # Ordenar por faturamento total
    df_tendencias = df_tendencias.sort_values('faturamento_total_ano', ascending=False)
    
    # =========================================================================
    # RESUMO EXECUTIVO
    # =========================================================================
    st.header("📈 Resumo Executivo")
    
    col1, col2, col3, col4, col5 = st.columns(5)
    
    total_vendedores = len(df_tendencias)
    vendedores_crescendo = len(df_tendencias[df_tendencias['tendencia'] == 'Crescente'])
    vendedores_caindo = len(df_tendencias[df_tendencias['tendencia'] == 'Decrescente'])
    faturamento_total = df_tendencias['faturamento_total_ano'].sum()
    margem_media = df_tendencias['margem_media_ano'].mean()
    
    with col1:
        st.metric("Total de Vendedores", total_vendedores)
    
    with col2:
        st.metric("Vendedores em Crescimento", vendedores_crescendo, 
                 delta=f"{vendedores_crescendo/total_vendedores*100:.1f}%")
    
    with col3:
        st.metric("Vendedores em Queda", vendedores_caindo,
                 delta=f"-{vendedores_caindo/total_vendedores*100:.1f}%")
    
    with col4:
        st.metric("Faturamento Total do Ano", 
                 f"R$ {faturamento_total:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
    
    with col5:
        st.metric("Margem Média", f"{margem_media:.2f}%")
    
    st.markdown("---")
    
    # =========================================================================
    # ANÁLISE DE TENDÊNCIAS
    # =========================================================================
    st.header("📊 Análise de Tendências")
    
    # Tabs para diferentes visualizações
    tab1, tab2, tab3, tab4, tab5 = st.tabs([
        "🎯 Visão Geral", 
        "📈 Crescimento/Declínio", 
        "💰 Faturamento e Margem",
        "👥 Positivação e Itens",
        "🔍 Detalhamento por Vendedor"
    ])
    
    with tab1:
        st.subheader("Ranking de Vendedores - Faturamento Total do Ano")
        
        # Top 20
        top_n = st.slider("Quantos vendedores mostrar?", 10, 50, 20)
        df_top = df_tendencias.head(top_n).copy()
        
        # Gráfico de barras - Faturamento Total
        fig = px.bar(
            df_top,
            x='faturamento_total_ano',
            y='nome_vendedor',
            orientation='h',
            title=f'Top {top_n} Vendedores - Faturamento Total do Ano',
            labels={'faturamento_total_ano': 'Faturamento (R$)', 'nome_vendedor': 'Vendedor'},
            color='margem_media_ano',
            color_continuous_scale='RdYlGn',
            text='faturamento_total_ano'
        )
        fig.update_layout(
            height=600, 
            yaxis={'categoryorder': 'total ascending'},
            xaxis=dict(tickformat='$,.0f')
        )
        st.plotly_chart(fig, use_container_width=True)
        
        # Tabela detalhada
        st.subheader("Tabela Detalhada")
        df_display = df_top[[
            'rca', 'nome_vendedor', 'faturamento_total_ano', 'faturamento_medio_mes',
            'var_faturamento_pct', 'margem_media_ano', 'tendencia', 'meses_com_venda'
        ]].copy()
        df_display.columns = [
            'RCA', 'Nome', 'Faturamento Total', 'Faturamento Médio/Mês',
            'Variação %', 'Margem Média %', 'Tendência', 'Meses com Venda'
        ]
        df_display['Faturamento Total'] = df_display['Faturamento Total'].apply(
            lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
        )
        df_display['Faturamento Médio/Mês'] = df_display['Faturamento Médio/Mês'].apply(
            lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
        )
        df_display['Variação %'] = df_display['Variação %'].apply(lambda x: f"{x:+.2f}%")
        df_display['Margem Média %'] = df_display['Margem Média %'].apply(lambda x: f"{x:.2f}%")
        
        st.dataframe(df_display, use_container_width=True, hide_index=True)
    
    with tab2:
        st.subheader("Análise de Crescimento vs Declínio")
        
        # Filtrar por tendência
        filtro_tendencia = st.multiselect(
            "Filtrar por Tendência",
            options=['Crescente', 'Decrescente', 'Estável'],
            default=['Crescente', 'Decrescente']
        )
        
        df_filtrado = df_tendencias[df_tendencias['tendencia'].isin(filtro_tendencia)].copy()
        
        if not df_filtrado.empty:
            # Criar coluna para size (valores sempre positivos)
            # Normalizar margem para valores entre 0 e 100, garantindo sempre positivo
            df_filtrado['size_margem'] = df_filtrado['margem_media_ano'].apply(
                lambda x: max(1, abs(x))  # Mínimo 1 para garantir visibilidade
            )
            
            # Gráfico de dispersão - Variação % vs Faturamento Total
            fig = px.scatter(
                df_filtrado,
                x='faturamento_total_ano',
                y='var_faturamento_pct',
                color='tendencia',
                size='size_margem',
                hover_data=['nome_vendedor', 'margem_media_ano'],
                title='Variação de Faturamento vs Faturamento Total',
                labels={
                    'faturamento_total_ano': 'Faturamento Total do Ano (R$)',
                    'var_faturamento_pct': 'Variação % (Primeiro vs Último Mês)',
                    'tendencia': 'Tendência'
                },
                color_discrete_map={
                    'Crescente': 'green',
                    'Decrescente': 'red',
                    'Estável': 'gray'
                }
            )
            fig.add_hline(y=0, line_dash="dash", line_color="black")
            st.plotly_chart(fig, use_container_width=True)
            
            # Top crescente e top caindo
            col1, col2 = st.columns(2)
            
            with col1:
                st.subheader("🏆 Top 10 em Crescimento")
                df_crescendo = df_filtrado[df_filtrado['tendencia'] == 'Crescente'].nlargest(10, 'var_faturamento_pct')
                if not df_crescendo.empty:
                    for idx, row in df_crescendo.iterrows():
                        st.markdown(f"""
                        **{row['nome_vendedor']}** (RCA: {row['rca']})
                        - Variação: {row['var_faturamento_pct']:+.2f}%
                        - Faturamento Total: R$ {row['faturamento_total_ano']:,.2f}
                        - Margem Média: {row['margem_media_ano']:.2f}%
                        """)
                else:
                    st.info("Nenhum vendedor em crescimento encontrado")
            
            with col2:
                st.subheader("⚠️ Top 10 em Declínio")
                df_caindo = df_filtrado[df_filtrado['tendencia'] == 'Decrescente'].nsmallest(10, 'var_faturamento_pct')
                if not df_caindo.empty:
                    for idx, row in df_caindo.iterrows():
                        st.markdown(f"""
                        **{row['nome_vendedor']}** (RCA: {row['rca']})
                        - Variação: {row['var_faturamento_pct']:+.2f}%
                        - Faturamento Total: R$ {row['faturamento_total_ano']:,.2f}
                        - Margem Média: {row['margem_media_ano']:.2f}%
                        """)
                else:
                    st.info("Nenhum vendedor em declínio encontrado")
    
    with tab3:
        st.subheader("Análise de Faturamento e Margem")
        
        # Gráfico de linha - Evolução mensal dos top vendedores
        top_vendedores = st.slider("Top N vendedores para análise mensal", 5, 20, 10)
        top_rcas = df_tendencias.head(top_vendedores)['rca'].tolist()
        df_top_mensal = df_ano[df_ano['rca'].isin(top_rcas)].copy()
        
        if not df_top_mensal.empty:
            # Criar coluna de mês/ano para ordenação
            df_top_mensal['mes_ano'] = df_top_mensal.apply(
                lambda row: f"{row['ano']}-{row['mes']:02d}", axis=1
            )
            df_top_mensal = df_top_mensal.sort_values('mes_ano')
            
            # Garantir que valores numéricos não sejam NaN
            df_top_mensal['faturamento'] = pd.to_numeric(df_top_mensal['faturamento'], errors='coerce').fillna(0)
            df_top_mensal['margem_percentual'] = pd.to_numeric(df_top_mensal['margem_percentual'], errors='coerce').fillna(0)
            
            # Gráfico de evolução de faturamento
            try:
                fig = px.line(
                    df_top_mensal,
                    x='mes_ano',
                    y='faturamento',
                    color='nome_vendedor',
                    title='Evolução Mensal de Faturamento - Top Vendedores',
                    labels={'faturamento': 'Faturamento (R$)', 'mes_ano': 'Mês/Ano'}
                )
                st.plotly_chart(fig, use_container_width=True)
            except Exception as e:
                st.error(f"Erro ao gerar gráfico de faturamento: {e}")
            
            # Gráfico de evolução de margem
            try:
                fig2 = px.line(
                    df_top_mensal,
                    x='mes_ano',
                    y='margem_percentual',
                    color='nome_vendedor',
                    title='Evolução Mensal de Margem % - Top Vendedores',
                    labels={'margem_percentual': 'Margem (%)', 'mes_ano': 'Mês/Ano'}
                )
                st.plotly_chart(fig2, use_container_width=True)
            except Exception as e:
                st.error(f"Erro ao gerar gráfico de margem: {e}")
        else:
            st.info("Não há dados mensais disponíveis para os vendedores selecionados")
        
        # Heatmap de margem por vendedor e mês
        st.subheader("Heatmap de Margem por Vendedor e Mês")
        if not df_top_mensal.empty:
            try:
                df_heatmap = df_top_mensal.pivot_table(
                    index='nome_vendedor',
                    columns='mes_ano',
                    values='margem_percentual',
                    aggfunc='mean'
                )
                # Preencher NaN com 0 para o heatmap
                df_heatmap = df_heatmap.fillna(0)
                
                if not df_heatmap.empty:
                    fig3 = px.imshow(
                        df_heatmap,
                        labels=dict(x="Mês", y="Vendedor", color="Margem %"),
                        title="Margem % por Vendedor e Mês",
                        color_continuous_scale='RdYlGn',
                        aspect="auto"
                    )
                    st.plotly_chart(fig3, use_container_width=True)
                else:
                    st.info("Não há dados suficientes para gerar o heatmap")
            except Exception as e:
                st.warning(f"Não foi possível gerar o heatmap: {e}")
        else:
            st.info("Não há dados para exibir")
    
    with tab4:
        st.subheader("Análise de Positivação e Quantidade de Itens")
        
        # Garantir que valores numéricos estão corretos
        df_tendencias_display_tab4 = df_tendencias.copy()
        df_tendencias_display_tab4['positivacao_media'] = pd.to_numeric(
            df_tendencias_display_tab4['positivacao_media'], errors='coerce'
        ).fillna(0)
        df_tendencias_display_tab4['qtd_itens_media'] = pd.to_numeric(
            df_tendencias_display_tab4['qtd_itens_media'], errors='coerce'
        ).fillna(0)
        
        # Gráfico de barras - Positivação média
        try:
            df_positivacao = df_tendencias_display_tab4.nlargest(20, 'positivacao_media')
            if not df_positivacao.empty:
                fig = px.bar(
                    df_positivacao,
                    x='nome_vendedor',
                    y='positivacao_media',
                    title='Top 20 Vendedores - Positivação Média (Clientes Diferentes)',
                    labels={'positivacao_media': 'Positivação Média', 'nome_vendedor': 'Vendedor'}
                )
                fig.update_layout(xaxis_tickangle=-45, height=500)
                st.plotly_chart(fig, use_container_width=True)
            else:
                st.info("Não há dados de positivação para exibir")
        except Exception as e:
            st.error(f"Erro ao gerar gráfico de positivação: {e}")
        
        # Gráfico de barras - Quantidade de itens
        try:
            df_itens = df_tendencias_display_tab4.nlargest(20, 'qtd_itens_media')
            if not df_itens.empty:
                fig2 = px.bar(
                    df_itens,
                    x='nome_vendedor',
                    y='qtd_itens_media',
                    title='Top 20 Vendedores - Quantidade Média de Itens Vendidos',
                    labels={'qtd_itens_media': 'Qtd Itens Média', 'nome_vendedor': 'Vendedor'}
                )
                fig2.update_layout(xaxis_tickangle=-45, height=500)
                st.plotly_chart(fig2, use_container_width=True)
            else:
                st.info("Não há dados de quantidade de itens para exibir")
        except Exception as e:
            st.error(f"Erro ao gerar gráfico de quantidade de itens: {e}")
        
        # Comparação de variações
        st.subheader("Variações: Positivação vs Quantidade de Itens")
        
        try:
            # Criar coluna para size (garantir valores sempre positivos)
            df_tendencias_display = df_tendencias.copy()
            
            # Garantir que valores numéricos estão corretos
            df_tendencias_display['var_positivacao'] = pd.to_numeric(
                df_tendencias_display['var_positivacao'], errors='coerce'
            ).fillna(0)
            df_tendencias_display['var_qtd_itens'] = pd.to_numeric(
                df_tendencias_display['var_qtd_itens'], errors='coerce'
            ).fillna(0)
            df_tendencias_display['faturamento_total_ano'] = pd.to_numeric(
                df_tendencias_display['faturamento_total_ano'], errors='coerce'
            ).fillna(0)
            
            df_tendencias_display['size_faturamento'] = df_tendencias_display['faturamento_total_ano'].apply(
                lambda x: max(1, abs(x))  # Garantir sempre positivo
            )
            
            if not df_tendencias_display.empty:
                fig3 = px.scatter(
                    df_tendencias_display,
                    x='var_positivacao',
                    y='var_qtd_itens',
                    color='tendencia',
                    size='size_faturamento',
                    hover_data=['nome_vendedor', 'faturamento_total_ano'],
                    title='Variação de Positivação vs Variação de Quantidade de Itens',
                    labels={
                        'var_positivacao': 'Variação de Positivação',
                        'var_qtd_itens': 'Variação de Qtd Itens',
                        'tendencia': 'Tendência'
                    }
                )
                st.plotly_chart(fig3, use_container_width=True)
            else:
                st.info("Não há dados para exibir o gráfico de variações")
        except Exception as e:
            st.error(f"Erro ao gerar gráfico de variações: {e}")
    
    with tab5:
        st.subheader("Detalhamento Individual por Vendedor")
        
        # Selecionar vendedor
        vendedor_selecionado = st.selectbox(
            "Selecione um vendedor",
            options=df_tendencias['nome_vendedor'].tolist()
        )
        
        if vendedor_selecionado:
            rca_selecionado = df_tendencias[df_tendencias['nome_vendedor'] == vendedor_selecionado]['rca'].iloc[0]
            df_vendedor = df_ano[df_ano['rca'] == rca_selecionado].sort_values(['ano', 'mes'])
            info_vendedor = df_tendencias[df_tendencias['rca'] == rca_selecionado].iloc[0]
            
            # Informações do vendedor
            col1, col2, col3, col4 = st.columns(4)
            with col1:
                st.metric("Faturamento Total", 
                         f"R$ {info_vendedor['faturamento_total_ano']:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
            with col2:
                st.metric("Variação %", f"{info_vendedor['var_faturamento_pct']:+.2f}%",
                         delta=f"{info_vendedor['var_faturamento']:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
            with col3:
                st.metric("Margem Média", f"{info_vendedor['margem_media_ano']:.2f}%")
            with col4:
                st.metric("Tendência", info_vendedor['tendencia'])
            
            if not df_vendedor.empty:
                # Criar coluna de mês/ano
                df_vendedor['mes_ano'] = df_vendedor.apply(
                    lambda row: f"{row['ano']}-{row['mes']:02d}", axis=1
                )
                
                # Garantir que valores numéricos estão corretos
                df_vendedor['faturamento'] = pd.to_numeric(df_vendedor['faturamento'], errors='coerce').fillna(0)
                df_vendedor['margem_percentual'] = pd.to_numeric(df_vendedor['margem_percentual'], errors='coerce').fillna(0)
                df_vendedor['positivacao'] = pd.to_numeric(df_vendedor['positivacao'], errors='coerce').fillna(0)
                df_vendedor['qtd_itens_total'] = pd.to_numeric(df_vendedor['qtd_itens_total'], errors='coerce').fillna(0)
                
                # Gráficos de evolução mensal
                try:
                    fig = make_subplots(
                        rows=2, cols=2,
                        subplot_titles=('Faturamento Mensal', 'Margem % Mensal', 
                                       'Positivação Mensal', 'Quantidade de Itens Mensal'),
                        specs=[[{"secondary_y": False}, {"secondary_y": False}],
                               [{"secondary_y": False}, {"secondary_y": False}]]
                    )
                    
                    # Faturamento
                    fig.add_trace(
                        go.Scatter(x=df_vendedor['mes_ano'], y=df_vendedor['faturamento'],
                                  mode='lines+markers', name='Faturamento', line=dict(color='blue')),
                        row=1, col=1
                    )
                    
                    # Margem
                    fig.add_trace(
                        go.Scatter(x=df_vendedor['mes_ano'], y=df_vendedor['margem_percentual'],
                                  mode='lines+markers', name='Margem %', line=dict(color='green')),
                        row=1, col=2
                    )
                    
                    # Positivação
                    fig.add_trace(
                        go.Scatter(x=df_vendedor['mes_ano'], y=df_vendedor['positivacao'],
                                  mode='lines+markers', name='Positivação', line=dict(color='orange')),
                        row=2, col=1
                    )
                    
                    # Qtd Itens
                    fig.add_trace(
                        go.Scatter(x=df_vendedor['mes_ano'], y=df_vendedor['qtd_itens_total'],
                                  mode='lines+markers', name='Qtd Itens', line=dict(color='red')),
                        row=2, col=2
                    )
                    
                    fig.update_layout(height=700, showlegend=False, title_text=f"Evolução Mensal - {vendedor_selecionado}")
                    st.plotly_chart(fig, use_container_width=True)
                except Exception as e:
                    st.error(f"Erro ao gerar gráficos de evolução: {e}")
            else:
                st.warning("Não há dados mensais disponíveis para este vendedor")
            
            # Tabela mensal detalhada
            st.subheader("📋 Dados Mensais Detalhados")
            
            # Criar DataFrame com todas as métricas
            df_display = df_vendedor[[
                'mes', 'ano', 'faturamento', 'margem_percentual', 'margem_valor',
                'positivacao', 'qtd_caixas', 'qtd_itens_total', 'qtd_produtos_diferentes'
            ]].copy()
            
            # Adicionar colunas de comparação (último vs melhor vs pior)
            melhor_fat = info_vendedor['faturamento_melhor_mes']
            pior_fat = info_vendedor['faturamento_pior_mes']
            melhor_margem = info_vendedor['margem_melhor_mes']
            pior_margem = info_vendedor['margem_pior_mes']
            melhor_mix = info_vendedor['mix_melhor_mes']
            pior_mix = info_vendedor['mix_pior_mes']
            
            # Função para determinar status
            def get_status_fat(valor):
                if valor >= melhor_fat:
                    return "🟢 Melhor"
                elif valor <= pior_fat:
                    return "🔴 Pior"
                else:
                    return "🟡 Médio"
            
            def get_status_margem(valor):
                if valor >= melhor_margem:
                    return "🟢 Melhor"
                elif valor <= pior_margem:
                    return "🔴 Pior"
                else:
                    return "🟡 Médio"
            
            def get_status_mix(valor):
                if valor >= melhor_mix:
                    return "🟢 Melhor"
                elif valor <= pior_mix:
                    return "🔴 Pior"
                else:
                    return "🟡 Médio"
            
            # Adicionar colunas de status
            df_display['Status Fat'] = df_display['faturamento'].apply(get_status_fat)
            df_display['Status Margem'] = df_display['margem_percentual'].apply(get_status_margem)
            df_display['Status MIX'] = df_display['qtd_produtos_diferentes'].apply(get_status_mix)
            
            # Renomear colunas
            df_display.columns = [
                'Mês', 'Ano', 'Faturamento', 'Margem %', 'Margem Valor',
                'Positivação', 'Qtd Caixas', 'Qtd Itens', 'MIX', 'Status Fat', 'Status Margem', 'Status MIX'
            ]
            
            # Formatação
            df_display['Faturamento'] = df_display['Faturamento'].apply(
                lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
            )
            df_display['Margem Valor'] = df_display['Margem Valor'].apply(
                lambda x: f"R$ {x:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.')
            )
            df_display['Margem %'] = df_display['Margem %'].apply(lambda x: f"{x:.2f}%")
            df_display['MIX'] = df_display['MIX'].apply(lambda x: f"{int(x)}")
            
            # Reordenar colunas
            df_display = df_display[[
                'Mês', 'Ano', 'Faturamento', 'Status Fat', 
                'Margem %', 'Margem Valor', 'Status Margem',
                'MIX', 'Status MIX',
                'Positivação', 'Qtd Caixas', 'Qtd Itens'
            ]]
            
            st.dataframe(df_display, use_container_width=True, hide_index=True)
            
            # Análise comparativa: Melhor, Pior e Último Mês
            st.subheader("📊 Análise Comparativa: Melhor, Pior e Último Mês")
            
            meses_nomes = {
                1: 'Janeiro', 2: 'Fevereiro', 3: 'Março', 4: 'Abril',
                5: 'Maio', 6: 'Junho', 7: 'Julho', 8: 'Agosto',
                9: 'Setembro', 10: 'Outubro', 11: 'Novembro', 12: 'Dezembro'
            }
            
            # Buscar dados dos meses específicos
            try:
                melhor_mes_data = df_vendedor[df_vendedor['faturamento'] == info_vendedor['faturamento_melhor_mes']].iloc[0]
            except:
                melhor_mes_data = df_vendedor.iloc[df_vendedor['faturamento'].idxmax()]
            
            try:
                pior_mes_data = df_vendedor[df_vendedor['faturamento'] == info_vendedor['faturamento_pior_mes']].iloc[0]
            except:
                pior_mes_data = df_vendedor.iloc[df_vendedor['faturamento'].idxmin()]
            
            ultimo_mes_data = df_vendedor.iloc[-1]
            
            mes_melhor_nome = meses_nomes.get(int(info_vendedor['mes_melhor']), f"Mês {int(info_vendedor['mes_melhor'])}")
            mes_pior_nome = meses_nomes.get(int(info_vendedor['mes_pior']), f"Mês {int(info_vendedor['mes_pior'])}")
            
            col1, col2, col3 = st.columns(3)
            
            with col1:
                st.markdown(f"### 🏆 Melhor Mês do Ano ({mes_melhor_nome})")
                st.metric("Faturamento", f"R$ {info_vendedor['faturamento_melhor_mes']:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
                st.metric("Positivação", int(info_vendedor['positivacao_melhor_mes']))
                st.metric("Qtd Itens", f"{info_vendedor['qtd_itens_melhor_mes']:,.0f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
                st.metric("MIX (Produtos Diferentes)", int(info_vendedor['mix_melhor_mes']))
                st.metric("Margem %", f"{info_vendedor['margem_melhor_mes']:.2f}%")
            
            with col2:
                st.markdown(f"### 📉 Pior Mês do Ano ({mes_pior_nome})")
                st.metric("Faturamento", f"R$ {info_vendedor['faturamento_pior_mes']:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
                st.metric("Positivação", int(info_vendedor['positivacao_pior_mes']))
                st.metric("Qtd Itens", f"{info_vendedor['qtd_itens_pior_mes']:,.0f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
                st.metric("MIX (Produtos Diferentes)", int(info_vendedor['mix_pior_mes']))
                st.metric("Margem %", f"{info_vendedor['margem_pior_mes']:.2f}%")
            
            with col3:
                st.markdown("### 📅 Último Mês (Atual)")
                st.metric("Faturamento", f"R$ {info_vendedor['faturamento_ultimo_mes']:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
                st.metric("Positivação", int(info_vendedor['positivacao_ultimo_mes']))
                st.metric("Qtd Itens", f"{info_vendedor['qtd_itens_ultimo_mes']:,.0f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
                st.metric("MIX (Produtos Diferentes)", int(info_vendedor['mix_ultimo_mes']))
                st.metric("Margem %", f"{info_vendedor['margem_ultimo_mes']:.2f}%")
            
            # Comparações
            st.markdown("---")
            st.subheader("📈 Comparações: Último Mês vs Melhor e Pior Mês")
            
            col1, col2 = st.columns(2)
            
            with col1:
                st.markdown("### Último Mês vs Melhor Mês")
                
                # Faturamento
                if info_vendedor['var_vs_melhor_fat_pct'] < 0:
                    st.error(f"❌ Faturamento: {abs(info_vendedor['var_vs_melhor_fat_pct']):.2f}% abaixo do melhor mês")
                    st.metric("Diferença", f"R$ {abs(info_vendedor['var_vs_melhor_fat']):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
                else:
                    st.success(f"✅ Faturamento: {info_vendedor['var_vs_melhor_fat_pct']:.2f}% acima do melhor mês")
                    st.metric("Diferença", f"R$ {info_vendedor['var_vs_melhor_fat']:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
                
                # Positivação
                if info_vendedor['var_vs_melhor_posit'] < 0:
                    st.warning(f"⚠️ Positivação: {abs(int(info_vendedor['var_vs_melhor_posit']))} clientes a menos")
                elif info_vendedor['var_vs_melhor_posit'] > 0:
                    st.success(f"✅ Positivação: {int(info_vendedor['var_vs_melhor_posit'])} clientes a mais")
                else:
                    st.info("➡️ Positivação: igual ao melhor mês")
                
                # Qtd Itens
                if info_vendedor['var_vs_melhor_itens'] < 0:
                    st.warning(f"⚠️ Qtd Itens: {abs(info_vendedor['var_vs_melhor_itens']):,.0f} itens a menos".replace(',', 'X').replace('.', ',').replace('X', '.'))
                elif info_vendedor['var_vs_melhor_itens'] > 0:
                    st.success(f"✅ Qtd Itens: {info_vendedor['var_vs_melhor_itens']:,.0f} itens a mais".replace(',', 'X').replace('.', ',').replace('X', '.'))
                else:
                    st.info("➡️ Qtd Itens: igual ao melhor mês")
                
                # Margem
                if info_vendedor['var_vs_melhor_margem'] < 0:
                    st.warning(f"⚠️ Margem: {abs(info_vendedor['var_vs_melhor_margem']):.2f}% abaixo")
                elif info_vendedor['var_vs_melhor_margem'] > 0:
                    st.success(f"✅ Margem: {info_vendedor['var_vs_melhor_margem']:.2f}% acima")
                else:
                    st.info("➡️ Margem: igual ao melhor mês")
                
                # MIX
                if info_vendedor['var_vs_melhor_mix'] < 0:
                    st.warning(f"⚠️ MIX: {abs(int(info_vendedor['var_vs_melhor_mix']))} produtos a menos")
                elif info_vendedor['var_vs_melhor_mix'] > 0:
                    st.success(f"✅ MIX: {int(info_vendedor['var_vs_melhor_mix'])} produtos a mais")
                else:
                    st.info("➡️ MIX: igual ao melhor mês")
            
            with col2:
                st.markdown("### Último Mês vs Pior Mês")
                
                # Faturamento
                if info_vendedor['var_vs_pior_fat_pct'] > 0:
                    st.success(f"✅ Faturamento: {info_vendedor['var_vs_pior_fat_pct']:.2f}% acima do pior mês")
                    st.metric("Diferença", f"R$ {info_vendedor['var_vs_pior_fat']:,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
                else:
                    st.error(f"❌ Faturamento: {abs(info_vendedor['var_vs_pior_fat_pct']):.2f}% abaixo do pior mês")
                    st.metric("Diferença", f"R$ {abs(info_vendedor['var_vs_pior_fat']):,.2f}".replace(',', 'X').replace('.', ',').replace('X', '.'))
                
                # Positivação
                if info_vendedor['var_vs_pior_posit'] > 0:
                    st.success(f"✅ Positivação: {int(info_vendedor['var_vs_pior_posit'])} clientes a mais")
                elif info_vendedor['var_vs_pior_posit'] < 0:
                    st.warning(f"⚠️ Positivação: {abs(int(info_vendedor['var_vs_pior_posit']))} clientes a menos")
                else:
                    st.info("➡️ Positivação: igual ao pior mês")
                
                # Qtd Itens
                if info_vendedor['var_vs_pior_itens'] > 0:
                    st.success(f"✅ Qtd Itens: {info_vendedor['var_vs_pior_itens']:,.0f} itens a mais".replace(',', 'X').replace('.', ',').replace('X', '.'))
                elif info_vendedor['var_vs_pior_itens'] < 0:
                    st.warning(f"⚠️ Qtd Itens: {abs(info_vendedor['var_vs_pior_itens']):,.0f} itens a menos".replace(',', 'X').replace('.', ',').replace('X', '.'))
                else:
                    st.info("➡️ Qtd Itens: igual ao pior mês")
                
                # Margem
                if info_vendedor['var_vs_pior_margem'] > 0:
                    st.success(f"✅ Margem: {info_vendedor['var_vs_pior_margem']:.2f}% acima")
                elif info_vendedor['var_vs_pior_margem'] < 0:
                    st.warning(f"⚠️ Margem: {abs(info_vendedor['var_vs_pior_margem']):.2f}% abaixo")
                else:
                    st.info("➡️ Margem: igual ao pior mês")
                
                # MIX
                if info_vendedor['var_vs_pior_mix'] > 0:
                    st.success(f"✅ MIX: {int(info_vendedor['var_vs_pior_mix'])} produtos a mais")
                elif info_vendedor['var_vs_pior_mix'] < 0:
                    st.warning(f"⚠️ MIX: {abs(int(info_vendedor['var_vs_pior_mix']))} produtos a menos")
                else:
                    st.info("➡️ MIX: igual ao pior mês")
            
            # Resumo final
            st.markdown("---")
            st.subheader("📋 Resumo da Performance")
            
            # Determinar status geral
            if info_vendedor['var_vs_melhor_fat_pct'] >= 0:
                st.success("🎉 **EXCELENTE**: O último mês foi igual ou melhor que o melhor mês do ano!")
            elif info_vendedor['var_vs_pior_fat_pct'] > 0:
                st.info("📊 **BOM**: O último mês está melhor que o pior mês, mas abaixo do melhor mês.")
            else:
                st.error("⚠️ **ATENÇÃO**: O último mês está abaixo até mesmo do pior mês do ano!")
    
    # =========================================================================
    # EXPORTAÇÃO
    # =========================================================================
    st.markdown("---")
    st.header("💾 Exportar Dados")
    
    col1, col2, col3 = st.columns(3)
    
    with col1:
        # Excel completo com análises
        excel_buffer = gerar_excel_analise_completa(df_tendencias, df_ano, ano)
        st.download_button(
            label="📊 Baixar Análise Completa (Excel)",
            data=excel_buffer,
            file_name=f"analise_completa_vendedores_{ano}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    
    with col2:
        # CSV de tendências
        csv_tendencias = df_tendencias.to_csv(index=False, encoding='utf-8-sig')
        st.download_button(
            label="📥 Baixar Análise de Tendências (CSV)",
            data=csv_tendencias,
            file_name=f"analise_tendencias_vendedores_{ano}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv"
        )
    
    with col3:
        # CSV de dados mensais
        csv_mensal = df_ano.to_csv(index=False, encoding='utf-8-sig')
        st.download_button(
            label="📥 Baixar Dados Mensais (CSV)",
            data=csv_mensal,
            file_name=f"dados_mensais_vendedores_{ano}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv",
            mime="text/csv"
        )

if __name__ == "__main__":
    main()

